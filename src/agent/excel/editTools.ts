import {
  ToolError,
  getSheet,
  loadUserSheets,
  localAddress,
  optionalBoolean,
  optionalEnum,
  optionalString,
  requireApi,
  requireString,
  resolveRange,
  stringList,
  topLeft,
  columnIndex,
  type Args,
} from './common';
import { snapshotRange } from './undo';

const PASTE_TYPES = { all: 'All', values: 'Values', formats: 'Formats', formulas: 'Formulas' } as const;

export const copyRange = async (context: Excel.RequestContext, args: Args) => {
  const move = optionalBoolean(args, 'move') ?? false;
  const paste = optionalEnum(args, 'paste', ['all', 'values', 'formats', 'formulas'] as const) ?? 'all';
  const transpose = optionalBoolean(args, 'transpose') ?? false;
  if (move) requireApi('1.11', 'Moving ranges');
  if (move && (paste !== 'all' || transpose)) throw new ToolError('"move" cannot be combined with "paste" or "transpose".');

  const source = await resolveRange(context, args, 'source_range', 'source_sheet');
  const destinationSheet = await getSheet(context, optionalString(args, 'destination_sheet') ?? source.sheetName);
  destinationSheet.load('name');
  const anchor = destinationSheet.getRange(localAddress(requireString(args, 'destination_cell')));
  const [rows, cols] = transpose ? [source.range.columnCount, source.range.rowCount] : [source.range.rowCount, source.range.columnCount];
  const destination = anchor.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  destination.load('address');
  await context.sync();

  const destinationSnapshot = await snapshotRange(context, destinationSheet.name, destination);
  const sourceSnapshot = move ? await snapshotRange(context, source.sheetName, source.range) : null;

  if (move) source.range.moveTo(destination);
  else destination.copyFrom(source.range, PASTE_TYPES[paste], false, transpose);
  await context.sync();

  return {
    sheet: destinationSheet.name,
    address: localAddress(destination.address),
    from: source.range.address,
    operation: move ? 'moved' : `copied (${paste}${transpose ? ', transposed' : ''})`,
    undoAvailable: destinationSnapshot !== null && (!move || sourceSnapshot !== null),
    ...(move ? { note: 'Formulas that referred to the moved cells now refer to the new location; undo does not change them back.' } : {}),
  };
};

const FILL_TYPES = ['FillDefault', 'FillCopy', 'FillSeries', 'FillFormats', 'FillValues', 'FillDays', 'FillWeekdays', 'FillMonths', 'FillYears', 'LinearTrend', 'GrowthTrend', 'FlashFill'] as const;

export const fillRange = async (context: Excel.RequestContext, args: Args) => {
  const type = optionalEnum(args, 'fill_type', FILL_TYPES) ?? 'FillDefault';
  const { sheet, sheetName, range: source } = await resolveRange(context, args, 'source_range');
  const destination = sheet.getRange(localAddress(requireString(args, 'destination_range')));
  destination.load('address, cellCount');
  try {
    await context.sync();
  } catch (e) {
    throw new ToolError('"destination_range" is not a valid range address.', { cause: e });
  }

  const snapshot = await snapshotRange(context, sheetName, destination);
  source.autoFill(destination, type);
  await context.sync();
  return { sheet: sheetName, address: localAddress(destination.address), filledFrom: localAddress(source.address), fillType: type, undoAvailable: snapshot !== null };
};

export const findReplace = async (context: Excel.RequestContext, args: Args) => {
  const find = requireString(args, 'find');
  const replacement = typeof args.replace === 'string' ? args.replace : String(args.replace ?? '');
  const criteria = {
    completeMatch: optionalBoolean(args, 'whole_cell') ?? false,
    matchCase: optionalBoolean(args, 'match_case') ?? false,
  };
  const sheetName = optionalString(args, 'sheet_name');
  const rangeAddress = optionalString(args, 'range_address');

  let targets: { sheet: string; range: Excel.Range }[];
  if (rangeAddress) {
    const { sheetName: name, range } = await resolveRange(context, args, 'range_address');
    targets = [{ sheet: name, range }];
  } else {
    let sheets: Excel.Worksheet[];
    if (sheetName) {
      sheets = [await getSheet(context, sheetName)];
    } else {
      sheets = await loadUserSheets(context);
    }
    targets = [];
    for (const sheet of sheets) {
      const used = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (!used.isNullObject) targets.push({ sheet: sheet.name, range: used });
    }
  }

  // Only snapshot ranges that actually contain matches.
  const withMatches: typeof targets = [];
  for (const target of targets) {
    const found = target.range.findOrNullObject(find, criteria);
    await context.sync();
    if (!found.isNullObject) withMatches.push(target);
  }

  let undoable = true;
  const counts: { sheet: string; result: OfficeExtension.ClientResult<number> }[] = [];
  for (const target of withMatches) {
    if (!(await snapshotRange(context, target.sheet, target.range))) undoable = false;
    counts.push({ sheet: target.sheet, result: target.range.replaceAll(find, replacement, criteria) });
  }
  await context.sync();

  const replaced = counts.reduce((sum, c) => sum + c.result.value, 0);
  return {
    find,
    replace: replacement,
    replaced,
    bySheet: counts.filter(c => c.result.value > 0).map(c => ({ sheet: c.sheet, replaced: c.result.value })),
    undoAvailable: undoable,
    ...(withMatches.length === 1 ? { sheet: withMatches[0].sheet, address: localAddress(withMatches[0].range.address) } : {}),
  };
};

export const removeDuplicates = async (context: Excel.RequestContext, args: Args) => {
  const hasHeaders = optionalBoolean(args, 'has_headers') ?? true;
  const resolved = await resolveRange(context, args, 'range_address');
  const range = resolved.range.cellCount === 1 ? resolved.range.getSurroundingRegion() : resolved.range;
  const header = range.getRow(0);
  header.load('values');
  range.load('address, columnCount');
  await context.sync();

  // Columns to compare: header names or column letters; all columns by default.
  const headers = header.values[0].map(v => String(v).trim().toLowerCase());
  const start = topLeft(range.address).col;
  const requested = stringList(args, 'columns');
  const columns = requested.length === 0
    ? Array.from({ length: range.columnCount }, (_, i) => i)
    : requested.map(name => {
      const byName = hasHeaders ? headers.indexOf(name.toLowerCase()) : -1;
      if (byName !== -1) return byName;
      if (/^[A-Z]{1,3}$/i.test(name)) {
        const index = columnIndex(name) - start;
        if (index >= 0 && index < range.columnCount) return index;
      }
      throw new ToolError(`Column "${name}" is not in the range.${hasHeaders ? ` Headers: ${header.values[0].map(h => `"${h}"`).join(', ')}.` : ''}`);
    });

  const snapshot = await snapshotRange(context, resolved.sheetName, range);
  const result = range.removeDuplicates(columns, hasHeaders);
  result.load('removed, uniqueRemaining');
  await context.sync();
  return {
    sheet: resolved.sheetName,
    address: localAddress(range.address),
    removed: result.removed,
    remaining: result.uniqueRemaining,
    undoAvailable: snapshot !== null,
  };
};
