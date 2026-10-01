import { READ_CELLS_LIMITS } from '../../utils/storage';
import {
  ToolError,
  cellAt,
  getSheet,
  loadUserSheets,
  localAddress,
  optionalBoolean,
  optionalString,
  requireString,
  resolveRange,
  type Args,
  type ToolOptions,
} from './common';
import { recordUndo } from './undo';

const MAX_READ_COLUMNS = 50;
const MAX_HEADER_COLUMNS = 30;
const SAMPLE_ROWS = 3;
const MAX_TABLES_IN_CONTEXT = 30;
const MAX_NAMES_IN_CONTEXT = 50;
const MAX_SEARCH_RESULTS = 100;
/** Above this many matches in a sheet only addresses are returned, not values. */
const MAX_SEARCH_VALUES = 5000;

const loadSheets = async (context: Excel.RequestContext) => {
  const active = context.workbook.worksheets.getActiveWorksheet();
  active.load('name');
  const sheets = await loadUserSheets(context, 'items/name, items/visibility');

  const used = sheets.map(sheet => {
    const range = sheet.getUsedRangeOrNullObject(true);
    range.load('address, rowCount, columnCount');
    return { sheet, range };
  });
  await context.sync();
  return { activeSheet: active.name, used };
};

const getSelectionAddress = async (context: Excel.RequestContext): Promise<string | null> => {
  try {
    const range = context.workbook.getSelectedRange();
    range.load('address');
    await context.sync();
    return range.address;
  } catch {
    // The selection is not a range (e.g. a chart or shape).
    return null;
  }
};

/** Sheet-qualified address of the selected range, or null when the selection is not a range. */
export const getSelectedRangeAddress = (): Promise<string | null> => Excel.run(getSelectionAddress);

export interface WorkbookOverview {
  activeSheet: string;
  selection: string | null;
  sheets: { name: string; usedRange: string | null; hidden: boolean }[];
}

/** Lightweight summary used in the system prompt. */
export const getWorkbookOverview = (): Promise<WorkbookOverview> =>
  Excel.run(async context => {
    const { activeSheet, used } = await loadSheets(context);
    return {
      activeSheet,
      selection: await getSelectionAddress(context),
      sheets: used.map(({ sheet, range }) => ({
        name: sheet.name,
        usedRange: range.isNullObject ? null : localAddress(range.address),
        hidden: sheet.visibility !== 'Visible',
      })),
    };
  });

export const getWorkbookContext = async (context: Excel.RequestContext) => {
  const { activeSheet, used } = await loadSheets(context);

  const tables = context.workbook.tables;
  tables.load('items/name');
  const names = context.workbook.names;
  names.load('items/name, items/formula, items/visible');
  const objects = used.map(({ sheet }) => {
    const charts = sheet.charts;
    charts.load('items/name, items/chartType');
    const pivots = sheet.pivotTables;
    pivots.load('items/name');
    return { charts, pivots };
  });
  await context.sync();

  const tableInfo = tables.items.slice(0, MAX_TABLES_IN_CONTEXT).map(table => {
    const sheet = table.worksheet;
    sheet.load('name');
    const range = table.getRange();
    range.load('address');
    const header = table.getHeaderRowRange();
    header.load('values');
    return { table, sheet, range, header };
  });

  // Header row of every sheet, plus a few sample rows of the active sheet.
  const previews = used.map(({ sheet, range }) => {
    if (range.isNullObject) return null;
    const rows = Math.min(range.rowCount, sheet.name === activeSheet ? 1 + SAMPLE_ROWS : 1);
    const cols = Math.min(range.columnCount, MAX_HEADER_COLUMNS);
    const preview = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    preview.load('values');
    return preview;
  });
  await context.sync();
  const selection = await getSelectionAddress(context);

  const visibleNames = names.items.filter(n => n.visible);
  return {
    activeSheet,
    selection,
    sheets: used.map(({ sheet, range }, i) => {
      const values = previews[i]?.values ?? [];
      const { charts, pivots } = objects[i];
      return {
        name: sheet.name,
        ...(sheet.visibility !== 'Visible' ? { visibility: sheet.visibility } : {}),
        ...(range.isNullObject ? { empty: true } : {
          usedRange: localAddress(range.address),
          rows: range.rowCount,
          columns: range.columnCount,
          firstRow: values[0],
          ...(sheet.name === activeSheet ? { sampleRows: values.slice(1) } : {}),
        }),
        ...(charts.items.length ? { charts: charts.items.map(c => ({ name: c.name, type: c.chartType })) } : {}),
        ...(pivots.items.length ? { pivotTables: pivots.items.map(p => p.name) } : {}),
      };
    }),
    tables: tableInfo.map(({ table, sheet, range, header }) => ({
      name: table.name,
      sheet: sheet.name,
      range: localAddress(range.address),
      headers: header.values[0],
    })),
    ...(tables.items.length > MAX_TABLES_IN_CONTEXT ? { tablesNote: `Only the first ${MAX_TABLES_IN_CONTEXT} of ${tables.items.length} tables are listed.` } : {}),
    ...(visibleNames.length ? { namedRanges: visibleNames.slice(0, MAX_NAMES_IN_CONTEXT).map(n => ({ name: n.name, refersTo: n.formula })) } : {}),
  };
};

export const readRange = async (context: Excel.RequestContext, args: Args, options: ToolOptions) => {
  const maxCells = options.maxReadCells ?? READ_CELLS_LIMITS.default;
  const { sheet, sheetName, range } = await resolveRange(context, args, 'range_address');
  const includeFormulas = optionalBoolean(args, 'include_formulas') ?? false;
  const requested = localAddress(range.address);

  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) return { sheet: sheetName, address: requested, values: [], note: 'The sheet is empty.' };

  // Only read the part that contains data, so that "A:A" doesn't load a million empty cells.
  const data = range.getIntersectionOrNullObject(used);
  data.load('address, rowCount, columnCount');
  await context.sync();
  if (data.isNullObject) return { sheet: sheetName, address: requested, values: [], note: 'The range contains no data.' };

  const cols = Math.min(data.columnCount, MAX_READ_COLUMNS);
  const rows = Math.min(data.rowCount, Math.max(1, Math.floor(maxCells / cols)));
  const part = data.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  part.load(includeFormulas ? 'address, values, formulas' : 'address, values');
  await context.sync();

  const truncated = rows < data.rowCount || cols < data.columnCount;
  return {
    sheet: sheetName,
    address: localAddress(part.address),
    values: part.values,
    ...(includeFormulas ? { formulas: part.formulas } : {}),
    ...(truncated ? {
      truncated: true,
      dataRange: localAddress(data.address),
      note: `Only the first ${rows} rows × ${cols} columns of ${data.rowCount} × ${data.columnCount} were returned. Read the rest with smaller ranges.`,
    } : {}),
  };
};

export const searchWorkbook = async (context: Excel.RequestContext, args: Args) => {
  const query = requireString(args, 'query');
  const sheetName = optionalString(args, 'sheet_name');
  const criteria = {
    completeMatch: optionalBoolean(args, 'whole_cell') ?? false,
    matchCase: optionalBoolean(args, 'match_case') ?? false,
  };

  let sheets: Excel.Worksheet[];
  if (sheetName) {
    sheets = [await getSheet(context, sheetName)];
  } else {
    sheets = await loadUserSheets(context);
  }

  const found = sheets.map(sheet => {
    const areas = sheet.findAllOrNullObject(query, criteria);
    areas.load('cellCount');
    return { sheet, areas };
  });
  await context.sync();

  const hits = found.filter(f => !f.areas.isNullObject);
  for (const hit of hits) {
    hit.areas.areas.load(hit.areas.cellCount <= MAX_SEARCH_VALUES ? 'items/address, items/values' : 'items/address, items/rowCount, items/columnCount');
  }
  await context.sync();

  const total = hits.reduce((sum, hit) => sum + hit.areas.cellCount, 0);
  const results: { sheet: string; cell: string; value?: unknown }[] = [];
  collect: for (const hit of hits) {
    for (const area of hit.areas.areas.items) {
      const withValues = hit.areas.cellCount <= MAX_SEARCH_VALUES;
      const rowCount = withValues ? area.values.length : area.rowCount;
      const colCount = withValues ? area.values[0].length : area.columnCount;
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          if (results.length >= MAX_SEARCH_RESULTS) break collect;
          results.push({ sheet: hit.sheet.name, cell: cellAt(area.address, r, c), ...(withValues ? { value: area.values[r][c] } : {}) });
        }
      }
    }
  }

  return {
    query,
    totalMatches: total,
    results,
    ...(total > results.length ? { note: `Only the first ${results.length} of ${total} matches are listed.` } : {}),
  };
};

export const activateWorksheet = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'name'));
  sheet.activate();
  await context.sync();
  return { activated: sheet.name };
};

const validateSheetName = (name: string) => {
  if (name.length > 31) throw new ToolError('Sheet names can have at most 31 characters.');
  if (/[\\/?*[\]:]/.test(name)) throw new ToolError('Sheet names cannot contain \\ / ? * [ ] or :');
  if (name.startsWith("'") || name.endsWith("'")) throw new ToolError('Sheet names cannot start or end with an apostrophe.');
  if (name.toLowerCase() === 'history') throw new ToolError('"History" is a reserved sheet name.');
};

const findSheetNamed = async (context: Excel.RequestContext, name: string) => {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  await context.sync();
  return sheets.items.find(s => s.name.toLowerCase() === name.toLowerCase());
};

export const createWorksheet = async (context: Excel.RequestContext, args: Args) => {
  const name = requireString(args, 'name');
  validateSheetName(name);

  const active = context.workbook.worksheets.getActiveWorksheet();
  active.load('name');
  const existing = await findSheetNamed(context, name);
  if (existing) throw new ToolError(`A worksheet named "${existing.name}" already exists. Use it or choose another name.`);

  const sheet = context.workbook.worksheets.add(name);
  sheet.activate();
  await context.sync();
  recordUndo({ kind: 'createdSheet', sheetName: name, previousActiveSheet: active.name });
  return { created: name, activated: true };
};

export const renameWorksheet = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'name'));
  const newName = requireString(args, 'new_name');
  validateSheetName(newName);

  const existing = await findSheetNamed(context, newName);
  if (existing && existing.name !== sheet.name) throw new ToolError(`A worksheet named "${existing.name}" already exists.`);

  const oldName = sheet.name;
  sheet.name = newName;
  await context.sync();
  recordUndo({ kind: 'renamedSheet', from: oldName, to: newName });
  return { renamed: oldName, to: newName };
};

export const createNamedRange = async (context: Excel.RequestContext, args: Args) => {
  const name = requireString(args, 'name');
  const comment = optionalString(args, 'comment');
  const { range } = await resolveRange(context, args, 'range_address');

  const existing = context.workbook.names.getItemOrNullObject(name);
  await context.sync();
  if (!existing.isNullObject) throw new ToolError(`The name "${name}" already exists in this workbook.`);

  context.workbook.names.add(name, range, comment);
  await context.sync();
  recordUndo({ kind: 'name', name });
  return { name, refersTo: range.address };
};
