import { READ_CELLS_LIMITS } from '../../utils/storage';
import { getAttachment, type AttachmentCell } from '../attachments';
import {
  DEFAULT_TABLE_STYLE,
  ToolError,
  columnLetters,
  getTable,
  localAddress,
  optionalBoolean,
  optionalNumber,
  optionalString,
  requireString,
  resolveRange,
  topLeft,
  type Args,
  type ToolOptions,
} from './common';
import { recordUndo, snapshotRange } from './undo';

const PROFILE_CHUNK_CELLS = 100000;
const MAX_PROFILE_CELLS = 2000000;
const MAX_DISTINCT = 10000;
const TOP_VALUES = 5;
const IMPORT_CHUNK_CELLS = 20000;
const MAX_TEXT_READ = 20000;

interface ColumnAccumulator {
  numbers: number[];
  counts: Map<string, number>;
  distinctOverflow: boolean;
  text: number;
  booleans: number;
  errors: number;
  blanks: number;
  minLength: number;
  maxLength: number;
}

const newAccumulator = (): ColumnAccumulator => ({
  numbers: [], counts: new Map(), distinctOverflow: false, text: 0, booleans: 0, errors: 0, blanks: 0, minLength: Infinity, maxLength: 0,
});

/** Whether a number format displays dates (ignoring quoted text and [color]/[locale] blocks). */
const isDateFormat = (format: unknown) =>
  typeof format === 'string' && /[dmy]/i.test(format.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')) && format !== 'General';

const serialToIso = (serial: number) => new Date(Math.round((serial - 25569) * 86400000)).toISOString().slice(0, 10);

const round = (n: number) => Math.round(n * 10000) / 10000;

const accumulate = (acc: ColumnAccumulator, value: unknown, type: string) => {
  if (type === 'Empty' || value === '') {
    acc.blanks++;
    return;
  }
  if (type === 'Error') {
    acc.errors++;
    return;
  }
  if (typeof value === 'number') acc.numbers.push(value);
  else if (typeof value === 'boolean') acc.booleans++;
  else {
    acc.text++;
    const length = String(value).length;
    acc.minLength = Math.min(acc.minLength, length);
    acc.maxLength = Math.max(acc.maxLength, length);
  }
  const key = String(value);
  if (acc.counts.has(key)) acc.counts.set(key, acc.counts.get(key)! + 1);
  else if (acc.counts.size < MAX_DISTINCT) acc.counts.set(key, 1);
  else acc.distinctOverflow = true;
};

const summarize = (acc: ColumnAccumulator, name: string, letter: string, isDate: boolean) => {
  const filled = acc.numbers.length + acc.text + acc.booleans;
  const kinds = [acc.numbers.length && 'number', acc.text && 'text', acc.booleans && 'boolean'].filter(Boolean);
  const type = kinds.length === 0 ? 'empty' : kinds.length > 1 ? 'mixed' : kinds[0] === 'number' && isDate ? 'date' : kinds[0];
  const base = {
    column: name,
    letter,
    type,
    filled,
    blanks: acc.blanks,
    ...(acc.errors ? { errors: acc.errors } : {}),
    distinct: acc.distinctOverflow ? `${MAX_DISTINCT}+` : acc.counts.size,
  };

  if (acc.numbers.length) {
    const sorted = [...acc.numbers].sort((a, b) => a - b);
    const sum = sorted.reduce((s, n) => s + n, 0);
    const mean = sum / sorted.length;
    const middle = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
    const stdDev = Math.sqrt(sorted.reduce((s, n) => s + (n - mean) ** 2, 0) / sorted.length);
    Object.assign(base, isDate
      ? { min: serialToIso(sorted[0]), max: serialToIso(sorted[sorted.length - 1]) }
      : { min: sorted[0], max: sorted[sorted.length - 1], sum: round(sum), mean: round(mean), median: round(median), stdDev: round(stdDev) });
  }
  if (acc.text || acc.booleans || type === 'mixed') {
    const top = [...acc.counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, TOP_VALUES).map(([value, count]) => ({ value, count }));
    Object.assign(base, { topValues: top });
    if (acc.text) Object.assign(base, { minLength: acc.minLength, maxLength: acc.maxLength });
  }
  return base;
};

export const profileData = async (context: Excel.RequestContext, args: Args) => {
  const tableName = optionalString(args, 'table_name');
  let sheetName: string;
  let data: Excel.Range;
  let hasHeaders = optionalBoolean(args, 'has_headers') ?? true;

  if (tableName) {
    const table = await getTable(context, tableName);
    const sheet = table.worksheet;
    sheet.load('name');
    data = table.getRange();
    await context.sync();
    sheetName = sheet.name;
    hasHeaders = true;
  } else {
    const resolved = await resolveRange(context, args, 'range_address');
    sheetName = resolved.sheetName;
    const target = resolved.range.cellCount === 1 ? resolved.range.getSurroundingRegion() : resolved.range;
    const used = resolved.sheet.getUsedRangeOrNullObject(true);
    await context.sync();
    if (used.isNullObject) return { sheet: sheetName, note: 'The sheet is empty.' };
    data = target.getIntersectionOrNullObject(used);
    await context.sync();
    if (data.isNullObject) return { sheet: sheetName, note: 'The range contains no data.' };
  }
  data.load('address, rowCount, columnCount');
  await context.sync();

  const cols = data.columnCount;
  const headerRows = hasHeaders ? 1 : 0;
  const dataRows = data.rowCount - headerRows;
  const maxRows = Math.min(dataRows, Math.floor(MAX_PROFILE_CELLS / cols), Math.round(optionalNumber(args, 'max_rows') ?? Infinity));
  const start = topLeft(data.address);

  let headers = Array.from({ length: cols }, (_, i) => columnLetters(start.col + i));
  if (hasHeaders) {
    const header = data.getRow(0);
    header.load('values');
    await context.sync();
    headers = header.values[0].map((v, i) => String(v) || headers[i]);
  }
  if (maxRows <= 0) return { sheet: sheetName, address: localAddress(data.address), columns: headers, rows: 0, note: 'There are no data rows.' };

  // Date columns are detected from the number format of the first data row.
  const firstRow = data.getRow(headerRows);
  firstRow.load('numberFormat');
  await context.sync();
  const dateColumns = firstRow.numberFormat[0].map(isDateFormat);

  const accumulators = headers.map(newAccumulator);
  const chunkRows = Math.max(1, Math.floor(PROFILE_CHUNK_CELLS / cols));
  for (let offset = 0; offset < maxRows; offset += chunkRows) {
    const rows = Math.min(chunkRows, maxRows - offset);
    const chunk = data.getCell(headerRows + offset, 0).getResizedRange(rows - 1, cols - 1);
    chunk.load('values, valueTypes');
    await context.sync();
    chunk.values.forEach((row, r) => row.forEach((value, c) => accumulate(accumulators[c], value, chunk.valueTypes[r][c])));
  }

  return {
    sheet: sheetName,
    address: localAddress(data.address),
    rows: dataRows,
    ...(maxRows < dataRows ? { profiledRows: maxRows, note: `Only the first ${maxRows} of ${dataRows} rows were profiled.` } : {}),
    columns: accumulators.map((acc, i) => summarize(acc, headers[i], columnLetters(start.col + i), dateColumns[i])),
  };
};

const requireAttachment = (args: Args) => {
  const id = requireString(args, 'attachment_id');
  const attachment = getAttachment(id);
  if (!attachment) throw new ToolError(`Attachment "${id}" is no longer available (attachments are kept only while the add-in is open). Ask the user to attach the file again.`);
  return attachment;
};

export const readAttachmentTool = async (_context: Excel.RequestContext, args: Args, options: ToolOptions) => {
  const attachment = requireAttachment(args);
  if (attachment.kind === 'text') {
    const startChar = Math.max(0, Math.round(optionalNumber(args, 'start_row') ?? 0));
    const text = attachment.text ?? '';
    return { attachment: attachment.id, start: startChar, length: text.length, text: text.slice(startChar, startChar + MAX_TEXT_READ) };
  }
  if (attachment.kind !== 'table') throw new ToolError('Only table and text attachments can be read with this tool; images and PDFs are already in the conversation.');

  const rows = attachment.rows ?? [];
  const width = rows[0]?.length ?? 1;
  const maxCells = options.maxReadCells ?? READ_CELLS_LIMITS.default;
  const startRow = Math.max(1, Math.round(optionalNumber(args, 'start_row') ?? 1));
  const count = Math.min(Math.round(optionalNumber(args, 'row_count') ?? Infinity), Math.max(1, Math.floor(maxCells / width)));
  const slice = rows.slice(startRow, startRow + count);
  return {
    attachment: attachment.id,
    headers: rows[0],
    startRow,
    rows: slice,
    totalDataRows: rows.length - 1,
    ...(startRow + slice.length < rows.length ? { note: `More rows are available from start_row ${startRow + slice.length}.` } : {}),
  };
};

/** Text that Excel would parse as a formula is stored as literal text. */
const safeCell = (value: AttachmentCell): AttachmentCell =>
  typeof value === 'string' && /^[=+\-@]/.test(value) ? `'${value}` : value;

export const importAttachment = async (context: Excel.RequestContext, args: Args) => {
  const attachment = requireAttachment(args);
  if (attachment.kind !== 'table') throw new ToolError('Only table attachments (CSV, TSV, JSON arrays) can be imported.');
  const asTable = optionalBoolean(args, 'as_table') ?? true;
  const tableName = optionalString(args, 'table_name');
  const rows = (attachment.rows ?? []).map(row => row.map(safeCell));
  if (rows.length === 0) throw new ToolError('The attachment is empty.');
  const width = rows[0].length;

  const { sheet, sheetName, range } = await resolveRange(context, args, 'start_cell');
  const target = range.getCell(0, 0).getResizedRange(rows.length - 1, width - 1);
  target.load('address');
  const used = sheet.getUsedRangeOrNullObject();
  await context.sync();
  const overlap = used.isNullObject ? null : target.getIntersectionOrNullObject(used);
  if (overlap) await context.sync();
  const targetIsEmpty = !overlap || overlap.isNullObject;

  const address = localAddress(target.address);
  const undoStep = targetIsEmpty ? { kind: 'newContent' as const, sheetName, address, tableName: undefined as string | undefined } : null;
  if (undoStep) recordUndo(undoStep);
  const snapshot = targetIsEmpty ? null : await snapshotRange(context, sheetName, target);

  const chunkRows = Math.max(1, Math.floor(IMPORT_CHUNK_CELLS / width));
  for (let offset = 0; offset < rows.length; offset += chunkRows) {
    const part = rows.slice(offset, offset + chunkRows);
    target.getCell(offset, 0).getResizedRange(part.length - 1, width - 1).values = part;
    await context.sync();
  }

  let createdTable: string | undefined;
  if (asTable) {
    const table = sheet.tables.add(target, true);
    table.load('name');
    await context.sync();
    const track = (name: string) => {
      createdTable = name;
      if (undoStep) undoStep.tableName = name;
      if (snapshot) snapshot.tableName = name;
    };
    track(table.name);
    if (tableName) {
      table.name = tableName;
      await context.sync();
      track(tableName);
    }
    table.style = DEFAULT_TABLE_STYLE;
  }
  target.format.autofitColumns();
  await context.sync();

  return {
    sheet: sheetName,
    address,
    rows: rows.length - 1,
    columns: width,
    ...(createdTable ? { table: createdTable } : {}),
    undoAvailable: targetIsEmpty || snapshot !== null,
  };
};
