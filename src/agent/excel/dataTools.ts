import {
  DEFAULT_TABLE_STYLE,
  ToolError,
  collectFormulaErrors,
  columnIndex,
  getTable,
  localAddress,
  matchName,
  optionalBoolean,
  optionalEnum,
  optionalString,
  parseArray,
  requireMatrix,
  requireString,
  resolveRange,
  stringList,
  toCell,
  topLeft,
  type Args,
} from './common';
import { recordUndo, snapshotRange, type RangeSnapshot } from './undo';
import { repairCollapsed } from './structureTools';

const MAX_RETURNED_RESULTS = 50;

/** Names, styles and auto-fits a freshly created table, keeping the undo snapshot in sync. */
const finishTable = async (
  context: Excel.RequestContext,
  table: Excel.Table,
  snapshot: RangeSnapshot | null,
  tableName: string | undefined,
  style: string
) => {
  table.load('name');
  await context.sync();
  if (snapshot) snapshot.tableName = table.name;

  if (tableName && tableName !== table.name) {
    table.name = tableName;
    await context.sync();
    if (snapshot) snapshot.tableName = tableName;
  }
  table.style = style;
  table.getRange().format.autofitColumns();
};

export const writeTable = async (context: Excel.RequestContext, args: Args) => {
  const headers = parseArray(args, 'headers').map(h => String(toCell(h)).trim());
  if (headers.length === 0) throw new ToolError('"headers" must contain at least one column name.');
  if (headers.some(h => !h)) throw new ToolError('Header names cannot be empty.');
  const duplicate = headers.find((h, i) => headers.findIndex(o => o.toLowerCase() === h.toLowerCase()) !== i);
  if (duplicate) throw new ToolError(`Header names must be unique; "${duplicate}" is repeated.`);
  const rows = requireMatrix(args, 'rows', { allowEmpty: true, width: headers.length });
  const tableName = optionalString(args, 'table_name');
  const style = optionalString(args, 'table_style') ?? DEFAULT_TABLE_STYLE;

  const { sheet, sheetName, range } = await resolveRange(context, args, 'start_cell');
  // A table always has at least one body row.
  const target = range.getCell(0, 0).getResizedRange(Math.max(rows.length, 1), headers.length - 1);
  const snapshot = await snapshotRange(context, sheetName, target);

  target.getRow(0).values = [headers];
  const table = sheet.tables.add(target, true);
  await finishTable(context, table, snapshot, tableName, style);

  // Body formulas are written after the table exists so structured references resolve.
  const body = table.getDataBodyRange();
  if (rows.length) body.formulas = rows;
  target.format.autofitColumns();
  body.load('address, values, valueTypes');
  await context.sync();
  await repairCollapsed(context, sheet, target, { rows: true, columns: true });

  const formulaErrors = collectFormulaErrors(body);
  return {
    table: tableName ?? table.name,
    sheet: sheetName,
    address: localAddress(target.address),
    rows: rows.length,
    columns: headers.length,
    ...(formulaErrors.length ? { formulaErrors } : {}),
    undoAvailable: snapshot !== null,
  };
};

export const convertRangeToTable = async (context: Excel.RequestContext, args: Args) => {
  const hasHeaders = optionalBoolean(args, 'has_headers') ?? true;
  const tableName = optionalString(args, 'table_name');
  const style = optionalString(args, 'table_style') ?? DEFAULT_TABLE_STYLE;
  const { sheet, sheetName, range } = await resolveRange(context, args, 'range_address');

  // A single cell means "the data around this cell".
  const target = range.cellCount === 1 ? range.getSurroundingRegion() : range;
  const snapshot = await snapshotRange(context, sheetName, target);

  const table = sheet.tables.add(target, hasHeaders);
  await finishTable(context, table, snapshot, tableName, style);
  const header = table.getHeaderRowRange();
  header.load('values');
  target.load('address');
  await context.sync();

  return {
    table: tableName ?? table.name,
    sheet: sheetName,
    address: localAddress(target.address),
    headers: header.values[0],
    undoAvailable: snapshot !== null,
  };
};

export const setRangeValuesOrFormulas = async (context: Excel.RequestContext, args: Args) => {
  const values = requireMatrix(args, 'values');
  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  const rows = values.length;
  const cols = values[0].length;

  let target = range;
  if (range.cellCount === 1) {
    target = range.getResizedRange(rows - 1, cols - 1);
  } else if (range.rowCount !== rows || range.columnCount !== cols) {
    throw new ToolError(
      `Range ${localAddress(range.address)} is ${range.rowCount}×${range.columnCount} but "values" is ${rows}×${cols}. ` +
      'Pass a range of the same size, or only its top-left cell.'
    );
  }

  const snapshot = await snapshotRange(context, sheetName, target);
  // `formulas` accepts both constants and formulas, parsed as if typed by the user.
  target.formulas = values;
  target.load('address, values, valueTypes');
  await context.sync();

  const formulaErrors = collectFormulaErrors(target);
  return {
    sheet: sheetName,
    address: localAddress(target.address),
    rows,
    columns: cols,
    ...(rows * cols <= MAX_RETURNED_RESULTS ? { results: target.values } : {}),
    ...(formulaErrors.length ? { formulaErrors } : {}),
    undoAvailable: snapshot !== null,
  };
};

const CLEAR_TARGETS = { contents: 'Contents', formats: 'Formats', all: 'All' } as const;

export const clearRange = async (context: Excel.RequestContext, args: Args) => {
  const what = optionalEnum(args, 'what', ['contents', 'formats', 'all'] as const) ?? 'contents';
  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  const snapshot = await snapshotRange(context, sheetName, range);
  range.clear(CLEAR_TARGETS[what]);
  await context.sync();
  return {
    sheet: sheetName,
    address: localAddress(range.address),
    cleared: what,
    undoAvailable: snapshot !== null,
    ...(snapshot ? {} : { note: 'The range was too large to snapshot, so this change cannot be undone.' }),
  };
};

/** Resolves a sort/filter column given as a header name or a column letter to an index within the range. */
const resolveColumn = (column: string, headers: string[], rangeAddress: string, useHeaders: boolean): number => {
  if (useHeaders) {
    const byName = headers.findIndex(h => h.toLowerCase() === column.trim().toLowerCase());
    if (byName !== -1) return byName;
  }
  if (/^[A-Z]{1,3}$/i.test(column.trim())) {
    const index = columnIndex(column.trim()) - topLeft(rangeAddress).col;
    if (index >= 0 && index < headers.length) return index;
  }
  throw new ToolError(`Column "${column}" is not in the range. ${useHeaders ? `Headers: ${headers.map(h => `"${h}"`).join(', ')}.` : 'Use a column letter.'}`);
};

export const sortRange = async (context: Excel.RequestContext, args: Args) => {
  const keys = parseArray(args, 'sort_by').map((item, i) => {
    if (typeof item === 'string') return { column: item, ascending: true };
    const { column, ascending } = (item ?? {}) as { column?: unknown; ascending?: unknown };
    if (typeof column !== 'string' && typeof column !== 'number') throw new ToolError(`"sort_by[${i}].column" is required.`);
    return { column: String(column), ascending: ascending !== false && ascending !== 'false' };
  });
  if (keys.length === 0) throw new ToolError('"sort_by" must contain at least one column.');

  const tableName = optionalString(args, 'table_name');
  let hasHeaders = optionalBoolean(args, 'has_headers') ?? true;
  let sheetName: string;
  let range: Excel.Range;
  let table: Excel.Table | null = null;

  if (tableName) {
    table = await getTable(context, tableName);
    const sheet = table.worksheet;
    sheet.load('name');
    range = table.getRange();
    await context.sync();
    sheetName = sheet.name;
    hasHeaders = true;
  } else {
    const resolved = await resolveRange(context, args, 'range_address');
    sheetName = resolved.sheetName;
    range = resolved.range.cellCount === 1 ? resolved.range.getSurroundingRegion() : resolved.range;
  }

  const firstRow = range.getRow(0);
  firstRow.load('values');
  range.load('address');
  await context.sync();
  const headers = firstRow.values[0].map(v => String(v));
  const fields = keys.map(k => ({ key: resolveColumn(k.column, headers, range.address, hasHeaders), ascending: k.ascending }));

  const snapshot = await snapshotRange(context, sheetName, range);
  if (table) table.sort.apply(fields);
  else range.sort.apply(fields, false, hasHeaders);
  await context.sync();

  return {
    sheet: sheetName,
    address: localAddress(range.address),
    sortedBy: fields.map(f => `${hasHeaders ? headers[f.key] : f.key} ${f.ascending ? 'ascending' : 'descending'}`),
    undoAvailable: snapshot !== null,
  };
};

export const filterTable = async (context: Excel.RequestContext, args: Args) => {
  const table = await getTable(context, requireString(args, 'table_name'));
  const clear = optionalBoolean(args, 'clear') ?? false;
  const values = stringList(args, 'values');
  const criteria1 = optionalString(args, 'criteria1');
  const criteria2 = optionalString(args, 'criteria2');
  const operator = optionalEnum(args, 'operator', ['And', 'Or'] as const);
  if (!clear && values.length === 0 && !criteria1) throw new ToolError('Provide "values", "criteria1", or "clear": true.');

  table.columns.load('items/name');
  await context.sync();
  const columnName = matchName(table.columns.items.map(c => c.name), requireString(args, 'column'), 'column');
  const column = table.columns.getItem(columnName);

  let previous: Excel.FilterCriteria | null = null;
  try {
    column.filter.load('criteria');
    await context.sync();
    previous = column.filter.criteria ?? null;
  } catch {
    // No readable filter state; undo will clear the filter.
  }
  recordUndo({ kind: 'filter', tableName: table.name, columnName, criteria: previous });

  if (clear) column.filter.clear();
  else if (values.length) column.filter.applyValuesFilter(values);
  else column.filter.applyCustomFilter(criteria1!, criteria2, operator);
  const visible = table.getDataBodyRange().getVisibleView();
  visible.load('rowCount');
  await context.sync();

  return { table: table.name, column: columnName, filter: clear ? 'cleared' : values.length ? { values } : { criteria1, criteria2, operator }, visibleRows: visible.rowCount };
};

const VALIDATION_TYPES = ['list', 'whole_number', 'decimal', 'date', 'text_length', 'custom'] as const;
const VALIDATION_OPERATORS = ['Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThan', 'LessThan', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo'] as const;

export const addDataValidation = async (context: Excel.RequestContext, args: Args) => {
  const type = optionalEnum(args, 'type', VALIDATION_TYPES);
  if (!type) throw new ToolError(`"type" is required: ${VALIDATION_TYPES.join(', ')}.`);

  let rule: Excel.DataValidationRule;
  if (type === 'list') {
    const values = stringList(args, 'list_values');
    const source = optionalString(args, 'list_source') ?? values.join(',');
    if (!source) throw new ToolError('Provide "list_values" or "list_source" (e.g. "=Lists!$A$2:$A$10").');
    if (source.length > 255 && !source.startsWith('=')) throw new ToolError('The list is longer than 255 characters; put the values in cells and use "list_source".');
    rule = { list: { inCellDropDown: true, source } };
  } else if (type === 'custom') {
    rule = { custom: { formula: requireString(args, 'formula') } };
  } else {
    const operator = optionalEnum(args, 'operator', VALIDATION_OPERATORS) ?? 'Between';
    const formula1 = requireString(args, 'value1');
    const formula2 = operator === 'Between' || operator === 'NotBetween' ? requireString(args, 'value2') : undefined;
    const basic = { formula1, formula2, operator };
    rule = type === 'whole_number' ? { wholeNumber: basic }
      : type === 'decimal' ? { decimal: basic }
      : type === 'date' ? { date: basic }
      : { textLength: basic };
  }

  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  const validation = range.dataValidation;
  validation.load('type');
  await context.sync();
  let previous: Excel.DataValidationRule | null = null;
  if (validation.type !== 'None' && validation.type !== 'Inconsistent') {
    validation.load('rule');
    await context.sync();
    previous = validation.rule;
  }
  recordUndo({ kind: 'validation', sheetName, address: localAddress(range.address), rule: previous });

  validation.clear();
  validation.rule = rule;
  validation.ignoreBlanks = optionalBoolean(args, 'allow_blank') ?? true;
  const errorMessage = optionalString(args, 'error_message');
  if (errorMessage) validation.errorAlert = { message: errorMessage, showAlert: true, style: 'Stop', title: 'Invalid value' };
  const inputMessage = optionalString(args, 'input_message');
  if (inputMessage) validation.prompt = { message: inputMessage, showPrompt: true, title: '' };
  await context.sync();

  return { sheet: sheetName, address: localAddress(range.address), validation: type };
};
