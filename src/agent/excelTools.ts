import { READ_CELLS_LIMITS } from '../utils/storage';
import type { ToolName } from './tools';

/** An error caused by bad tool arguments or workbook state; its message is sent back to the model. */
export class ToolError extends Error {}

const MAX_READ_COLUMNS = 50;
const MAX_HEADER_COLUMNS = 30;
const SAMPLE_ROWS = 3;
const MAX_TABLES_IN_CONTEXT = 30;
const MAX_UNDO_CELLS = 20000;
const MAX_NUMBER_FORMAT_CELLS = 200000;
const MAX_UNDO_GROUPS = 20;
const MAX_REPORTED_ERRORS = 20;
const MAX_RETURNED_RESULTS = 50;

type Cell = string | number | boolean;
type Args = Record<string, unknown>;

export interface ToolOptions {
  /** Maximum number of cells returned by read_range. */
  maxReadCells?: number;
}

// ---------------------------------------------------------------------------
// Argument helpers
// ---------------------------------------------------------------------------

const requireString = (args: Args, key: string): string => {
  const value = args[key];
  if (typeof value !== 'string' || !value.trim()) throw new ToolError(`"${key}" must be a non-empty string.`);
  return value.trim();
};

const optionalString = (args: Args, key: string): string | undefined => {
  const value = args[key];
  if (value === undefined || value === null || value === '') return undefined;
  if (typeof value !== 'string') throw new ToolError(`"${key}" must be a string.`);
  return value.trim();
};

const optionalBoolean = (args: Args, key: string): boolean | undefined => {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'boolean') return value;
  if (value === 'true' || value === 'false') return value === 'true';
  throw new ToolError(`"${key}" must be a boolean.`);
};

/** Some models send arrays as JSON-encoded strings. */
const parseArray = (args: Args, key: string): unknown[] => {
  let value = args[key];
  if (typeof value === 'string') {
    try {
      value = JSON.parse(value);
    } catch {
      // Reported below.
    }
  }
  if (!Array.isArray(value)) throw new ToolError(`"${key}" must be an array.`);
  return value;
};

const toCell = (value: unknown): Cell => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  throw new ToolError(`Unsupported cell value ${JSON.stringify(value)}; use strings, numbers or booleans.`);
};

const requireMatrix = (args: Args, key: string, options: { allowEmpty?: boolean; width?: number } = {}): Cell[][] => {
  const matrix = parseArray(args, key).map((row, i) => {
    if (!Array.isArray(row)) throw new ToolError(`"${key}[${i}]" must be an array of cells.`);
    return row.map(toCell);
  });
  if (matrix.length === 0) {
    if (options.allowEmpty) return matrix;
    throw new ToolError(`"${key}" must contain at least one row.`);
  }
  const width = options.width ?? matrix[0].length;
  if (width === 0) throw new ToolError(`"${key}" rows must contain at least one cell.`);
  const bad = matrix.findIndex(row => row.length !== width);
  if (bad !== -1) {
    throw new ToolError(`Every row of "${key}" must have ${width} cells, but row ${bad} has ${matrix[bad].length}.`);
  }
  return matrix;
};

// ---------------------------------------------------------------------------
// Address helpers
// ---------------------------------------------------------------------------

/** Splits "'My Sheet'!A1:B2" into its sheet and local address parts. */
const splitAddress = (address: string): { sheet?: string; local: string } => {
  const bang = address.lastIndexOf('!');
  if (bang === -1) return { local: address.trim() };
  let sheet = address.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, local: address.slice(bang + 1).trim() };
};

const localAddress = (address: string) => splitAddress(address).local;

const columnLetters = (index: number): string => {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
};

const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** Addresses of the cells that evaluate to an error (#NAME?, #REF!, ...), as feedback for the model. */
const collectFormulaErrors = (range: Excel.Range): string[] => {
  const match = /^\$?([A-Z]+)\$?(\d+)/i.exec(localAddress(range.address));
  const startCol = match ? columnIndex(match[1]) : 0;
  const startRow = match ? Number(match[2]) : 1;
  const errors: string[] = [];
  range.valueTypes.forEach((row, r) => row.forEach((type, c) => {
    if (String(type) === 'Error' && errors.length < MAX_REPORTED_ERRORS) {
      errors.push(`${columnLetters(startCol + c)}${startRow + r}: ${range.values[r][c]}`);
    }
  }));
  return errors;
};

const getSheet = async (context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> => {
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (sheet.isNullObject) {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    throw new ToolError(`Worksheet "${name}" does not exist. Existing sheets: ${sheets.items.map(s => `"${s.name}"`).join(', ')}.`);
  }
  sheet.load('name');
  return sheet;
};

/** Resolves `sheet_name` + an address argument to a loaded range (address, size) on the right sheet. */
const resolveRange = async (context: Excel.RequestContext, args: Args, addressKey: string) => {
  const { sheet: addressSheet, local } = splitAddress(requireString(args, addressKey));
  const sheetName = optionalString(args, 'sheet_name') ?? addressSheet;
  if (!sheetName) throw new ToolError('"sheet_name" is required.');
  if (addressSheet && addressSheet.toLowerCase() !== sheetName.toLowerCase()) {
    throw new ToolError(`"sheet_name" is "${sheetName}" but "${addressKey}" refers to sheet "${addressSheet}".`);
  }
  if (!local) throw new ToolError(`"${addressKey}" is missing the cell address.`);

  const sheet = await getSheet(context, sheetName);
  const range = sheet.getRange(local);
  range.load('address, rowCount, columnCount, cellCount');
  try {
    await context.sync();
  } catch (e) {
    throw new ToolError(`"${local}" is not a valid range address.`, { cause: e });
  }
  return { sheet, sheetName: sheet.name, range };
};

// ---------------------------------------------------------------------------
// Undo journal: every mutating tool records how to revert it. All changes made
// while answering one user message form a group that is undone together.
// ---------------------------------------------------------------------------

interface RangeSnapshot {
  kind: 'range';
  sheetName: string;
  address: string;
  formulas: unknown[][];
  numberFormat: unknown[][];
  cellProperties: Excel.CellProperties[][];
  /** Table created on top of this range, converted back to a plain range on undo. */
  tableName?: string;
}

interface CreatedSheet {
  kind: 'sheet';
  sheetName: string;
  previousActiveSheet: string;
}

type UndoStep = RangeSnapshot | CreatedSheet;

const undoStack: UndoStep[][] = [];
let currentGroup: UndoStep[] | null = null;

const record = (step: UndoStep) => {
  if (currentGroup) currentGroup.push(step);
  else undoStack.push([step]);
};

export const beginUndoGroup = () => {
  currentGroup = [];
};

export const endUndoGroup = () => {
  if (currentGroup?.length) undoStack.push(currentGroup);
  currentGroup = null;
  if (undoStack.length > MAX_UNDO_GROUPS) undoStack.shift();
};

export const canUndo = () => undoStack.length > 0;

const CELL_PROPERTIES: Excel.CellPropertiesLoadOptions = {
  format: {
    fill: { color: true, pattern: true },
    font: { bold: true, italic: true, color: true },
    horizontalAlignment: true,
  },
};

/**
 * Captures formulas, number formats and per-cell formatting before a change.
 * Returns null (no undo) for ranges too large to snapshot.
 */
const snapshotRange = async (context: Excel.RequestContext, sheetName: string, range: Excel.Range): Promise<RangeSnapshot | null> => {
  range.load('address, cellCount');
  await context.sync();
  if (range.cellCount > MAX_UNDO_CELLS) return null;

  range.load('formulas, numberFormat');
  const properties = range.getCellProperties(CELL_PROPERTIES);
  await context.sync();

  const snapshot: RangeSnapshot = {
    kind: 'range',
    sheetName,
    address: localAddress(range.address),
    formulas: range.formulas,
    numberFormat: range.numberFormat,
    cellProperties: properties.value,
  };
  // Recorded before the change is applied: a failed batch may still have been partially applied.
  record(snapshot);
  return snapshot;
};

const restoreRange = async (context: Excel.RequestContext, step: RangeSnapshot) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject(step.sheetName);
  await context.sync();
  if (sheet.isNullObject) return;

  if (step.tableName) {
    const table = sheet.tables.getItemOrNullObject(step.tableName);
    await context.sync();
    if (!table.isNullObject) table.convertToRange();
  }

  const range = sheet.getRange(step.address);
  // Clearing first and only re-applying real fills keeps "no fill" cells transparent
  // (restoring their reported #FFFFFF would hide the gridlines).
  range.format.fill.clear();
  range.numberFormat = step.numberFormat;
  range.formulas = step.formulas;
  range.setCellProperties(step.cellProperties.map(row => row.map(({ format }) => ({
    format: {
      font: format?.font,
      horizontalAlignment: format?.horizontalAlignment,
      ...(format?.fill?.pattern && format.fill.pattern !== 'None' ? { fill: { color: format.fill.color } } : {}),
    },
  }))));
  await context.sync();
};

const deleteCreatedSheet = async (context: Excel.RequestContext, step: CreatedSheet) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject(step.sheetName);
  const previous = context.workbook.worksheets.getItemOrNullObject(step.previousActiveSheet);
  await context.sync();
  if (sheet.isNullObject) return;
  if (!previous.isNullObject) previous.activate();
  sheet.delete();
  await context.sync();
};

/** Reverts every change made while answering the most recent user message. */
export const undoLastGroup = async (): Promise<void> => {
  const group = undoStack.pop();
  if (!group) throw new Error('Nothing to undo.');
  try {
    await Excel.run(async context => {
      for (const step of [...group].reverse()) {
        if (step.kind === 'range') await restoreRange(context, step);
        else await deleteCreatedSheet(context, step);
      }
    });
  } catch (e) {
    // Every step is idempotent, so keeping the group allows retrying.
    undoStack.push(group);
    throw e;
  }
};

// ---------------------------------------------------------------------------
// Workbook overview (used for the system prompt and get_workbook_context)
// ---------------------------------------------------------------------------

const loadSheets = async (context: Excel.RequestContext) => {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name, items/visibility');
  const active = sheets.getActiveWorksheet();
  active.load('name');
  await context.sync();

  const used = sheets.items.map(sheet => {
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

// ---------------------------------------------------------------------------
// Tool executors
// ---------------------------------------------------------------------------

const getWorkbookContext = async (context: Excel.RequestContext) => {
  const { activeSheet, used } = await loadSheets(context);

  const tables = context.workbook.tables;
  tables.load('items/name');
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

  return {
    activeSheet,
    selection,
    sheets: used.map(({ sheet, range }, i) => {
      if (range.isNullObject) return { name: sheet.name, empty: true };
      const values = previews[i]?.values ?? [];
      return {
        name: sheet.name,
        ...(sheet.visibility !== 'Visible' ? { visibility: sheet.visibility } : {}),
        usedRange: localAddress(range.address),
        rows: range.rowCount,
        columns: range.columnCount,
        firstRow: values[0],
        ...(sheet.name === activeSheet ? { sampleRows: values.slice(1) } : {}),
      };
    }),
    tables: tableInfo.map(({ table, sheet, range, header }) => ({
      name: table.name,
      sheet: sheet.name,
      range: localAddress(range.address),
      headers: header.values[0],
    })),
    ...(tables.items.length > MAX_TABLES_IN_CONTEXT ? { note: `Only the first ${MAX_TABLES_IN_CONTEXT} of ${tables.items.length} tables are listed.` } : {}),
  };
};

const readRange = async (context: Excel.RequestContext, args: Args, options: ToolOptions) => {
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

const activateWorksheet = async (context: Excel.RequestContext, args: Args) => {
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

const createWorksheet = async (context: Excel.RequestContext, args: Args) => {
  const name = requireString(args, 'name');
  validateSheetName(name);

  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  const active = sheets.getActiveWorksheet();
  active.load('name');
  await context.sync();

  const existing = sheets.items.find(s => s.name.toLowerCase() === name.toLowerCase());
  if (existing) throw new ToolError(`A worksheet named "${existing.name}" already exists. Use it or choose another name.`);

  const sheet = sheets.add(name);
  sheet.activate();
  await context.sync();
  record({ kind: 'sheet', sheetName: name, previousActiveSheet: active.name });
  return { created: name, activated: true };
};

const writeTable = async (context: Excel.RequestContext, args: Args) => {
  const headers = parseArray(args, 'headers').map(h => String(toCell(h)).trim());
  if (headers.length === 0) throw new ToolError('"headers" must contain at least one column name.');
  if (headers.some(h => !h)) throw new ToolError('Header names cannot be empty.');
  const duplicate = headers.find((h, i) => headers.findIndex(o => o.toLowerCase() === h.toLowerCase()) !== i);
  if (duplicate) throw new ToolError(`Header names must be unique; "${duplicate}" is repeated.`);
  const rows = requireMatrix(args, 'rows', { allowEmpty: true, width: headers.length });
  const tableName = optionalString(args, 'table_name');
  const style = optionalString(args, 'table_style') ?? 'TableStyleMedium2';

  const { sheet, sheetName, range } = await resolveRange(context, args, 'start_cell');
  // A table always has at least one body row.
  const target = range.getCell(0, 0).getResizedRange(Math.max(rows.length, 1), headers.length - 1);
  const snapshot = await snapshotRange(context, sheetName, target);

  target.getRow(0).values = [headers];
  const table = sheet.tables.add(target, true);
  table.load('name');
  await context.sync();
  if (snapshot) snapshot.tableName = table.name;

  if (tableName && tableName !== table.name) {
    table.name = tableName;
    await context.sync();
    if (snapshot) snapshot.tableName = tableName;
  }

  table.style = style;
  // Body formulas are written after the table exists so structured references resolve.
  const body = table.getDataBodyRange();
  if (rows.length) body.formulas = rows;
  target.format.autofitColumns();
  body.load('address, values, valueTypes');
  target.load('address');
  await context.sync();

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

const setRangeValuesOrFormulas = async (context: Excel.RequestContext, args: Args) => {
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

// String literals rather than Excel.* enums: the Excel namespace is only defined once Office has loaded.
const ALIGNMENTS: Record<string, 'Left' | 'Center' | 'Right'> = { left: 'Left', center: 'Center', right: 'Right' };

const formatRange = async (context: Excel.RequestContext, args: Args) => {
  const bold = optionalBoolean(args, 'font_bold');
  const italic = optionalBoolean(args, 'font_italic');
  const fontColor = optionalString(args, 'font_color');
  const fillColor = optionalString(args, 'fill_color');
  const numberFormat = optionalString(args, 'num_format');
  const alignmentArg = optionalString(args, 'horizontal_alignment');
  const autofit = optionalBoolean(args, 'autofit_columns');

  const alignment = alignmentArg ? ALIGNMENTS[alignmentArg.toLowerCase()] : undefined;
  if (alignmentArg && !alignment) throw new ToolError('"horizontal_alignment" must be Left, Center or Right.');
  const applied = Object.entries({ bold, italic, fontColor, fillColor, numberFormat, alignment, autofit })
    .filter(([, value]) => value !== undefined)
    .map(([key]) => key);
  if (applied.length === 0) throw new ToolError('Provide at least one formatting property.');

  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  if (numberFormat && range.cellCount > MAX_NUMBER_FORMAT_CELLS) {
    throw new ToolError(`"num_format" can be applied to at most ${MAX_NUMBER_FORMAT_CELLS} cells at once, but the range has ${range.cellCount}. Limit it to the rows that contain data.`);
  }

  const snapshot = await snapshotRange(context, sheetName, range);
  if (bold !== undefined) range.format.font.bold = bold;
  if (italic !== undefined) range.format.font.italic = italic;
  if (fontColor) range.format.font.color = fontColor;
  if (fillColor) {
    if (['none', 'transparent', 'no fill'].includes(fillColor.toLowerCase())) range.format.fill.clear();
    else range.format.fill.color = fillColor;
  }
  if (alignment) range.format.horizontalAlignment = alignment;
  if (numberFormat) {
    range.numberFormat = Array.from({ length: range.rowCount }, () => new Array<string>(range.columnCount).fill(numberFormat));
  }
  if (autofit) range.format.autofitColumns();
  await context.sync();

  return {
    sheet: sheetName,
    address: localAddress(range.address),
    applied,
    undoAvailable: snapshot !== null,
    ...(snapshot ? {} : { note: 'The range was too large to snapshot, so this change cannot be undone.' }),
  };
};

const EXECUTORS: Record<ToolName, (context: Excel.RequestContext, args: Args, options: ToolOptions) => Promise<unknown>> = {
  get_workbook_context: getWorkbookContext,
  read_range: readRange,
  activate_worksheet: activateWorksheet,
  create_worksheet: createWorksheet,
  write_table: writeTable,
  set_range_values_or_formulas: setRangeValuesOrFormulas,
  format_range: formatRange,
};

/** Runs one tool in its own Excel batch. Throws ToolError or Office errors on failure. */
export const executeTool = (name: ToolName, args: Args, options: ToolOptions = {}): Promise<unknown> =>
  Excel.run(context => EXECUTORS[name](context, args, options));
