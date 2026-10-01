/** An error caused by bad tool arguments or workbook state; its message is sent back to the model. */
export class ToolError extends Error {}

export type Cell = string | number | boolean;
export type Args = Record<string, unknown>;

export interface ToolOptions {
  /** Maximum number of cells returned by read_range. */
  maxReadCells?: number;
}

export type Executor = (context: Excel.RequestContext, args: Args, options: ToolOptions) => Promise<unknown>;

export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2';

/** Deleted sheets are kept as very hidden copies with this prefix until their undo expires. */
export const BACKUP_SHEET_PREFIX = '__xlai_bak_';
export const isBackupSheet = (name: string) => name.startsWith(BACKUP_SHEET_PREFIX);

/** All worksheets except internal backups. */
export const loadUserSheets = async (context: Excel.RequestContext, properties = 'items/name') => {
  const sheets = context.workbook.worksheets;
  sheets.load(properties);
  await context.sync();
  return sheets.items.filter(s => !isBackupSheet(s.name));
};

/** Whether this Excel supports the given ExcelApi requirement set (assumed outside Office, e.g. in tests). */
export const isApiSupported = (version: string): boolean => {
  try {
    if (typeof Office === 'undefined' || !Office.context?.requirements) return true;
    return Office.context.requirements.isSetSupported('ExcelApi', version);
  } catch {
    return true;
  }
};

export const requireApi = (version: string, feature: string) => {
  if (!isApiSupported(version)) throw new ToolError(`${feature} requires ExcelApi ${version}, which this version of Excel does not support.`);
};

/** Where a tool's change happened, so the chat can link to it. */
export interface ChangeLocation {
  sheet: string;
  address: string;
}

// ---------------------------------------------------------------------------
// Argument helpers
// ---------------------------------------------------------------------------

export const requireString = (args: Args, key: string): string => {
  const value = args[key];
  if (typeof value === 'number') return String(value);
  if (typeof value !== 'string' || !value.trim()) throw new ToolError(`"${key}" must be a non-empty string.`);
  return value.trim();
};

export const optionalString = (args: Args, key: string): string | undefined => {
  const value = args[key];
  if (value === undefined || value === null || value === '') return undefined;
  if (typeof value === 'number') return String(value);
  if (typeof value !== 'string') throw new ToolError(`"${key}" must be a string.`);
  return value.trim();
};

export const optionalBoolean = (args: Args, key: string): boolean | undefined => {
  const value = args[key];
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'boolean') return value;
  if (value === 'true' || value === 'false') return value === 'true';
  throw new ToolError(`"${key}" must be a boolean.`);
};

export const optionalNumber = (args: Args, key: string): number | undefined => {
  const value = args[key];
  if (value === undefined || value === null || value === '') return undefined;
  const n = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(n)) throw new ToolError(`"${key}" must be a number.`);
  return n;
};

/** Case-insensitive match against a fixed list; returns the canonical spelling. */
export const optionalEnum = <T extends string>(args: Args, key: string, allowed: readonly T[]): T | undefined => {
  const value = optionalString(args, key);
  if (value === undefined) return undefined;
  const match = allowed.find(a => a.toLowerCase() === value.toLowerCase());
  if (!match) throw new ToolError(`"${key}" must be one of: ${allowed.join(', ')}.`);
  return match;
};

/** Some models send arrays as JSON-encoded strings. */
export const parseArray = (args: Args, key: string): unknown[] => {
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

export const toCell = (value: unknown): Cell => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  throw new ToolError(`Unsupported cell value ${JSON.stringify(value)}; use strings, numbers or booleans.`);
};

/** Optional list of strings; missing means empty. */
export const stringList = (args: Args, key: string): string[] =>
  args[key] === undefined || args[key] === null ? [] : parseArray(args, key).map(v => String(toCell(v)).trim()).filter(Boolean);

export const requireMatrix = (args: Args, key: string, options: { allowEmpty?: boolean; width?: number } = {}): Cell[][] => {
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

/** Finds `wanted` in `names` ignoring case, or fails listing the valid names. */
export const matchName = (names: string[], wanted: string, what: string): string => {
  const match = names.find(n => n.toLowerCase() === wanted.trim().toLowerCase());
  if (!match) throw new ToolError(`Unknown ${what} "${wanted}". Available: ${names.map(n => `"${n}"`).join(', ')}.`);
  return match;
};

// ---------------------------------------------------------------------------
// Address helpers
// ---------------------------------------------------------------------------

/** Splits "'My Sheet'!A1:B2" into its sheet and local address parts. */
export const splitAddress = (address: string): { sheet?: string; local: string } => {
  const bang = address.lastIndexOf('!');
  if (bang === -1) return { local: address.trim() };
  let sheet = address.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, local: address.slice(bang + 1).trim() };
};

export const localAddress = (address: string) => splitAddress(address).local;

export const columnLetters = (index: number): string => {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
};

export const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** Zero-based column and one-based row of the top-left cell of an address like "Sheet1!$B$2:D9". */
export const topLeft = (address: string): { col: number; row: number } => {
  const match = /^\$?([A-Z]+)\$?(\d+)/i.exec(localAddress(address));
  return match ? { col: columnIndex(match[1]), row: Number(match[2]) } : { col: 0, row: 1 };
};

/** A1 address of the cell at offset (r, c) from the top-left cell of `address`. */
export const cellAt = (address: string, r: number, c: number): string => {
  const { col, row } = topLeft(address);
  return `${columnLetters(col + c)}${row + r}`;
};

/** Addresses of the cells that evaluate to an error (#NAME?, #REF!, ...), as feedback for the model. */
export const collectFormulaErrors = (range: Excel.Range, max = 20): string[] => {
  const errors: string[] = [];
  range.valueTypes.forEach((row, r) => row.forEach((type, c) => {
    if (String(type) === 'Error' && errors.length < max) errors.push(`${cellAt(range.address, r, c)}: ${range.values[r][c]}`);
  }));
  return errors;
};

// ---------------------------------------------------------------------------
// Workbook object lookup
// ---------------------------------------------------------------------------

export const getSheet = async (context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> => {
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

export const getTable = async (context: Excel.RequestContext, name: string): Promise<Excel.Table> => {
  const table = context.workbook.tables.getItemOrNullObject(name);
  await context.sync();
  if (table.isNullObject) {
    const tables = context.workbook.tables;
    tables.load('items/name');
    await context.sync();
    const existing = tables.items.length ? `Existing tables: ${tables.items.map(t => `"${t.name}"`).join(', ')}.` : 'The workbook has no tables.';
    throw new ToolError(`Table "${name}" does not exist. ${existing}`);
  }
  table.load('name');
  return table;
};

/**
 * Resolves `sheet_name` + an address argument to a loaded range (address, size) on the right sheet.
 * The sheet may also be given inside the address ("Sheet1!A1").
 */
export const resolveRange = async (context: Excel.RequestContext, args: Args, addressKey: string, sheetKey = 'sheet_name') => {
  const { sheet: addressSheet, local } = splitAddress(requireString(args, addressKey));
  const sheetName = optionalString(args, sheetKey) ?? addressSheet;
  if (!sheetName) throw new ToolError(`"${sheetKey}" is required.`);
  if (addressSheet && addressSheet.toLowerCase() !== sheetName.toLowerCase()) {
    throw new ToolError(`"${sheetKey}" is "${sheetName}" but "${addressKey}" refers to sheet "${addressSheet}".`);
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
