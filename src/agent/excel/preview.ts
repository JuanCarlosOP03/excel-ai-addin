import { getAttachment } from '../attachments';
import { localAddress, splitAddress } from './common';

const PREVIEW_ROWS = 8;
const PREVIEW_COLUMNS = 6;

/** What a pending change will do to a range, shown in the approval card. */
export interface GridPreview {
  sheet: string;
  address: string;
  /** Total size of the change (the grids are clipped to a few rows and columns). */
  rows: number;
  columns: number;
  before: unknown[][];
  /** Missing when the cells are removed (delete_range). */
  after?: unknown[][];
}

const clip = (matrix: unknown[][]) => matrix.slice(0, PREVIEW_ROWS).map(row => row.slice(0, PREVIEW_COLUMNS));

const asMatrix = (value: unknown): unknown[][] | null => {
  let parsed = value;
  if (typeof parsed === 'string') {
    try {
      parsed = JSON.parse(parsed);
    } catch {
      return null;
    }
  }
  return Array.isArray(parsed) && parsed.every(Array.isArray) ? (parsed as unknown[][]) : null;
};

/** Target top-left cell and the values that will be written, per tool. */
const plannedWrite = (name: string, args: Record<string, unknown>): { address: string; after: unknown[][] | null; width?: number; height?: number } | null => {
  switch (name) {
    case 'set_range_values_or_formulas': {
      const values = asMatrix(args.values);
      return values ? { address: String(args.range_address ?? ''), after: values } : null;
    }
    case 'write_table': {
      const headers = Array.isArray(args.headers) ? args.headers : null;
      const rows = asMatrix(args.rows) ?? [];
      return headers ? { address: String(args.start_cell ?? ''), after: [headers, ...rows] } : null;
    }
    case 'import_attachment': {
      const rows = getAttachment(String(args.attachment_id ?? ''))?.rows;
      return rows ? { address: String(args.start_cell ?? ''), after: rows } : null;
    }
    case 'clear_range':
      return args.what === 'formats' ? null : { address: String(args.range_address ?? ''), after: null };
    case 'delete_range':
      return { address: String(args.range_address ?? ''), after: null };
    default:
      return null;
  }
};

/** Reads the current contents of the cells a tool call would change. Returns null when not applicable. */
export const previewToolCall = async (name: string, args: Record<string, unknown>): Promise<GridPreview | null> => {
  const plan = plannedWrite(name, args);
  if (!plan?.address) return null;
  try {
    return await Excel.run(async context => {
      const { sheet: addressSheet, local } = splitAddress(plan.address);
      const sheetName = typeof args.sheet_name === 'string' && args.sheet_name ? args.sheet_name : addressSheet;
      const sheet = sheetName ? context.workbook.worksheets.getItem(sheetName) : context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      let target = sheet.getRange(local);
      target.load('cellCount, rowCount, columnCount');
      await context.sync();

      if (plan.after && target.cellCount === 1) {
        target = target.getResizedRange(plan.after.length - 1, Math.max(1, plan.after[0]?.length ?? 1) - 1);
        target.load('rowCount, columnCount');
      }
      target.load('address');
      await context.sync();

      const rows = target.rowCount;
      const columns = target.columnCount;
      const visible = target.getCell(0, 0).getResizedRange(Math.min(rows, PREVIEW_ROWS) - 1, Math.min(columns, PREVIEW_COLUMNS) - 1);
      visible.load('formulas');
      await context.sync();

      const before = visible.formulas;
      const after = name === 'delete_range' ? undefined
        : plan.after ? clip(plan.after)
        : before.map(row => row.map(() => ''));
      return { sheet: sheet.name, address: localAddress(target.address), rows, columns, before, after };
    });
  } catch {
    return null;
  }
};

/** Activates the sheet and selects the range of a reference like "Sales!B5" or "'Q1 Data'!A1:D9". */
export const selectReference = (reference: string) =>
  Excel.run(async context => {
    const { sheet: sheetName, local } = splitAddress(reference);
    const sheet = sheetName ? context.workbook.worksheets.getItem(sheetName) : context.workbook.worksheets.getActiveWorksheet();
    sheet.activate();
    sheet.getRange(local).select();
    await context.sync();
  });
