import {
  BACKUP_SHEET_PREFIX,
  ToolError,
  columnLetters,
  getSheet,
  isApiSupported,
  localAddress,
  optionalBoolean,
  optionalEnum,
  optionalNumber,
  requireApi,
  requireString,
  resolveRange,
  topLeft,
  type Args,
} from './common';
import { captureUsedPart, recordUndo, snapshotRange, type DimensionState } from './undo';

const MAX_DIMENSION_UNDO = 2000;

const isEntireRows = (address: string) => /^\$?\d+(:\$?\d+)?$/.test(localAddress(address));
const isEntireColumns = (address: string) => /^\$?[A-Z]{1,3}(:\$?[A-Z]{1,3})?$/i.test(localAddress(address));

export const insertRange = async (context: Excel.RequestContext, args: Args) => {
  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  range.load('isEntireRow, isEntireColumn');
  await context.sync();
  // Whole rows can only shift down and whole columns only right.
  const shift = range.isEntireRow ? 'Down' : range.isEntireColumn ? 'Right'
    : optionalEnum(args, 'shift', ['Down', 'Right'] as const) ?? 'Down';

  const inserted = range.insert(shift);
  inserted.load('address');
  await context.sync();
  const address = localAddress(inserted.address);
  recordUndo({ kind: 'insertedRange', sheetName, address, shift });
  return { sheet: sheetName, address, inserted: range.isEntireRow ? 'rows' : range.isEntireColumn ? 'columns' : `cells (shifted ${shift.toLowerCase()})` };
};

export const deleteRange = async (context: Excel.RequestContext, args: Args) => {
  const { sheet, sheetName, range } = await resolveRange(context, args, 'range_address');
  range.load('isEntireRow, isEntireColumn');
  await context.sync();
  const shift = range.isEntireRow ? 'Up' : range.isEntireColumn ? 'Left'
    : optionalEnum(args, 'shift', ['Up', 'Left'] as const) ?? 'Up';

  const { snapshot, empty } = await captureUsedPart(context, sheet, sheetName, range);
  const address = localAddress(range.address);
  const undoable = empty || snapshot !== null;
  if (undoable) recordUndo({ kind: 'deletedRange', sheetName, address, shift, snapshot });

  range.delete(shift);
  await context.sync();
  return {
    sheet: sheetName,
    address,
    deleted: range.isEntireRow ? 'rows' : range.isEntireColumn ? 'columns' : `cells (shifted ${shift.toLowerCase()})`,
    undoAvailable: undoable,
    note: undoable
      ? 'Formulas elsewhere that referred to the deleted cells now show #REF!; undo restores the cells but not those references.'
      : 'The deleted area was too large to back up, so this change cannot be undone.',
  };
};

export const deleteWorksheet = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'name'));
  sheet.load('name, position');
  sheet.tables.load('items/name');
  await context.sync();
  const { name, position } = sheet;

  // Keep a very hidden copy so the deletion can be undone.
  const backupName = `${BACKUP_SHEET_PREFIX}${Math.random().toString(36).slice(2, 10)}`;
  const backup = sheet.copy('End');
  backup.name = backupName;
  backup.visibility = 'VeryHidden';
  await context.sync();
  recordUndo({ kind: 'deletedSheet', sheetName: name, backupName, position, tableNames: sheet.tables.items.map(t => t.name) });

  sheet.delete();
  await context.sync();
  return {
    deleted: name,
    undoAvailable: true,
    note: 'Formulas in other sheets that referred to this sheet now show #REF!; undo restores the sheet but not those references.',
  };
};

export const setWorksheetVisibility = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'name'));
  const visible = optionalBoolean(args, 'visible');
  if (visible === undefined) throw new ToolError('"visible" is required.');
  sheet.load('name, visibility');
  await context.sync();
  if (sheet.name.startsWith(BACKUP_SHEET_PREFIX)) throw new ToolError('That sheet is an internal backup and cannot be changed.');

  recordUndo({ kind: 'sheetVisibility', sheetName: sheet.name, visibility: sheet.visibility });
  sheet.visibility = visible ? 'Visible' : 'Hidden';
  await context.sync();
  return { sheet: sheet.name, visibility: visible ? 'visible' : 'hidden' };
};

/** Records width/hidden of each column (or height/hidden of each row) for undo. */
const captureDimensions = async (range: Excel.Range, kind: 'columns' | 'rows', count: number, start: number, context: Excel.RequestContext) => {
  if (count > MAX_DIMENSION_UNDO) return null;
  const items = Array.from({ length: count }, (_, i) => {
    const part = kind === 'columns' ? range.getColumn(i) : range.getRow(i);
    part.load(kind === 'columns' ? 'format/columnWidth, columnHidden' : 'format/rowHeight, rowHidden');
    return part;
  });
  await context.sync();
  return items.map((part, i): DimensionState => ({
    address: kind === 'columns' ? columnLetters(start + i) : String(start + i),
    size: kind === 'columns' ? part.format.columnWidth : part.format.rowHeight,
    hidden: kind === 'columns' ? part.columnHidden : part.rowHidden,
  }));
};

const MAX_ROW_HEIGHT = 409;
const MAX_COLUMN_WIDTH = 1700;
/** Excel's default column width (8.43 characters) in points. */
const DEFAULT_COLUMN_WIDTH = 48;
const MAX_REPAIR_CHECK = 5000;

/** Validates a row height / column width in points. 0 would hide the row or column. */
const optionalSize = (args: Args, key: string, max: number): number | undefined => {
  const value = optionalNumber(args, key);
  if (value === undefined) return undefined;
  if (value <= 0) {
    throw new ToolError(`"${key}" must be greater than 0 (in points). A size of 0 hides the row/column: use "hidden": true to hide it, or "autofit_rows"/"autofit_columns" for automatic sizing.`);
  }
  return Math.min(value, max);
};

/**
 * Rows or columns that end up with zero size without being hidden on purpose look like they
 * disappeared (some models send 0 meaning "automatic", and auto-fit can collapse empty rows on
 * some hosts). Restores them to the sheet's standard row height / a default column width.
 */
export const repairCollapsed = async (
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
  check: { rows?: boolean; columns?: boolean }
): Promise<number> => {
  range.load('rowCount, columnCount');
  sheet.load('standardHeight');
  await context.sync();
  const rows = check.rows && range.rowCount <= MAX_REPAIR_CHECK
    ? Array.from({ length: range.rowCount }, (_, i) => {
      const part = range.getRow(i);
      part.load('format/rowHeight');
      return part;
    })
    : [];
  const columns = check.columns && range.columnCount <= MAX_REPAIR_CHECK
    ? Array.from({ length: range.columnCount }, (_, i) => {
      const part = range.getColumn(i);
      part.load('format/columnWidth');
      return part;
    })
    : [];
  if (rows.length === 0 && columns.length === 0) return 0;
  await context.sync();

  let repaired = 0;
  for (const part of rows) {
    if (part.format.rowHeight <= 0) {
      part.format.rowHeight = sheet.standardHeight || 15;
      repaired++;
    }
  }
  for (const part of columns) {
    if (part.format.columnWidth <= 0) {
      part.format.columnWidth = DEFAULT_COLUMN_WIDTH;
      repaired++;
    }
  }
  if (repaired) await context.sync();
  return repaired;
};

export const setRowsColumns = async (context: Excel.RequestContext, args: Args) => {
  const hidden = optionalBoolean(args, 'hidden');
  const autofitColumns = optionalBoolean(args, 'autofit_columns');
  const autofitRows = optionalBoolean(args, 'autofit_rows');
  // A size sent together with auto-fit (often 0, meaning "auto") is ignored in favor of auto-fit.
  const columnWidth = autofitColumns ? undefined : optionalSize(args, 'column_width', MAX_COLUMN_WIDTH);
  const rowHeight = autofitRows ? undefined : optionalSize(args, 'row_height', MAX_ROW_HEIGHT);
  if ([hidden, columnWidth, rowHeight, autofitColumns, autofitRows].every(v => v === undefined)) {
    throw new ToolError('Provide at least one of hidden, column_width, row_height, autofit_columns, autofit_rows.');
  }

  const { sheet, sheetName, range } = await resolveRange(context, args, 'range_address');
  const address = localAddress(range.address);
  const rowsTarget = isEntireRows(address) || rowHeight !== undefined || autofitRows;
  const columnsTarget = isEntireColumns(address) || columnWidth !== undefined || autofitColumns;
  // "hidden" applies to whole rows or whole columns, depending on the address.
  const hideRows = hidden !== undefined && isEntireRows(address);
  const hideColumns = hidden !== undefined && !hideRows;

  const { col, row } = topLeft(range.address);
  const columns = columnsTarget || hideColumns ? await captureDimensions(range, 'columns', range.columnCount, col, context) : undefined;
  const rows = rowsTarget || hideRows ? await captureDimensions(range, 'rows', range.rowCount, row, context) : undefined;
  const undoable = columns !== null && rows !== null;
  if (undoable) recordUndo({ kind: 'dimensions', sheetName, ...(columns ? { columns } : {}), ...(rows ? { rows } : {}) });

  if (hideRows) range.rowHidden = hidden!;
  if (hideColumns) range.columnHidden = hidden!;
  if (columnWidth !== undefined) range.format.columnWidth = columnWidth;
  if (rowHeight !== undefined) range.format.rowHeight = rowHeight;
  if (autofitColumns) range.format.autofitColumns();
  if (autofitRows) range.format.autofitRows();
  await context.sync();

  // Only rows/columns that were resized (not deliberately hidden) are checked.
  const repaired = await repairCollapsed(context, sheet, range, {
    rows: (rowHeight !== undefined || autofitRows) && !(hideRows && hidden),
    columns: (columnWidth !== undefined || autofitColumns) && !(hideColumns && hidden),
  });
  return {
    sheet: sheetName,
    address,
    undoAvailable: undoable,
    ...(repaired ? { note: `${repaired} rows/columns had collapsed to zero size and were restored to the default size.` } : {}),
  };
};

export const freezePanes = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'sheet_name'));
  const rows = Math.max(0, Math.round(optionalNumber(args, 'rows') ?? 0));
  const columns = Math.max(0, Math.round(optionalNumber(args, 'columns') ?? 0));

  const location = sheet.freezePanes.getLocationOrNullObject();
  location.load('address');
  await context.sync();
  recordUndo({ kind: 'freeze', sheetName: sheet.name, location: location.isNullObject ? null : localAddress(location.address) });

  sheet.freezePanes.unfreeze();
  if (rows && columns) sheet.freezePanes.freezeAt(sheet.getRange(`A1:${columnLetters(columns - 1)}${rows}`));
  else if (rows) sheet.freezePanes.freezeRows(rows);
  else if (columns) sheet.freezePanes.freezeColumns(columns);
  await context.sync();
  return { sheet: sheet.name, frozenRows: rows, frozenColumns: columns };
};

export const mergeCells = async (context: Excel.RequestContext, args: Args) => {
  const unmerge = optionalBoolean(args, 'unmerge') ?? false;
  const across = optionalBoolean(args, 'across') ?? false;
  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  const address = localAddress(range.address);

  if (unmerge) {
    let mergedAreas: string[] | undefined;
    if (isApiSupported('1.13')) {
      const areas = range.getMergedAreasOrNullObject();
      areas.load('address');
      await context.sync();
      mergedAreas = areas.isNullObject ? [] : areas.address.split(',').map(localAddress);
    }
    if (mergedAreas) recordUndo({ kind: 'merge', sheetName, address, action: 'unmerged', mergedAreas });
    range.unmerge();
    await context.sync();
    return { sheet: sheetName, address, unmerged: true, undoAvailable: mergedAreas !== undefined };
  }

  // Merging keeps only the top-left value of each merged area, so snapshot the contents.
  const snapshot = await snapshotRange(context, sheetName, range);
  recordUndo({ kind: 'merge', sheetName, address, action: 'merged', across });
  range.merge(across);
  await context.sync();
  return { sheet: sheetName, address, merged: across ? 'each row' : 'whole range', undoAvailable: snapshot !== null };
};

export const addComment = async (context: Excel.RequestContext, args: Args) => {
  requireApi('1.10', 'Comments');
  const text = requireString(args, 'text');
  const { sheetName, range } = await resolveRange(context, args, 'cell');
  if (range.cellCount !== 1) throw new ToolError('"cell" must be a single cell.');

  const comment = context.workbook.comments.add(range, text);
  comment.load('id');
  await context.sync();
  recordUndo({ kind: 'comment', id: comment.id });
  return { sheet: sheetName, address: localAddress(range.address), comment: text };
};
