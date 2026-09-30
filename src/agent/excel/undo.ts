import { localAddress } from './common';

// Every mutating tool records how to revert its change. All changes made while answering
// one user message form a group that is undone together, newest first.

const MAX_UNDO_CELLS = 20000;
const MAX_UNDO_GROUPS = 20;

export interface RangeSnapshot {
  kind: 'range';
  sheetName: string;
  address: string;
  formulas: unknown[][];
  numberFormat: unknown[][];
  cellProperties: Excel.CellProperties[][];
  /** Table created on top of this range, converted back to a plain range on undo. */
  tableName?: string;
}

export type UndoStep =
  | RangeSnapshot
  | { kind: 'createdSheet'; sheetName: string; previousActiveSheet: string }
  | { kind: 'renamedSheet'; from: string; to: string }
  | { kind: 'chart'; sheetName: string; chartName: string }
  | { kind: 'pivot'; sheetName: string; pivotName: string }
  | { kind: 'conditionalFormat'; sheetName: string; address: string; id: string }
  | { kind: 'filter'; tableName: string; columnName: string; criteria: Excel.FilterCriteria | null }
  | { kind: 'validation'; sheetName: string; address: string; rule: Excel.DataValidationRule | null }
  | { kind: 'name'; name: string };

const undoStack: UndoStep[][] = [];
let currentGroup: UndoStep[] | null = null;

export const recordUndo = (step: UndoStep) => {
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

const EDGES = ['top', 'bottom', 'left', 'right'] as const;
const BORDER_INDEXES = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideVertical', 'InsideHorizontal', 'DiagonalDown', 'DiagonalUp'] as const;

const CELL_PROPERTIES: Excel.CellPropertiesLoadOptions = {
  format: {
    fill: { color: true, pattern: true },
    font: { bold: true, italic: true, color: true, size: true, name: true },
    borders: { style: true, color: true, weight: true },
    horizontalAlignment: true,
    wrapText: true,
  },
};

/**
 * Captures formulas, number formats and per-cell formatting before a change and records it.
 * Returns null (no undo) for ranges too large to snapshot.
 */
export const snapshotRange = async (context: Excel.RequestContext, sheetName: string, range: Excel.Range): Promise<RangeSnapshot | null> => {
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
  recordUndo(snapshot);
  return snapshot;
};

/** Settable version of a snapshot cell: fills and borders only where they really existed. */
const toSettable = ({ format }: Excel.CellProperties): Excel.SettableCellProperties => {
  const borders: Excel.CellBorderCollection = {};
  for (const edge of EDGES) {
    const border = format?.borders?.[edge];
    if (border?.style && border.style !== 'None') borders[edge] = { style: border.style, color: border.color, weight: border.weight };
  }
  return {
    format: {
      font: format?.font,
      horizontalAlignment: format?.horizontalAlignment,
      wrapText: format?.wrapText,
      borders,
      ...(format?.fill?.pattern && format.fill.pattern !== 'None' ? { fill: { color: format.fill.color } } : {}),
    },
  };
};

const getSheetOrNull = async (context: Excel.RequestContext, name: string) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  return sheet.isNullObject ? null : sheet;
};

const restoreRange = async (context: Excel.RequestContext, step: RangeSnapshot) => {
  const sheet = await getSheetOrNull(context, step.sheetName);
  if (!sheet) return;

  if (step.tableName) {
    const table = sheet.tables.getItemOrNullObject(step.tableName);
    await context.sync();
    if (!table.isNullObject) table.convertToRange();
  }

  const range = sheet.getRange(step.address);
  // Clear fills and borders first and only re-apply the ones that existed: restoring the
  // reported "#FFFFFF, no pattern" of empty cells would hide the gridlines.
  range.format.fill.clear();
  for (const index of BORDER_INDEXES) range.format.borders.getItem(index).style = 'None';
  range.numberFormat = step.numberFormat;
  range.formulas = step.formulas;
  range.setCellProperties(step.cellProperties.map(row => row.map(toSettable)));
  await context.sync();
};

const undoStep = async (context: Excel.RequestContext, step: UndoStep) => {
  const workbook = context.workbook;
  switch (step.kind) {
    case 'range':
      return restoreRange(context, step);

    case 'createdSheet': {
      const sheet = workbook.worksheets.getItemOrNullObject(step.sheetName);
      const previous = workbook.worksheets.getItemOrNullObject(step.previousActiveSheet);
      await context.sync();
      if (sheet.isNullObject) return;
      if (!previous.isNullObject) previous.activate();
      sheet.delete();
      return context.sync();
    }

    case 'renamedSheet': {
      const sheet = await getSheetOrNull(context, step.to);
      if (!sheet) return;
      sheet.name = step.from;
      return context.sync();
    }

    case 'chart': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      const chart = sheet.charts.getItemOrNullObject(step.chartName);
      await context.sync();
      if (!chart.isNullObject) chart.delete();
      return context.sync();
    }

    case 'pivot': {
      const pivot = workbook.pivotTables.getItemOrNullObject(step.pivotName);
      await context.sync();
      if (!pivot.isNullObject) pivot.delete();
      return context.sync();
    }

    case 'conditionalFormat': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.getRange(step.address).conditionalFormats.getItem(step.id).delete();
      try {
        await context.sync();
      } catch {
        // Already removed by the user.
      }
      return;
    }

    case 'filter': {
      const table = workbook.tables.getItemOrNullObject(step.tableName);
      await context.sync();
      if (table.isNullObject) return;
      const column = table.columns.getItemOrNullObject(step.columnName);
      await context.sync();
      if (column.isNullObject) return;
      if (step.criteria) column.filter.apply(step.criteria);
      else column.filter.clear();
      return context.sync();
    }

    case 'validation': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      const validation = sheet.getRange(step.address).dataValidation;
      validation.clear();
      if (step.rule) validation.rule = step.rule;
      return context.sync();
    }

    case 'name': {
      const name = workbook.names.getItemOrNullObject(step.name);
      await context.sync();
      if (!name.isNullObject) name.delete();
      return context.sync();
    }
  }
};

/** Reverts every change made while answering the most recent user message. */
export const undoLastGroup = async (): Promise<void> => {
  const group = undoStack.pop();
  if (!group) throw new Error('Nothing to undo.');
  try {
    await Excel.run(async context => {
      for (const step of [...group].reverse()) await undoStep(context, step);
    });
  } catch (e) {
    // Every step is idempotent, so keeping the group allows retrying.
    undoStack.push(group);
    throw e;
  }
};
