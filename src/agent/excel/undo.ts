import { BACKUP_SHEET_PREFIX, localAddress } from './common';

// Every mutating tool records how to revert its change. All changes made while answering
// one user message form a group that is undone together, newest first. The stack is plain
// JSON so it can be persisted per workbook.

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

export interface DimensionState {
  /** Column letter(s) like "C" or row number like "5". */
  address: string;
  size: number;
  hidden: boolean;
}

export type UndoStep =
  | RangeSnapshot
  /** Content written into an area that was empty: undone by clearing it (cheaper than a snapshot). */
  | { kind: 'newContent'; sheetName: string; address: string; tableName?: string }
  | { kind: 'createdSheet'; sheetName: string; previousActiveSheet: string }
  | { kind: 'renamedSheet'; from: string; to: string }
  | { kind: 'deletedSheet'; sheetName: string; backupName: string; position: number; tableNames: string[] }
  | { kind: 'sheetVisibility'; sheetName: string; visibility: Excel.SheetVisibility | 'Visible' | 'Hidden' | 'VeryHidden' }
  | { kind: 'insertedRange'; sheetName: string; address: string; shift: 'Down' | 'Right' }
  | { kind: 'deletedRange'; sheetName: string; address: string; shift: 'Up' | 'Left'; snapshot: RangeSnapshot | null }
  | { kind: 'dimensions'; sheetName: string; columns?: DimensionState[]; rows?: DimensionState[] }
  | { kind: 'freeze'; sheetName: string; location: string | null }
  | { kind: 'merge'; sheetName: string; address: string; action: 'merged' | 'unmerged'; mergedAreas?: string[]; across?: boolean }
  | { kind: 'comment'; id: string }
  | { kind: 'chart'; sheetName: string; chartName: string }
  | { kind: 'chartProps'; sheetName: string; chartName: string; props: ChartProps }
  | { kind: 'pivot'; sheetName: string; pivotName: string }
  | { kind: 'pivotLayout'; pivotName: string; layout: PivotLayout }
  | { kind: 'conditionalFormat'; sheetName: string; address: string; id: string }
  | { kind: 'filter'; tableName: string; columnName: string; criteria: Excel.FilterCriteria | null }
  | { kind: 'validation'; sheetName: string; address: string; rule: Excel.DataValidationRule | null }
  | { kind: 'name'; name: string };

export interface ChartProps {
  name: string;
  chartType: string;
  title: string | null;
  legendVisible: boolean;
  legendPosition: string;
  top: number;
  left: number;
  width: number;
  height: number;
}

export interface PivotLayout {
  rows: string[];
  columns: string[];
  filters: string[];
  values: { field: string; name: string; summarizeBy: string; numberFormat: string }[];
}

let undoStack: UndoStep[][] = [];
let currentGroup: UndoStep[] | null = null;
let persist: ((stack: UndoStep[][]) => void) | null = null;

const save = () => persist?.(undoStack);

/** Loads a persisted stack and saves every later change through `onChange`. */
export const initUndo = (stack: UndoStep[][], onChange: (stack: UndoStep[][]) => void) => {
  undoStack = stack;
  persist = onChange;
};

export const recordUndo = (step: UndoStep) => {
  if (currentGroup) {
    currentGroup.push(step);
  } else {
    undoStack.push([step]);
    save();
  }
};

export const beginUndoGroup = () => {
  currentGroup = [];
};

const backupsIn = (group: UndoStep[]) =>
  group.flatMap(step => (step.kind === 'deletedSheet' ? [step.backupName] : []));

const deleteSheets = (names: string[]) =>
  names.length === 0 ? Promise.resolve() : Excel.run(async context => {
    for (const name of names) context.workbook.worksheets.getItemOrNullObject(name).delete();
    await context.sync();
  }).catch(e => console.error('Failed to delete backup sheets', e));

export const endUndoGroup = () => {
  if (currentGroup?.length) undoStack.push(currentGroup);
  currentGroup = null;
  // Groups that fall off the stack can't be undone anymore: drop their backup sheets.
  while (undoStack.length > MAX_UNDO_GROUPS) void deleteSheets(backupsIn(undoStack.shift()!));
  save();
};

export const canUndo = () => undoStack.length > 0;

/** Deletes backup sheets that no undo step refers to (e.g. after the undo history was lost). */
export const cleanupOrphanBackups = () =>
  Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const referenced = new Set(undoStack.flatMap(backupsIn));
    for (const sheet of sheets.items) {
      if (sheet.name.startsWith(BACKUP_SHEET_PREFIX) && !referenced.has(sheet.name)) sheet.delete();
    }
    await context.sync();
  }).catch(e => console.error('Failed to clean up backup sheets', e));

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

/** Captures formulas, number formats and per-cell formatting, or null for ranges too large to snapshot. */
export const captureRange = async (context: Excel.RequestContext, sheetName: string, range: Excel.Range): Promise<RangeSnapshot | null> => {
  range.load('address, cellCount');
  await context.sync();
  if (range.cellCount > MAX_UNDO_CELLS) return null;

  range.load('formulas, numberFormat');
  const properties = range.getCellProperties(CELL_PROPERTIES);
  await context.sync();
  return {
    kind: 'range',
    sheetName,
    address: localAddress(range.address),
    formulas: range.formulas,
    numberFormat: range.numberFormat,
    cellProperties: properties.value,
  };
};

/**
 * Captures a range before a change and records it for undo.
 * Returns null (no undo) for ranges too large to snapshot.
 */
export const snapshotRange = async (context: Excel.RequestContext, sheetName: string, range: Excel.Range): Promise<RangeSnapshot | null> => {
  const snapshot = await captureRange(context, sheetName, range);
  // Recorded before the change is applied: a failed batch may still have been partially applied.
  if (snapshot) recordUndo(snapshot);
  return snapshot;
};

/** Snapshot of only the part of `range` that contains data (for whole rows or columns). */
export const captureUsedPart = async (context: Excel.RequestContext, sheet: Excel.Worksheet, sheetName: string, range: Excel.Range) => {
  const used = sheet.getUsedRangeOrNullObject();
  await context.sync();
  if (used.isNullObject) return { snapshot: null, empty: true };
  const part = range.getIntersectionOrNullObject(used);
  await context.sync();
  if (part.isNullObject) return { snapshot: null, empty: true };
  return { snapshot: await captureRange(context, sheetName, part), empty: false };
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

const syncIgnoringMissing = async (context: Excel.RequestContext) => {
  try {
    await context.sync();
  } catch {
    // The object was already removed or changed by the user.
  }
};

const undoStep = async (context: Excel.RequestContext, step: UndoStep) => {
  const workbook = context.workbook;
  switch (step.kind) {
    case 'range':
      return restoreRange(context, step);

    case 'newContent': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      if (step.tableName) {
        const table = sheet.tables.getItemOrNullObject(step.tableName);
        await context.sync();
        if (!table.isNullObject) table.convertToRange();
      }
      sheet.getRange(step.address).clear('All');
      return context.sync();
    }

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

    case 'deletedSheet': {
      const backup = await getSheetOrNull(context, step.backupName);
      if (!backup) throw new Error(`The backup of sheet "${step.sheetName}" no longer exists, so it can't be restored.`);
      backup.visibility = 'Visible';
      backup.name = step.sheetName;
      backup.position = step.position;
      // Copying a sheet renames its tables; give them their original names back.
      backup.tables.load('items/name');
      await context.sync();
      backup.tables.items.forEach((table, i) => {
        if (step.tableNames[i]) table.name = step.tableNames[i];
      });
      backup.activate();
      return syncIgnoringMissing(context);
    }

    case 'sheetVisibility': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.visibility = step.visibility;
      return context.sync();
    }

    case 'insertedRange': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.getRange(step.address).delete(step.shift === 'Down' ? 'Up' : 'Left');
      return context.sync();
    }

    case 'deletedRange': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.getRange(step.address).insert(step.shift === 'Up' ? 'Down' : 'Right');
      await context.sync();
      if (step.snapshot) await restoreRange(context, step.snapshot);
      return;
    }

    case 'dimensions': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      for (const c of step.columns ?? []) {
        const range = sheet.getRange(`${c.address}:${c.address}`);
        range.format.columnWidth = c.size;
        range.columnHidden = c.hidden;
      }
      for (const r of step.rows ?? []) {
        const range = sheet.getRange(`${r.address}:${r.address}`);
        range.format.rowHeight = r.size;
        range.rowHidden = r.hidden;
      }
      return context.sync();
    }

    case 'freeze': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.freezePanes.unfreeze();
      if (step.location) sheet.freezePanes.freezeAt(step.location);
      return context.sync();
    }

    case 'merge': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      if (step.action === 'merged') sheet.getRange(step.address).unmerge();
      else for (const area of step.mergedAreas ?? []) sheet.getRange(area).merge(false);
      return context.sync();
    }

    case 'comment': {
      const comment = workbook.comments.getItemOrNullObject(step.id);
      await context.sync();
      if (!comment.isNullObject) comment.delete();
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

    case 'chartProps': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      const chart = sheet.charts.getItemOrNullObject(step.chartName);
      await context.sync();
      if (chart.isNullObject) return;
      const p = step.props;
      chart.chartType = p.chartType as Excel.ChartType;
      chart.title.text = p.title ?? '';
      chart.title.visible = p.title !== null;
      chart.legend.visible = p.legendVisible;
      if (p.legendVisible) chart.legend.position = p.legendPosition as Excel.ChartLegendPosition;
      chart.top = p.top;
      chart.left = p.left;
      chart.width = p.width;
      chart.height = p.height;
      chart.name = p.name;
      return syncIgnoringMissing(context);
    }

    case 'pivot': {
      const pivot = workbook.pivotTables.getItemOrNullObject(step.pivotName);
      await context.sync();
      if (!pivot.isNullObject) pivot.delete();
      return context.sync();
    }

    case 'pivotLayout': {
      const pivot = workbook.pivotTables.getItemOrNullObject(step.pivotName);
      await context.sync();
      if (pivot.isNullObject) return;
      const collections = [pivot.rowHierarchies, pivot.columnHierarchies, pivot.filterHierarchies, pivot.dataHierarchies];
      for (const c of collections) c.load('items');
      await context.sync();
      pivot.rowHierarchies.items.forEach(h => pivot.rowHierarchies.remove(h));
      pivot.columnHierarchies.items.forEach(h => pivot.columnHierarchies.remove(h));
      pivot.filterHierarchies.items.forEach(h => pivot.filterHierarchies.remove(h));
      pivot.dataHierarchies.items.forEach(h => pivot.dataHierarchies.remove(h));
      await context.sync();
      const { layout } = step;
      layout.rows.forEach(f => pivot.rowHierarchies.add(pivot.hierarchies.getItem(f)));
      layout.columns.forEach(f => pivot.columnHierarchies.add(pivot.hierarchies.getItem(f)));
      layout.filters.forEach(f => pivot.filterHierarchies.add(pivot.hierarchies.getItem(f)));
      for (const v of layout.values) {
        const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(v.field));
        data.summarizeBy = v.summarizeBy as Excel.AggregationFunction;
        if (v.numberFormat) data.numberFormat = v.numberFormat;
        data.name = v.name;
      }
      return syncIgnoringMissing(context);
    }

    case 'conditionalFormat': {
      const sheet = await getSheetOrNull(context, step.sheetName);
      if (!sheet) return;
      sheet.getRange(step.address).conditionalFormats.getItem(step.id).delete();
      return syncIgnoringMissing(context);
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
    save();
  } catch (e) {
    // Every step is idempotent, so keeping the group allows retrying.
    undoStack.push(group);
    throw e;
  }
};
