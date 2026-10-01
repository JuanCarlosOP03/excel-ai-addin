import {
  ToolError,
  cellAt,
  getSheet,
  loadUserSheets,
  localAddress,
  optionalEnum,
  optionalNumber,
  optionalString,
  requireApi,
  resolveRange,
  splitAddress,
  type Args,
} from './common';
import { quoteSheetName } from '../../utils/cellReferences';

const MAX_TRACE_DEPTH = 3;
const MAX_TRACE_CELLS = 60;
const MAX_ERRORS = 100;

interface TracedCell {
  cell: string;
  formula?: string;
  value: unknown;
  /** Distance from the starting cell (1 = direct). */
  level: number;
}

/** Expands sheet-qualified area addresses into individual cell addresses, up to `limit`. */
const loadAreaCells = async (context: Excel.RequestContext, addresses: string[], limit: number) => {
  const ranges = addresses.map(address => {
    const { sheet, local } = splitAddress(address);
    const range = sheet ? context.workbook.worksheets.getItem(sheet).getRange(local) : context.workbook.worksheets.getActiveWorksheet().getRange(local);
    range.load('address, formulas, values, cellCount');
    return { sheet: sheet ?? '', range };
  });
  await context.sync();

  const cells: { cell: string; formula: unknown; value: unknown }[] = [];
  let total = 0;
  for (const { sheet, range } of ranges) {
    total += range.cellCount;
    range.formulas.forEach((row, r) => row.forEach((formula, c) => {
      if (cells.length < limit) cells.push({ cell: `${sheet ? `${quoteSheet(sheet)}!` : ''}${cellAt(range.address, r, c)}`, formula, value: range.values[r][c] });
    }));
  }
  return { cells, total };
};

const quoteSheet = quoteSheetName;

export const traceFormula = async (context: Excel.RequestContext, args: Args) => {
  const direction = optionalEnum(args, 'direction', ['precedents', 'dependents'] as const) ?? 'precedents';
  requireApi(direction === 'precedents' ? '1.12' : '1.13', `Tracing ${direction}`);
  const depth = Math.min(MAX_TRACE_DEPTH, Math.max(1, Math.round(optionalNumber(args, 'depth') ?? 1)));
  const { sheetName, range } = await resolveRange(context, args, 'cell');
  if (range.cellCount !== 1) throw new ToolError('"cell" must be a single cell.');
  range.load('formulas, values');
  await context.sync();

  const start = `${quoteSheet(sheetName)}!${localAddress(range.address)}`;
  const visited = new Set<string>([start.toLowerCase()]);
  const traced: TracedCell[] = [];
  let frontier = [start];
  let truncated = false;

  for (let level = 1; level <= depth && frontier.length && !truncated; level++) {
    const next: string[] = [];
    for (const address of frontier) {
      const { sheet, local } = splitAddress(address);
      const cell = context.workbook.worksheets.getItem(sheet!).getRange(local);
      const areas = direction === 'precedents' ? cell.getDirectPrecedents() : cell.getDirectDependents();
      areas.load('addresses');
      try {
        await context.sync();
      } catch {
        // No precedents/dependents (Excel throws ItemNotFound).
        continue;
      }
      const { cells, total } = await loadAreaCells(context, areas.addresses, MAX_TRACE_CELLS - traced.length);
      if (cells.length < total) truncated = true;
      for (const c of cells) {
        if (visited.has(c.cell.toLowerCase())) continue;
        visited.add(c.cell.toLowerCase());
        const isFormula = typeof c.formula === 'string' && c.formula.startsWith('=');
        traced.push({ cell: c.cell, ...(isFormula ? { formula: c.formula as string } : {}), value: c.value, level });
        if (isFormula || direction === 'dependents') next.push(c.cell);
      }
    }
    frontier = next;
  }

  const startFormula = range.formulas[0][0];
  return {
    cell: start,
    formula: typeof startFormula === 'string' && startFormula.startsWith('=') ? startFormula : null,
    value: range.values[0][0],
    direction,
    [direction]: traced,
    ...(traced.length === 0 ? { note: direction === 'precedents' ? 'The cell has no precedents (it is a constant or its formula has no cell references).' : 'No formula depends on this cell.' } : {}),
    ...(truncated ? { note: `Only the first ${MAX_TRACE_CELLS} cells are listed.` } : {}),
  };
};

export const findFormulaErrors = async (context: Excel.RequestContext, args: Args) => {
  const sheetName = optionalString(args, 'sheet_name');
  const rangeAddress = optionalString(args, 'range_address');

  const targets: { sheet: string; range: Excel.Range }[] = [];
  if (rangeAddress) {
    const { sheetName: name, range } = await resolveRange(context, args, 'range_address');
    targets.push({ sheet: name, range });
  } else {
    let sheets: Excel.Worksheet[];
    if (sheetName) {
      sheets = [await getSheet(context, sheetName)];
    } else {
      sheets = await loadUserSheets(context);
    }
    for (const sheet of sheets) {
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      if (!used.isNullObject) targets.push({ sheet: sheet.name, range: used });
    }
  }

  const found = targets.map(t => {
    const areas = t.range.getSpecialCellsOrNullObject('Formulas', 'Errors');
    areas.load('cellCount');
    return { ...t, areas };
  });
  await context.sync();

  const hits = found.filter(f => !f.areas.isNullObject);
  for (const hit of hits) hit.areas.areas.load('items/address, items/values, items/formulas');
  await context.sync();

  const total = hits.reduce((sum, h) => sum + h.areas.cellCount, 0);
  const errors: { cell: string; error: unknown; formula: unknown }[] = [];
  const byType: Record<string, number> = {};
  for (const hit of hits) {
    for (const area of hit.areas.areas.items) {
      area.values.forEach((row, r) => row.forEach((value, c) => {
        byType[String(value)] = (byType[String(value)] ?? 0) + 1;
        if (errors.length < MAX_ERRORS) errors.push({ cell: `${quoteSheet(hit.sheet)}!${cellAt(area.address, r, c)}`, error: value, formula: area.formulas[r][c] });
      }));
    }
  }

  return {
    totalErrors: total,
    ...(total ? { byType } : {}),
    errors,
    ...(total > errors.length ? { note: `Only the first ${errors.length} of ${total} errors are listed.` } : {}),
    ...(total === 0 ? { note: 'No formula errors found.' } : {}),
  };
};
