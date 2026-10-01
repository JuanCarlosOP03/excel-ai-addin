import {
  ToolError,
  getSheet,
  localAddress,
  matchName,
  optionalBoolean,
  optionalEnum,
  optionalNumber,
  optionalString,
  parseArray,
  requireString,
  resolveRange,
  stringList,
  type Args,
} from './common';
import { recordUndo, type PivotLayout } from './undo';

export const CHART_TYPES = [
  'ColumnClustered', 'ColumnStacked', 'ColumnStacked100',
  'BarClustered', 'BarStacked', 'BarStacked100',
  'Line', 'LineMarkers', 'LineStacked',
  'Area', 'AreaStacked',
  'Pie', 'Doughnut',
  'XYScatter', 'XYScatterLines', 'XYScatterSmooth',
  'Radar', 'Histogram', 'Pareto', 'Waterfall', 'Treemap', 'Sunburst', 'Funnel', 'Boxwhisker',
] as const;
const CHARTS_WITHOUT_AXES = new Set<string>(['Pie', 'Doughnut', 'Treemap', 'Sunburst', 'Funnel']);
const LEGEND_POSITIONS = ['Right', 'Bottom', 'Top', 'Left', 'None'] as const;

export const createChart = async (context: Excel.RequestContext, args: Args) => {
  const chartType = optionalEnum(args, 'chart_type', CHART_TYPES) ?? 'ColumnClustered';
  const seriesBy = optionalEnum(args, 'series_by', ['Auto', 'Rows', 'Columns'] as const) ?? 'Auto';
  const legend = optionalEnum(args, 'legend_position', LEGEND_POSITIONS);
  const { sheet: dataSheet, sheetName: dataSheetName, range: data } = await resolveRange(context, args, 'data_range');

  const destinationName = optionalString(args, 'destination_sheet');
  const target = destinationName && destinationName.toLowerCase() !== dataSheetName.toLowerCase()
    ? await getSheet(context, destinationName)
    : dataSheet;
  const anchorArg = optionalString(args, 'anchor_cell');
  // By default the chart goes one column to the right of the data, or at B2 on another sheet.
  const anchor = anchorArg
    ? target.getRange(localAddress(anchorArg))
    : target === dataSheet ? data.getCell(0, 0).getOffsetRange(0, data.columnCount + 1) : target.getRange('B2');
  anchor.load('address');

  const chart = target.charts.add(chartType, data, seriesBy);
  chart.load('name');
  await context.sync();
  const sheetName = target.name;
  const undoStep = { kind: 'chart' as const, sheetName, chartName: chart.name };
  recordUndo(undoStep);

  const title = optionalString(args, 'title');
  if (title) {
    chart.title.text = title;
    chart.title.visible = true;
  }
  chart.setPosition(anchor);
  chart.width = optionalNumber(args, 'width') ?? 480;
  chart.height = optionalNumber(args, 'height') ?? 288;
  if (!CHARTS_WITHOUT_AXES.has(chartType)) {
    const xTitle = optionalString(args, 'x_axis_title');
    const yTitle = optionalString(args, 'y_axis_title');
    if (xTitle) chart.axes.categoryAxis.title.text = xTitle;
    if (yTitle) chart.axes.valueAxis.title.text = yTitle;
  }
  if (legend === 'None') {
    chart.legend.visible = false;
  } else if (legend) {
    chart.legend.visible = true;
    chart.legend.position = legend;
  }
  const chartName = optionalString(args, 'chart_name');
  if (chartName) chart.name = chartName;
  await context.sync();
  if (chartName) undoStep.chartName = chartName;

  return { chart: chartName ?? chart.name, sheet: sheetName, type: chartType, data: data.address, position: localAddress(anchor.address) };
};

const SUMMARIZE_FUNCTIONS = [
  'Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers',
  'StandardDeviation', 'StandardDeviationP', 'Variance', 'VarianceP',
] as const;

interface ValueField {
  field: string;
  summarizeBy?: (typeof SUMMARIZE_FUNCTIONS)[number];
  numberFormat?: string;
  name?: string;
}

const parseValueFields = (args: Args): ValueField[] =>
  (args.values === undefined || args.values === null ? [] : parseArray(args, 'values')).map((item, i) => {
    if (typeof item === 'string') return { field: item };
    if (!item || typeof item !== 'object') throw new ToolError(`"values[${i}]" must be an object with a "field".`);
    const value = item as Args;
    return {
      field: requireString(value, 'field'),
      summarizeBy: optionalEnum(value, 'summarize_by', SUMMARIZE_FUNCTIONS),
      numberFormat: optionalString(value, 'number_format'),
      name: optionalString(value, 'name'),
    };
  });

export const createPivotTable = async (context: Excel.RequestContext, args: Args) => {
  const rows = stringList(args, 'rows');
  const columns = stringList(args, 'columns');
  const filters = stringList(args, 'filters');
  const values = parseValueFields(args);
  if (rows.length === 0 && values.length === 0) throw new ToolError('Provide at least "rows" or "values".');

  // Source: a table name or a sheet-qualified range.
  const sourceArg = requireString(args, 'source');
  let source: Excel.Range | Excel.Table;
  let headerRow: Excel.Range;
  if (sourceArg.includes('!')) {
    const { range } = await resolveRange(context, { range_address: sourceArg }, 'range_address');
    source = range.cellCount === 1 ? range.getSurroundingRegion() : range;
    headerRow = source.getRow(0);
  } else {
    const table = context.workbook.tables.getItemOrNullObject(sourceArg);
    await context.sync();
    if (table.isNullObject) throw new ToolError(`"source" must be a table name or a sheet-qualified range like "Data!A1:F200"; "${sourceArg}" is neither.`);
    source = table;
    headerRow = table.getHeaderRowRange();
  }
  headerRow.load('values');

  const destination = await getSheet(context, requireString(args, 'destination_sheet'));
  const destinationCell = localAddress(optionalString(args, 'destination_cell') ?? 'A3');
  const pivots = context.workbook.pivotTables;
  pivots.load('items/name');
  await context.sync();

  // Validate every field before creating anything.
  const headers = headerRow.values[0].map(v => String(v)).filter(Boolean);
  const field = (name: string) => matchName(headers, name, 'field');
  const rowFields = rows.map(field);
  const columnFields = columns.map(field);
  const filterFields = filters.map(field);
  const valueFields = values.map(v => ({ ...v, field: field(v.field) }));

  const taken = new Set(pivots.items.map(p => p.name.toLowerCase()));
  let name = optionalString(args, 'name');
  if (name && taken.has(name.toLowerCase())) throw new ToolError(`A PivotTable named "${name}" already exists.`);
  if (!name) {
    let i = 1;
    while (taken.has(`pivottable${i}`)) i++;
    name = `PivotTable${i}`;
  }

  const pivot = destination.pivotTables.add(name, source, destination.getRange(destinationCell));
  await context.sync();
  recordUndo({ kind: 'pivot', sheetName: destination.name, pivotName: name });

  const hierarchy = (fieldName: string) => pivot.hierarchies.getItem(fieldName);
  rowFields.forEach(f => pivot.rowHierarchies.add(hierarchy(f)));
  columnFields.forEach(f => pivot.columnHierarchies.add(hierarchy(f)));
  filterFields.forEach(f => pivot.filterHierarchies.add(hierarchy(f)));
  for (const v of valueFields) {
    const data = pivot.dataHierarchies.add(hierarchy(v.field));
    if (v.summarizeBy) data.summarizeBy = v.summarizeBy;
    if (v.numberFormat) data.numberFormat = v.numberFormat;
    if (v.name) data.name = v.name;
  }
  const layout = pivot.layout.getRange();
  layout.load('address');
  await context.sync();

  return {
    pivotTable: name,
    sheet: destination.name,
    address: localAddress(layout.address),
    rows: rowFields,
    columns: columnFields,
    filters: filterFields,
    values: valueFields.map(v => `${v.summarizeBy ?? 'Sum'} of ${v.field}`),
  };
};

const getChart = async (context: Excel.RequestContext, sheet: Excel.Worksheet, name: string) => {
  const chart = sheet.charts.getItemOrNullObject(name);
  await context.sync();
  if (chart.isNullObject) {
    sheet.charts.load('items/name');
    await context.sync();
    const names = sheet.charts.items.map(c => `"${c.name}"`).join(', ');
    throw new ToolError(`Chart "${name}" was not found on sheet "${sheet.name}". ${names ? `Charts there: ${names}.` : 'That sheet has no charts.'}`);
  }
  return chart;
};

export const updateChart = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'sheet_name'));
  const chart = await getChart(context, sheet, requireString(args, 'chart_name'));
  const chartType = optionalEnum(args, 'chart_type', CHART_TYPES);
  const seriesBy = optionalEnum(args, 'series_by', ['Auto', 'Rows', 'Columns'] as const) ?? 'Auto';
  const legend = optionalEnum(args, 'legend_position', LEGEND_POSITIONS);
  const dataRange = optionalString(args, 'data_range');
  const title = typeof args.title === 'string' ? args.title.trim() : undefined;
  const anchorCell = optionalString(args, 'anchor_cell');
  const newName = optionalString(args, 'new_name');

  chart.load('name, chartType, top, left, width, height, title/text, title/visible, legend/visible, legend/position');
  await context.sync();
  const undoStep = {
    kind: 'chartProps' as const,
    sheetName: sheet.name,
    chartName: chart.name,
    props: {
      name: chart.name,
      chartType: chart.chartType,
      title: chart.title.visible ? chart.title.text : null,
      legendVisible: chart.legend.visible,
      legendPosition: chart.legend.position,
      top: chart.top,
      left: chart.left,
      width: chart.width,
      height: chart.height,
    },
  };
  recordUndo(undoStep);

  if (chartType) chart.chartType = chartType;
  if (dataRange) {
    const { range } = await resolveRange(context, { sheet_name: sheet.name, data_range: dataRange }, 'data_range');
    chart.setData(range, seriesBy);
  }
  if (title !== undefined) {
    chart.title.text = title;
    chart.title.visible = title !== '';
  }
  if (!CHARTS_WITHOUT_AXES.has(chartType ?? chart.chartType)) {
    const xTitle = optionalString(args, 'x_axis_title');
    const yTitle = optionalString(args, 'y_axis_title');
    if (xTitle) chart.axes.categoryAxis.title.text = xTitle;
    if (yTitle) chart.axes.valueAxis.title.text = yTitle;
  }
  if (legend === 'None') {
    chart.legend.visible = false;
  } else if (legend) {
    chart.legend.visible = true;
    chart.legend.position = legend;
  }
  if (anchorCell) chart.setPosition(sheet.getRange(localAddress(anchorCell)));
  const width = optionalNumber(args, 'width');
  const height = optionalNumber(args, 'height');
  if (width) chart.width = width;
  if (height) chart.height = height;
  if (newName) chart.name = newName;
  await context.sync();
  if (newName) undoStep.chartName = newName;

  return {
    chart: newName ?? undoStep.props.name,
    sheet: sheet.name,
    updated: Object.keys(args).filter(k => !['sheet_name', 'chart_name'].includes(k)),
    ...(dataRange ? { note: 'Undo restores the chart settings but not its previous data range.' } : {}),
  };
};

export const deleteChart = async (context: Excel.RequestContext, args: Args) => {
  const sheet = await getSheet(context, requireString(args, 'sheet_name'));
  const chart = await getChart(context, sheet, requireString(args, 'chart_name'));
  chart.load('name');
  await context.sync();
  const name = chart.name;
  chart.delete();
  await context.sync();
  return { deleted: name, sheet: sheet.name, undoAvailable: false };
};

const getPivot = async (context: Excel.RequestContext, name: string) => {
  const pivot = context.workbook.pivotTables.getItemOrNullObject(name);
  await context.sync();
  if (pivot.isNullObject) {
    const pivots = context.workbook.pivotTables;
    pivots.load('items/name');
    await context.sync();
    const names = pivots.items.map(p => `"${p.name}"`).join(', ');
    throw new ToolError(`PivotTable "${name}" does not exist. ${names ? `PivotTables: ${names}.` : 'The workbook has no PivotTables.'}`);
  }
  pivot.load('name');
  return pivot;
};

const captureLayout = async (context: Excel.RequestContext, pivot: Excel.PivotTable): Promise<PivotLayout> => {
  pivot.rowHierarchies.load('items/name');
  pivot.columnHierarchies.load('items/name');
  pivot.filterHierarchies.load('items/name');
  pivot.dataHierarchies.load('items/name, items/summarizeBy, items/numberFormat, items/field/name');
  await context.sync();
  return {
    rows: pivot.rowHierarchies.items.map(h => h.name),
    columns: pivot.columnHierarchies.items.map(h => h.name),
    filters: pivot.filterHierarchies.items.map(h => h.name),
    values: pivot.dataHierarchies.items.map(h => ({ field: h.field.name, name: h.name, summarizeBy: h.summarizeBy, numberFormat: h.numberFormat })),
  };
};

const sameName = (a: string, b: string | undefined) => a.toLowerCase() === b?.toLowerCase();

export const updatePivotTable = async (context: Excel.RequestContext, args: Args) => {
  const pivot = await getPivot(context, requireString(args, 'name'));
  const addRows = stringList(args, 'add_rows');
  const addColumns = stringList(args, 'add_columns');
  const addFilters = stringList(args, 'add_filters');
  const addValues = parseValueFields({ values: args.add_values });
  const removeFields = stringList(args, 'remove_fields');
  const refresh = optionalBoolean(args, 'refresh') ?? false;
  if (!refresh && [addRows, addColumns, addFilters, addValues, removeFields].every(list => list.length === 0)) {
    throw new ToolError('Nothing to change: provide fields to add or remove, or "refresh": true.');
  }

  pivot.hierarchies.load('items/name');
  const layout = await captureLayout(context, pivot);
  const field = (name: string) => matchName(pivot.hierarchies.items.map(h => h.name), name, 'field');
  // Validate everything before changing anything.
  const rows = addRows.map(field);
  const columns = addColumns.map(field);
  const filters = addFilters.map(field);
  const values = addValues.map(v => ({ ...v, field: field(v.field) }));
  const inUse = [...new Set([...layout.rows, ...layout.columns, ...layout.filters, ...layout.values.flatMap(v => [v.name, v.field])])];
  const removals = removeFields.map(name => matchName(inUse, name, 'field in this PivotTable'));

  recordUndo({ kind: 'pivotLayout', pivotName: pivot.name, layout });

  if (removals.length) {
    const removed = (name: string, fieldName?: string) => removals.some(r => sameName(r, name) || sameName(r, fieldName));
    pivot.rowHierarchies.items.filter(h => removed(h.name)).forEach(h => pivot.rowHierarchies.remove(h));
    pivot.columnHierarchies.items.filter(h => removed(h.name)).forEach(h => pivot.columnHierarchies.remove(h));
    pivot.filterHierarchies.items.filter(h => removed(h.name)).forEach(h => pivot.filterHierarchies.remove(h));
    pivot.dataHierarchies.items.filter(h => removed(h.name, h.field.name)).forEach(h => pivot.dataHierarchies.remove(h));
    await context.sync();
  }
  rows.forEach(f => pivot.rowHierarchies.add(pivot.hierarchies.getItem(f)));
  columns.forEach(f => pivot.columnHierarchies.add(pivot.hierarchies.getItem(f)));
  filters.forEach(f => pivot.filterHierarchies.add(pivot.hierarchies.getItem(f)));
  for (const v of values) {
    const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(v.field));
    if (v.summarizeBy) data.summarizeBy = v.summarizeBy;
    if (v.numberFormat) data.numberFormat = v.numberFormat;
    if (v.name) data.name = v.name;
  }
  if (refresh) pivot.refresh();
  const range = pivot.layout.getRange();
  range.load('address');
  const sheet = pivot.worksheet;
  sheet.load('name');
  await context.sync();

  return {
    pivotTable: pivot.name,
    sheet: sheet.name,
    address: localAddress(range.address),
    removed: removals,
    added: { rows, columns, filters, values: values.map(v => `${v.summarizeBy ?? 'Sum'} of ${v.field}`) },
    ...(refresh ? { refreshed: true } : {}),
  };
};

export const deletePivotTable = async (context: Excel.RequestContext, args: Args) => {
  const pivot = await getPivot(context, requireString(args, 'name'));
  await context.sync();
  const name = pivot.name;
  pivot.delete();
  await context.sync();
  return { deleted: name, undoAvailable: false };
};
