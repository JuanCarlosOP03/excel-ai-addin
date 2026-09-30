import {
  ToolError,
  getSheet,
  localAddress,
  matchName,
  optionalEnum,
  optionalNumber,
  optionalString,
  parseArray,
  requireString,
  resolveRange,
  stringList,
  type Args,
} from './common';
import { recordUndo } from './undo';

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
