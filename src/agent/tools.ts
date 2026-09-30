import type { ToolDefinition } from './llmClient';
import { CHART_TYPES } from './excel/analysisTools';

export type ToolName =
  | 'get_workbook_context'
  | 'read_range'
  | 'search_workbook'
  | 'activate_worksheet'
  | 'create_worksheet'
  | 'rename_worksheet'
  | 'create_named_range'
  | 'write_table'
  | 'convert_range_to_table'
  | 'set_range_values_or_formulas'
  | 'clear_range'
  | 'sort_range'
  | 'filter_table'
  | 'add_data_validation'
  | 'format_range'
  | 'add_conditional_format'
  | 'create_chart'
  | 'create_pivot_table';

/** Tools that don't change the workbook: they run without approval and need no undo. */
const READ_ONLY_TOOLS = new Set<ToolName>(['get_workbook_context', 'read_range', 'search_workbook', 'activate_worksheet']);
export const isMutatingTool = (name: ToolName) => !READ_ONLY_TOOLS.has(name);

// Cells are typed as strings for maximum provider compatibility (some reject union types).
// The executors also accept numbers, booleans and null.
const CELL = {
  type: 'string',
  description: 'Cell content. Numbers ("42", "0.15"), dates ("2024-01-31") and formulas ("=SUM(B2:B10)") are parsed as if typed into Excel.',
};
const MATRIX = { type: 'array', items: { type: 'array', items: CELL } };
const SHEET_NAME = { type: 'string', description: 'Exact name of the target worksheet.' };
const RANGE = (example: string) => ({ type: 'string', description: `A1-style range, e.g. "${example}".` });
const STRINGS = (description: string) => ({ type: 'array', items: { type: 'string' }, description });
const COLOR = (description: string) => ({ type: 'string', description: `${description} Hex color, e.g. "#FFC7CE".` });

const tool = (name: ToolName, description: string, properties: Record<string, unknown> = {}, required: string[] = []): ToolDefinition => ({
  type: 'function',
  function: { name, description, parameters: { type: 'object', properties, required } },
});

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  // --- Inspect -------------------------------------------------------------
  tool('get_workbook_context',
    'Inspect the workbook: all worksheets (used ranges, first row, charts, PivotTables), the active sheet with sample rows, the selection, Excel tables with headers, and named ranges.'),
  tool('read_range',
    'Read cell values (and optionally formulas) from a range. Large ranges are truncated to the per-call cell limit (see the system prompt); the result says which part was returned.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A1:F50'), include_formulas: { type: 'boolean', description: 'Also return the formulas.' } },
    ['sheet_name', 'range_address']),
  tool('search_workbook',
    'Find cells whose content contains a text, in one sheet or the whole workbook. Returns cell addresses and values.',
    {
      query: { type: 'string', description: 'Text to find.' },
      sheet_name: { type: 'string', description: 'Limit the search to this sheet (optional).' },
      whole_cell: { type: 'boolean', description: 'Match the entire cell content only.' },
      match_case: { type: 'boolean' },
    },
    ['query']),

  // --- Sheets and names ----------------------------------------------------
  tool('activate_worksheet', 'Switch the visible worksheet tab so the user can see it.', { name: SHEET_NAME }, ['name']),
  tool('create_worksheet', 'Create a new worksheet and activate it. Fails if the name exists.',
    { name: { type: 'string', description: 'Max 31 characters, none of \\ / ? * [ ] :' } }, ['name']),
  tool('rename_worksheet', 'Rename a worksheet.',
    { name: SHEET_NAME, new_name: { type: 'string', description: 'Max 31 characters, none of \\ / ? * [ ] :' } }, ['name', 'new_name']),
  tool('create_named_range', 'Define a workbook-level name for a range, usable in formulas (e.g. =SUM(Sales)).',
    { name: { type: 'string', description: 'Letters, digits, underscores; must not look like a cell address.' }, sheet_name: SHEET_NAME, range_address: RANGE('B2:B100'), comment: { type: 'string' } },
    ['name', 'sheet_name', 'range_address']),

  // --- Data ----------------------------------------------------------------
  tool('write_table',
    'Write new tabular data and make it a native Excel table (header row, table style, auto-fitted columns). Body cells may use formulas, including structured references like "=[@Price]*[@Qty]".',
    {
      sheet_name: SHEET_NAME,
      start_cell: { type: 'string', description: 'Top-left cell of the table, e.g. "A1".' },
      headers: STRINGS('Unique, non-empty column names.'),
      rows: { ...MATRIX, description: 'Data rows; every row must have exactly one value per header.' },
      table_name: { type: 'string', description: 'Optional table name (letters, digits, underscores; no spaces).' },
      table_style: { type: 'string', description: 'Optional built-in style, e.g. "TableStyleMedium2" (default), "TableStyleLight9".' },
    },
    ['sheet_name', 'start_cell', 'headers', 'rows']),
  tool('convert_range_to_table',
    'Turn existing data into a native Excel table. Pass the full data range, or any single cell inside the data to use the surrounding region.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('A1:F200'),
      has_headers: { type: 'boolean', description: 'Whether the first row contains headers (default true).' },
      table_name: { type: 'string' },
      table_style: { type: 'string' },
    },
    ['sheet_name', 'range_address']),
  tool('set_range_values_or_formulas',
    'Write values or formulas into a cell or range. Use English function names and comma separators (e.g. "=XLOOKUP(A2,Data!A:A,Data!B:B)"). If range_address is a single cell, the range expands from it to fit the values. The result reports formula errors such as #NAME? or #REF!.',
    { sheet_name: SHEET_NAME, range_address: RANGE('B2 or B2:D10'), values: { ...MATRIX, description: 'Rectangular 2D array (rows of cells).' } },
    ['sheet_name', 'range_address', 'values']),
  tool('clear_range', 'Clear the contents, the formatting, or both from a range.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A2:F100'), what: { type: 'string', enum: ['contents', 'formats', 'all'], description: 'Default "contents".' } },
    ['sheet_name', 'range_address']),
  tool('sort_range',
    'Sort data by one or more columns. Use table_name for Excel tables, or sheet_name + range_address (a single cell sorts the surrounding data).',
    {
      table_name: { type: 'string' },
      sheet_name: SHEET_NAME,
      range_address: RANGE('A1:F200'),
      sort_by: {
        type: 'array',
        description: 'Sort keys in priority order.',
        items: {
          type: 'object',
          properties: {
            column: { type: 'string', description: 'Header name, or column letter.' },
            ascending: { type: 'boolean', description: 'Default true.' },
          },
          required: ['column'],
        },
      },
      has_headers: { type: 'boolean', description: 'Whether the first row is a header row (default true).' },
    },
    ['sort_by']),
  tool('filter_table',
    'Filter an Excel table column: show only some values, filter by condition, or clear the filter.',
    {
      table_name: { type: 'string' },
      column: { type: 'string', description: 'Header name of the column.' },
      values: STRINGS('Show only rows whose value is in this list.'),
      criteria1: { type: 'string', description: 'Condition such as ">100", "<=2024-12-31", "=Pending", "*north*".' },
      criteria2: { type: 'string', description: 'Second condition, combined with "operator".' },
      operator: { type: 'string', enum: ['And', 'Or'] },
      clear: { type: 'boolean', description: 'Remove the filter from this column.' },
    },
    ['table_name', 'column']),
  tool('add_data_validation',
    'Restrict what can be entered in cells: dropdown lists, number/date/text-length limits, or a custom formula.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('C2:C200'),
      type: { type: 'string', enum: ['list', 'whole_number', 'decimal', 'date', 'text_length', 'custom'] },
      list_values: STRINGS('Dropdown options (for type "list").'),
      list_source: { type: 'string', description: 'Range with the options instead of list_values, e.g. "=Lists!$A$2:$A$10".' },
      operator: { type: 'string', enum: ['Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThan', 'LessThan', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo'] },
      value1: { type: 'string', description: 'Limit or first bound (number, date or formula).' },
      value2: { type: 'string', description: 'Second bound for Between/NotBetween.' },
      formula: { type: 'string', description: 'For type "custom", e.g. "=COUNTIF($A:$A,A2)=1".' },
      error_message: { type: 'string', description: 'Message shown when invalid data is entered.' },
      input_message: { type: 'string', description: 'Hint shown when the cell is selected.' },
      allow_blank: { type: 'boolean', description: 'Default true.' },
    },
    ['sheet_name', 'range_address', 'type']),

  // --- Formatting ------------------------------------------------------------
  tool('format_range', 'Apply visual and number formatting to a range. Only the provided properties change.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('A1:D1'),
      font_bold: { type: 'boolean' },
      font_italic: { type: 'boolean' },
      font_color: COLOR('Text color.'),
      font_size: { type: 'number' },
      fill_color: COLOR('Background color, or "none" to clear it.'),
      num_format: { type: 'string', description: 'Excel number format code, e.g. "$#,##0.00", "0.0%", "yyyy-mm-dd", "@" for text.' },
      horizontal_alignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
      wrap_text: { type: 'boolean' },
      borders: { type: 'string', enum: ['all', 'outline', 'none'], description: 'Thin borders around every cell, only around the range, or remove them.' },
      border_color: COLOR('Border color (default black).'),
      autofit_columns: { type: 'boolean', description: 'Auto-fit the width of the range columns.' },
    },
    ['sheet_name', 'range_address']),
  tool('add_conditional_format',
    'Add conditional formatting: color scales, data bars, icon sets, or highlighting cells by value, text, rank, or a formula.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('D2:D200'),
      type: { type: 'string', enum: ['color_scale', 'data_bar', 'icon_set', 'cell_value', 'text_contains', 'top_bottom', 'formula'] },
      min_color: COLOR('color_scale: color for the lowest value (default red).'),
      mid_color: COLOR('color_scale: optional color for the median.'),
      max_color: COLOR('color_scale: color for the highest value (default green).'),
      bar_color: COLOR('data_bar: bar color.'),
      icon_style: { type: 'string', description: 'icon_set style, e.g. "ThreeTrafficLights1", "ThreeArrows", "ThreeSymbols", "FourRating", "FiveArrows".' },
      operator: { type: 'string', enum: ['GreaterThan', 'LessThan', 'Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThanOrEqual', 'LessThanOrEqual'], description: 'cell_value comparison.' },
      value1: { type: 'string', description: 'cell_value: number or formula to compare with, e.g. "100" or "=$H$1".' },
      value2: { type: 'string', description: 'cell_value: second bound for Between/NotBetween.' },
      text: { type: 'string', description: 'text_contains: text to look for.' },
      rank: { type: 'number', description: 'top_bottom: number of items or percent (default 10).' },
      top_bottom_type: { type: 'string', enum: ['TopItems', 'BottomItems', 'TopPercent', 'BottomPercent'] },
      formula: { type: 'string', description: 'formula: condition relative to the top-left cell, e.g. "=$E2<TODAY()".' },
      fill_color: COLOR('Highlight fill for rule types (default light red).'),
      font_color: COLOR('Highlight text color.'),
      font_bold: { type: 'boolean' },
    },
    ['sheet_name', 'range_address', 'type']),

  // --- Analysis --------------------------------------------------------------
  tool('create_chart',
    'Create a chart from a data range. Include the header row and the category column; each numeric column becomes a series. For summaries, first build the summary table (or PivotTable) and chart it.',
    {
      sheet_name: { type: 'string', description: 'Sheet that contains the data.' },
      data_range: RANGE('A1:C13'),
      chart_type: { type: 'string', enum: [...CHART_TYPES], description: 'Default "ColumnClustered".' },
      title: { type: 'string' },
      series_by: { type: 'string', enum: ['Auto', 'Rows', 'Columns'] },
      destination_sheet: { type: 'string', description: 'Sheet to place the chart on (default: the data sheet).' },
      anchor_cell: { type: 'string', description: 'Top-left cell of the chart (default: right of the data).' },
      width: { type: 'number', description: 'Points (default 480).' },
      height: { type: 'number', description: 'Points (default 288).' },
      x_axis_title: { type: 'string' },
      y_axis_title: { type: 'string' },
      legend_position: { type: 'string', enum: ['Right', 'Bottom', 'Top', 'Left', 'None'] },
      chart_name: { type: 'string' },
    },
    ['sheet_name', 'data_range']),
  tool('create_pivot_table',
    'Create a PivotTable that summarizes data by category. The destination sheet must exist (create it first if needed) and must not overlap the source data.',
    {
      source: { type: 'string', description: 'Table name (e.g. "Sales") or sheet-qualified range including headers (e.g. "Data!A1:F500").' },
      destination_sheet: SHEET_NAME,
      destination_cell: { type: 'string', description: 'Top-left cell (default "A3").' },
      rows: STRINGS('Source column names to group rows by.'),
      columns: STRINGS('Source column names to spread across columns.'),
      values: {
        type: 'array',
        description: 'Fields to aggregate.',
        items: {
          type: 'object',
          properties: {
            field: { type: 'string', description: 'Source column name.' },
            summarize_by: { type: 'string', enum: ['Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers', 'StandardDeviation', 'StandardDeviationP', 'Variance', 'VarianceP'] },
            number_format: { type: 'string', description: 'e.g. "$#,##0".' },
            name: { type: 'string', description: 'Display name; must differ from every source column name.' },
          },
          required: ['field'],
        },
      },
      filters: STRINGS('Source column names to add as report filters.'),
      name: { type: 'string', description: 'Optional PivotTable name.' },
    },
    ['source', 'destination_sheet']),
];

const TOOL_NAMES = new Set(TOOL_DEFINITIONS.map(t => t.function.name));
export const isToolName = (name: string): name is ToolName => TOOL_NAMES.has(name);

const at = (args: Record<string, unknown>, key: string) => `${args.sheet_name ?? '?'}!${args[key] ?? '?'}`;
const count = (value: unknown) => (Array.isArray(value) ? value.length : '?');
const dims = (matrix: unknown) =>
  Array.isArray(matrix) ? `${matrix.length}×${Array.isArray(matrix[0]) ? matrix[0].length : 0}` : '?';
const list = (value: unknown) => (Array.isArray(value) ? value.map(v => (typeof v === 'object' && v ? (v as { field?: unknown; column?: unknown }).field ?? (v as { column?: unknown }).column : v)).join(', ') : '');

/** Short human-readable description shown in the chat and approval cards. */
export const describeToolCall = (name: string, args: Record<string, unknown>): string => {
  switch (name) {
    case 'get_workbook_context': return 'Inspect workbook structure';
    case 'read_range': return `Read ${at(args, 'range_address')}`;
    case 'search_workbook': return `Search for "${args.query}"${args.sheet_name ? ` in ${args.sheet_name}` : ''}`;
    case 'activate_worksheet': return `Switch to sheet "${args.name}"`;
    case 'create_worksheet': return `Create sheet "${args.name}"`;
    case 'rename_worksheet': return `Rename sheet "${args.name}" to "${args.new_name}"`;
    case 'create_named_range': return `Name ${at(args, 'range_address')} as "${args.name}"`;
    case 'write_table': return `Write table (${count(args.rows)} rows × ${count(args.headers)} columns) at ${at(args, 'start_cell')}`;
    case 'convert_range_to_table': return `Convert ${at(args, 'range_address')} to a table`;
    case 'set_range_values_or_formulas': return `Write ${dims(args.values)} cells at ${at(args, 'range_address')}`;
    case 'clear_range': return `Clear ${args.what ?? 'contents'} of ${at(args, 'range_address')}`;
    case 'sort_range': return `Sort ${args.table_name ? `table ${args.table_name}` : at(args, 'range_address')} by ${list(args.sort_by)}`;
    case 'filter_table': return args.clear ? `Clear filter on ${args.table_name}[${args.column}]` : `Filter ${args.table_name}[${args.column}]`;
    case 'add_data_validation': return `Add ${args.type ?? ''} validation to ${at(args, 'range_address')}`;
    case 'format_range': return `Format ${at(args, 'range_address')}`;
    case 'add_conditional_format': return `Add ${String(args.type ?? '').replace(/_/g, ' ')} conditional format to ${at(args, 'range_address')}`;
    case 'create_chart': return `Create ${args.chart_type ?? 'ColumnClustered'} chart from ${at(args, 'data_range')}`;
    case 'create_pivot_table': return `Create PivotTable from ${args.source} on "${args.destination_sheet}" (rows: ${list(args.rows) || '—'}; values: ${list(args.values) || '—'})`;
    default: return `Unknown tool "${name}"`;
  }
};
