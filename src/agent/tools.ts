import type { ToolDefinition } from './llmClient';
import { CHART_TYPES } from './excel/analysisTools';
import { isApiSupported } from './excel/common';

export type ToolName =
  | 'use_skill'
  | 'get_workbook_context' | 'read_range' | 'search_workbook' | 'profile_data'
  | 'trace_formula' | 'find_formula_errors' | 'read_attachment'
  | 'activate_worksheet' | 'create_worksheet' | 'rename_worksheet' | 'delete_worksheet' | 'set_worksheet_visibility'
  | 'create_named_range'
  | 'write_table' | 'convert_range_to_table' | 'import_attachment' | 'set_range_values_or_formulas' | 'clear_range'
  | 'copy_range' | 'fill_range' | 'find_replace' | 'remove_duplicates' | 'insert_range' | 'delete_range'
  | 'sort_range' | 'filter_table' | 'add_data_validation'
  | 'format_range' | 'add_conditional_format' | 'set_rows_columns' | 'freeze_panes' | 'merge_cells' | 'add_comment'
  | 'create_chart' | 'update_chart' | 'delete_chart'
  | 'create_pivot_table' | 'update_pivot_table' | 'delete_pivot_table';

interface ToolMeta {
  /** Changes the workbook: needs approval unless auto-approve is on. */
  mutating: boolean;
  /** Can't be undone: always needs approval, even with auto-approve. */
  irreversible?: boolean;
  /** Minimum ExcelApi requirement set. The manifest requires 1.9. */
  minApi?: string;
}

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
const BOOL = (description?: string) => ({ type: 'boolean', ...(description ? { description } : {}) });
const VALUE_FIELDS = {
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
};

const META: Partial<Record<ToolName, ToolMeta>> = {};

const tool = (
  name: ToolName,
  meta: ToolMeta,
  description: string,
  properties: Record<string, unknown> = {},
  required: string[] = []
): ToolDefinition => {
  META[name] = meta;
  return { type: 'function', function: { name, description, parameters: { type: 'object', properties, required } } };
};

const READ: ToolMeta = { mutating: false };
const WRITE: ToolMeta = { mutating: true };
const IRREVERSIBLE: ToolMeta = { mutating: true, irreversible: true };

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  // --- Skills ------------------------------------------------------------------
  tool('use_skill', READ,
    'Load the expert instructions of a skill (listed in the system prompt) before building something it covers. Load each skill once per conversation.',
    { skill_id: { type: 'string', description: 'Id from the skill list.' } },
    ['skill_id']),

  // --- Inspect and analyze ---------------------------------------------------
  tool('get_workbook_context', READ,
    'Inspect the workbook: all worksheets (used ranges, first row, charts, PivotTables), the active sheet with sample rows, the selection, Excel tables with headers, and named ranges.'),
  tool('read_range', READ,
    'Read cell values (and optionally formulas) from a range. Large ranges are truncated to the per-call cell limit; the result says which part was returned.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A1:F50'), include_formulas: BOOL('Also return the formulas.') },
    ['sheet_name', 'range_address']),
  tool('search_workbook', READ,
    'Find cells whose content contains a text, in one sheet or the whole workbook.',
    { query: { type: 'string' }, sheet_name: { type: 'string', description: 'Limit to this sheet (optional).' }, whole_cell: BOOL('Match the entire cell only.'), match_case: BOOL() },
    ['query']),
  tool('profile_data', READ,
    'Compute statistics for each column of a data set without reading every cell: type, filled/blank counts, distinct values, min/max/mean/median/sum for numbers and dates, top values for text. Use it first to understand large data.',
    { table_name: { type: 'string' }, sheet_name: SHEET_NAME, range_address: RANGE('A1:H5000 (or one cell inside the data)'), has_headers: BOOL('Default true.'), max_rows: { type: 'number' } }),
  tool('trace_formula', { mutating: false, minApi: '1.12' },
    'Explain how a cell is calculated: list the cells its formula depends on (precedents) or the formulas that use it (dependents), with their formulas and values, a few levels deep.',
    {
      sheet_name: SHEET_NAME,
      cell: { type: 'string', description: 'Single cell, e.g. "F20".' },
      direction: { type: 'string', enum: ['precedents', 'dependents'], description: 'Default "precedents".' },
      depth: { type: 'number', description: 'Levels to follow, 1-3 (default 1).' },
    },
    ['sheet_name', 'cell']),
  tool('find_formula_errors', READ,
    'Find every formula that returns an error (#REF!, #N/A, #DIV/0!, #NAME?, #VALUE!…) in a range, a sheet or the whole workbook, with its formula.',
    { sheet_name: { type: 'string', description: 'Optional; whole workbook when omitted.' }, range_address: RANGE('A1:Z500') }),
  tool('read_attachment', READ,
    'Read more rows of a table attachment (or more text of a text attachment, using start_row as the character offset).',
    { attachment_id: { type: 'string' }, start_row: { type: 'number', description: 'First data row (1 = first row after the header).' }, row_count: { type: 'number' } },
    ['attachment_id']),

  // --- Sheets and names --------------------------------------------------------
  tool('activate_worksheet', READ, 'Switch the visible worksheet tab so the user can see it.', { name: SHEET_NAME }, ['name']),
  tool('create_worksheet', WRITE, 'Create a new worksheet and activate it. Fails if the name exists.',
    { name: { type: 'string', description: 'Max 31 characters, none of \\ / ? * [ ] :' } }, ['name']),
  tool('rename_worksheet', WRITE, 'Rename a worksheet.',
    { name: SHEET_NAME, new_name: { type: 'string', description: 'Max 31 characters, none of \\ / ? * [ ] :' } }, ['name', 'new_name']),
  tool('delete_worksheet', WRITE, 'Delete a worksheet (a hidden backup is kept so it can be undone).', { name: SHEET_NAME }, ['name']),
  tool('set_worksheet_visibility', WRITE, 'Hide or unhide a worksheet.', { name: SHEET_NAME, visible: BOOL() }, ['name', 'visible']),
  tool('create_named_range', WRITE, 'Define a workbook-level name for a range, usable in formulas (e.g. =SUM(Sales)).',
    { name: { type: 'string', description: 'Letters, digits, underscores; must not look like a cell address.' }, sheet_name: SHEET_NAME, range_address: RANGE('B2:B100'), comment: { type: 'string' } },
    ['name', 'sheet_name', 'range_address']),

  // --- Data ----------------------------------------------------------------------
  tool('write_table', WRITE,
    'Write new tabular data and make it a native Excel table (header row, style, auto-fitted columns). Body cells may use formulas, including structured references like "=[@Price]*[@Qty]".',
    {
      sheet_name: SHEET_NAME,
      start_cell: { type: 'string', description: 'Top-left cell, e.g. "A1".' },
      headers: STRINGS('Unique, non-empty column names.'),
      rows: { ...MATRIX, description: 'Data rows; one value per header.' },
      table_name: { type: 'string', description: 'Optional (letters, digits, underscores).' },
      table_style: { type: 'string', description: 'e.g. "TableStyleMedium2" (default), "TableStyleLight9".' },
    },
    ['sheet_name', 'start_cell', 'headers', 'rows']),
  tool('convert_range_to_table', WRITE,
    'Turn existing data into a native Excel table. Pass the data range or one cell inside the data.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A1:F200'), has_headers: BOOL('Default true.'), table_name: { type: 'string' }, table_style: { type: 'string' } },
    ['sheet_name', 'range_address']),
  tool('import_attachment', WRITE,
    'Write a table attachment (CSV, TSV, JSON) into the workbook directly, without sending all its rows through the conversation.',
    { attachment_id: { type: 'string' }, sheet_name: SHEET_NAME, start_cell: { type: 'string', description: 'Top-left cell, e.g. "A1".' }, as_table: BOOL('Make it an Excel table (default true).'), table_name: { type: 'string' } },
    ['attachment_id', 'sheet_name', 'start_cell']),
  tool('set_range_values_or_formulas', WRITE,
    'Write values or formulas into a cell or range. Use English function names and comma separators (e.g. "=XLOOKUP(A2,Data!A:A,Data!B:B)"). A single-cell range_address expands to fit the values. The result reports formula errors.',
    { sheet_name: SHEET_NAME, range_address: RANGE('B2 or B2:D10'), values: { ...MATRIX, description: 'Rectangular 2D array (rows of cells).' } },
    ['sheet_name', 'range_address', 'values']),
  tool('clear_range', WRITE, 'Clear the contents, the formatting, or both from a range.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A2:F100'), what: { type: 'string', enum: ['contents', 'formats', 'all'], description: 'Default "contents".' } },
    ['sheet_name', 'range_address']),
  tool('copy_range', WRITE,
    'Copy or move a range to another place (same or another sheet). Can paste only values, formats or formulas, and transpose.',
    {
      source_sheet: SHEET_NAME,
      source_range: RANGE('A1:D20'),
      destination_sheet: { type: 'string', description: 'Default: the source sheet.' },
      destination_cell: { type: 'string', description: 'Top-left cell of the destination.' },
      paste: { type: 'string', enum: ['all', 'values', 'formats', 'formulas'], description: 'Default "all".' },
      transpose: BOOL(),
      move: BOOL('Cut and paste instead of copy.'),
    },
    ['source_sheet', 'source_range', 'destination_cell']),
  tool('fill_range', WRITE,
    'Fill a formula, value or series from source cells into a larger destination range (like dragging the fill handle). The destination must include the source.',
    {
      sheet_name: SHEET_NAME,
      source_range: RANGE('D2'),
      destination_range: RANGE('D2:D500'),
      fill_type: { type: 'string', enum: ['FillDefault', 'FillCopy', 'FillSeries', 'FillFormats', 'FillValues', 'FillDays', 'FillWeekdays', 'FillMonths', 'FillYears', 'LinearTrend', 'GrowthTrend', 'FlashFill'] },
    },
    ['sheet_name', 'source_range', 'destination_range']),
  tool('find_replace', WRITE, 'Replace text in cell contents, in a range, a sheet or the whole workbook.',
    { find: { type: 'string' }, replace: { type: 'string' }, sheet_name: { type: 'string' }, range_address: RANGE('A1:F500'), whole_cell: BOOL(), match_case: BOOL() },
    ['find', 'replace']),
  tool('remove_duplicates', WRITE, 'Remove duplicate rows from a range, comparing all or some columns.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A1:F500 (or one cell inside the data)'), columns: STRINGS('Header names or column letters to compare (default all).'), has_headers: BOOL('Default true.') },
    ['sheet_name', 'range_address']),
  tool('insert_range', WRITE,
    'Insert blank rows ("5:7"), columns ("C:D") or cells (shifting existing cells down or right).',
    { sheet_name: SHEET_NAME, range_address: RANGE('5:7, C:D or B2:C3'), shift: { type: 'string', enum: ['Down', 'Right'], description: 'Only for cells.' } },
    ['sheet_name', 'range_address']),
  tool('delete_range', WRITE,
    'Delete rows ("5:7"), columns ("C:D") or cells (shifting the rest up or left).',
    { sheet_name: SHEET_NAME, range_address: RANGE('5:7, C:D or B2:C3'), shift: { type: 'string', enum: ['Up', 'Left'], description: 'Only for cells.' } },
    ['sheet_name', 'range_address']),
  tool('sort_range', WRITE,
    'Sort data by one or more columns. Use table_name for Excel tables, or sheet_name + range_address (one cell sorts the surrounding data).',
    {
      table_name: { type: 'string' },
      sheet_name: SHEET_NAME,
      range_address: RANGE('A1:F200'),
      sort_by: {
        type: 'array',
        description: 'Sort keys in priority order.',
        items: { type: 'object', properties: { column: { type: 'string', description: 'Header name, or column letter.' }, ascending: BOOL('Default true.') }, required: ['column'] },
      },
      has_headers: BOOL('Default true.'),
    },
    ['sort_by']),
  tool('filter_table', WRITE, 'Filter an Excel table column: show only some values, filter by condition, or clear the filter.',
    {
      table_name: { type: 'string' },
      column: { type: 'string', description: 'Header name.' },
      values: STRINGS('Show only rows whose value is in this list.'),
      criteria1: { type: 'string', description: 'e.g. ">100", "<=2024-12-31", "=Pending", "*north*".' },
      criteria2: { type: 'string' },
      operator: { type: 'string', enum: ['And', 'Or'] },
      clear: BOOL('Remove the filter from this column.'),
    },
    ['table_name', 'column']),
  tool('add_data_validation', WRITE, 'Restrict what can be entered in cells: dropdown lists, number/date/text-length limits, or a custom formula.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('C2:C200'),
      type: { type: 'string', enum: ['list', 'whole_number', 'decimal', 'date', 'text_length', 'custom'] },
      list_values: STRINGS('Dropdown options.'),
      list_source: { type: 'string', description: 'Range with the options, e.g. "=Lists!$A$2:$A$10".' },
      operator: { type: 'string', enum: ['Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThan', 'LessThan', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo'] },
      value1: { type: 'string' },
      value2: { type: 'string' },
      formula: { type: 'string', description: 'For "custom", e.g. "=COUNTIF($A:$A,A2)=1".' },
      error_message: { type: 'string' },
      input_message: { type: 'string' },
      allow_blank: BOOL('Default true.'),
    },
    ['sheet_name', 'range_address', 'type']),

  // --- Formatting and layout ------------------------------------------------------
  tool('format_range', WRITE, 'Apply visual and number formatting to a range. Only the provided properties change.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('A1:D1'),
      font_bold: BOOL(),
      font_italic: BOOL(),
      font_color: COLOR('Text color.'),
      font_size: { type: 'number' },
      fill_color: COLOR('Background, or "none" to clear it.'),
      num_format: { type: 'string', description: 'e.g. "$#,##0.00", "0.0%", "yyyy-mm-dd", "@" for text.' },
      horizontal_alignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
      wrap_text: BOOL(),
      borders: { type: 'string', enum: ['all', 'outline', 'none'] },
      border_color: COLOR('Default black.'),
      autofit_columns: BOOL(),
    },
    ['sheet_name', 'range_address']),
  tool('add_conditional_format', WRITE,
    'Add conditional formatting: color scales, data bars, icon sets, or highlighting cells by value, text, rank, or formula.',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('D2:D200'),
      type: { type: 'string', enum: ['color_scale', 'data_bar', 'icon_set', 'cell_value', 'text_contains', 'top_bottom', 'formula'] },
      min_color: COLOR('color_scale lowest (default red).'),
      mid_color: COLOR('color_scale median (optional).'),
      max_color: COLOR('color_scale highest (default green).'),
      bar_color: COLOR('data_bar.'),
      icon_style: { type: 'string', description: 'e.g. "ThreeTrafficLights1", "ThreeArrows", "FourRating", "FiveArrows".' },
      operator: { type: 'string', enum: ['GreaterThan', 'LessThan', 'Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThanOrEqual', 'LessThanOrEqual'] },
      value1: { type: 'string', description: 'cell_value: number or formula, e.g. "100" or "=$H$1".' },
      value2: { type: 'string' },
      text: { type: 'string', description: 'text_contains.' },
      rank: { type: 'number', description: 'top_bottom (default 10).' },
      top_bottom_type: { type: 'string', enum: ['TopItems', 'BottomItems', 'TopPercent', 'BottomPercent'] },
      formula: { type: 'string', description: 'Relative to the top-left cell, e.g. "=$E2<TODAY()".' },
      fill_color: COLOR('Highlight fill (default light red).'),
      font_color: COLOR('Highlight text.'),
      font_bold: BOOL(),
    },
    ['sheet_name', 'range_address', 'type']),
  tool('set_rows_columns', WRITE,
    'Set column widths / row heights (in points), auto-fit them, or hide/unhide whole rows ("5:7") or columns ("C:D"). For automatic sizing use autofit_rows/autofit_columns and omit the sizes; never pass 0 (it would hide them).',
    {
      sheet_name: SHEET_NAME,
      range_address: RANGE('C:D, 5:7 or A1:F1'),
      hidden: BOOL('Hide (true) or unhide (false).'),
      column_width: { type: 'number', description: 'Points, > 0 (default column ≈ 48, wide text column ≈ 150).' },
      row_height: { type: 'number', description: 'Points, 1–409 (default row ≈ 15).' },
      autofit_columns: BOOL(),
      autofit_rows: BOOL(),
    },
    ['sheet_name', 'range_address']),
  tool('freeze_panes', WRITE, 'Freeze the top rows and/or left columns of a sheet (0 and 0 unfreezes).',
    { sheet_name: SHEET_NAME, rows: { type: 'number' }, columns: { type: 'number' } },
    ['sheet_name']),
  tool('merge_cells', WRITE, 'Merge a range into one cell (or each row with across=true), or unmerge it.',
    { sheet_name: SHEET_NAME, range_address: RANGE('A1:F1'), across: BOOL(), unmerge: BOOL() },
    ['sheet_name', 'range_address']),
  tool('add_comment', { mutating: true, minApi: '1.10' }, 'Add a comment to a cell.',
    { sheet_name: SHEET_NAME, cell: { type: 'string' }, text: { type: 'string' } },
    ['sheet_name', 'cell', 'text']),

  // --- Charts and PivotTables -----------------------------------------------------
  tool('create_chart', WRITE,
    'Create a chart from a range that includes the header row and the category column; each numeric column becomes a series. To chart a summary, build it first.',
    {
      sheet_name: { type: 'string', description: 'Sheet with the data.' },
      data_range: RANGE('A1:C13'),
      chart_type: { type: 'string', enum: [...CHART_TYPES], description: 'Default "ColumnClustered".' },
      title: { type: 'string' },
      series_by: { type: 'string', enum: ['Auto', 'Rows', 'Columns'] },
      destination_sheet: { type: 'string', description: 'Default: the data sheet.' },
      anchor_cell: { type: 'string', description: 'Top-left cell of the chart (default: right of the data).' },
      width: { type: 'number', description: 'Points (default 480).' },
      height: { type: 'number', description: 'Points (default 288).' },
      x_axis_title: { type: 'string' },
      y_axis_title: { type: 'string' },
      legend_position: { type: 'string', enum: ['Right', 'Bottom', 'Top', 'Left', 'None'] },
      chart_name: { type: 'string' },
    },
    ['sheet_name', 'data_range']),
  tool('update_chart', WRITE, 'Change an existing chart: type, data, title, axis titles, legend, position, size or name.',
    {
      sheet_name: SHEET_NAME,
      chart_name: { type: 'string', description: 'Name from get_workbook_context.' },
      chart_type: { type: 'string', enum: [...CHART_TYPES] },
      data_range: RANGE('A1:D13'),
      series_by: { type: 'string', enum: ['Auto', 'Rows', 'Columns'] },
      title: { type: 'string', description: 'Empty string hides the title.' },
      x_axis_title: { type: 'string' },
      y_axis_title: { type: 'string' },
      legend_position: { type: 'string', enum: ['Right', 'Bottom', 'Top', 'Left', 'None'] },
      anchor_cell: { type: 'string' },
      width: { type: 'number' },
      height: { type: 'number' },
      new_name: { type: 'string' },
    },
    ['sheet_name', 'chart_name']),
  tool('delete_chart', IRREVERSIBLE, 'Delete a chart. This cannot be undone.',
    { sheet_name: SHEET_NAME, chart_name: { type: 'string' } }, ['sheet_name', 'chart_name']),
  tool('create_pivot_table', WRITE,
    'Create a PivotTable that summarizes data by category. The destination sheet must exist and must not overlap the source data.',
    {
      source: { type: 'string', description: 'Table name or sheet-qualified range with headers, e.g. "Data!A1:F500".' },
      destination_sheet: SHEET_NAME,
      destination_cell: { type: 'string', description: 'Default "A3".' },
      rows: STRINGS('Columns to group rows by.'),
      columns: STRINGS('Columns to spread across.'),
      values: VALUE_FIELDS,
      filters: STRINGS('Columns to add as report filters.'),
      name: { type: 'string' },
    },
    ['source', 'destination_sheet']),
  tool('update_pivot_table', WRITE, 'Change an existing PivotTable: add or remove row/column/filter/value fields, or refresh it after the source data changed.',
    {
      name: { type: 'string', description: 'PivotTable name from get_workbook_context.' },
      add_rows: STRINGS('Fields to add as rows.'),
      add_columns: STRINGS('Fields to add as columns.'),
      add_filters: STRINGS('Fields to add as filters.'),
      add_values: VALUE_FIELDS,
      remove_fields: STRINGS('Fields (or value names like "Sum of Amount") to remove.'),
      refresh: BOOL(),
    },
    ['name']),
  tool('delete_pivot_table', IRREVERSIBLE, 'Delete a PivotTable. This cannot be undone.', { name: { type: 'string' } }, ['name']),
];

const TOOL_NAMES = new Set(TOOL_DEFINITIONS.map(t => t.function.name));
export const isToolName = (name: string): name is ToolName => TOOL_NAMES.has(name);
export const isMutatingTool = (name: ToolName) => META[name]?.mutating ?? true;
export const isIrreversibleTool = (name: ToolName) => META[name]?.irreversible ?? false;

/** Tools this Excel supports, plus the names of the ones it doesn't. */
export const getAvailableTools = () => {
  const available: ToolDefinition[] = [];
  const unavailable: string[] = [];
  for (const definition of TOOL_DEFINITIONS) {
    const minApi = META[definition.function.name as ToolName]?.minApi;
    if (!minApi || isApiSupported(minApi)) available.push(definition);
    else unavailable.push(definition.function.name);
  }
  return { available, unavailable };
};

const at = (args: Record<string, unknown>, key: string, sheetKey = 'sheet_name') => `${args[sheetKey] ?? '?'}!${args[key] ?? '?'}`;
const count = (value: unknown) => (Array.isArray(value) ? value.length : '?');
const dims = (matrix: unknown) =>
  Array.isArray(matrix) ? `${matrix.length}×${Array.isArray(matrix[0]) ? matrix[0].length : 0}` : '?';
const list = (value: unknown) =>
  Array.isArray(value)
    ? value.map(v => (v && typeof v === 'object' ? (v as { field?: unknown; column?: unknown }).field ?? (v as { column?: unknown }).column : v)).join(', ')
    : '';

/** Short human-readable description shown in the chat and approval cards. */
export const describeToolCall = (name: string, args: Record<string, unknown>): string => {
  switch (name) {
    case 'use_skill': return `Use skill "${args.skill_id}"`;
    case 'get_workbook_context': return 'Inspect workbook structure';
    case 'read_range': return `Read ${at(args, 'range_address')}`;
    case 'search_workbook': return `Search for "${args.query}"${args.sheet_name ? ` in ${args.sheet_name}` : ''}`;
    case 'profile_data': return `Profile data in ${args.table_name ?? at(args, 'range_address')}`;
    case 'trace_formula': return `Trace ${args.direction ?? 'precedents'} of ${at(args, 'cell')}`;
    case 'find_formula_errors': return `Find formula errors in ${args.range_address ? at(args, 'range_address') : args.sheet_name ?? 'the workbook'}`;
    case 'read_attachment': return `Read attachment ${args.attachment_id}`;
    case 'activate_worksheet': return `Switch to sheet "${args.name}"`;
    case 'create_worksheet': return `Create sheet "${args.name}"`;
    case 'rename_worksheet': return `Rename sheet "${args.name}" to "${args.new_name}"`;
    case 'delete_worksheet': return `Delete sheet "${args.name}"`;
    case 'set_worksheet_visibility': return `${args.visible === false || args.visible === 'false' ? 'Hide' : 'Unhide'} sheet "${args.name}"`;
    case 'create_named_range': return `Name ${at(args, 'range_address')} as "${args.name}"`;
    case 'write_table': return `Write table (${count(args.rows)} rows × ${count(args.headers)} columns) at ${at(args, 'start_cell')}`;
    case 'convert_range_to_table': return `Convert ${at(args, 'range_address')} to a table`;
    case 'import_attachment': return `Import attachment ${args.attachment_id} at ${at(args, 'start_cell')}`;
    case 'set_range_values_or_formulas': return `Write ${dims(args.values)} cells at ${at(args, 'range_address')}`;
    case 'clear_range': return `Clear ${args.what ?? 'contents'} of ${at(args, 'range_address')}`;
    case 'copy_range': return `${args.move ? 'Move' : 'Copy'} ${at(args, 'source_range', 'source_sheet')} to ${args.destination_sheet ?? args.source_sheet}!${args.destination_cell}`;
    case 'fill_range': return `Fill ${at(args, 'source_range')} into ${args.destination_range}`;
    case 'find_replace': return `Replace "${args.find}" with "${args.replace}"${args.sheet_name ? ` in ${args.sheet_name}` : ''}`;
    case 'remove_duplicates': return `Remove duplicates in ${at(args, 'range_address')}`;
    case 'insert_range': return `Insert ${at(args, 'range_address')}`;
    case 'delete_range': return `Delete ${at(args, 'range_address')}`;
    case 'sort_range': return `Sort ${args.table_name ? `table ${args.table_name}` : at(args, 'range_address')} by ${list(args.sort_by)}`;
    case 'filter_table': return args.clear ? `Clear filter on ${args.table_name}[${args.column}]` : `Filter ${args.table_name}[${args.column}]`;
    case 'add_data_validation': return `Add ${args.type ?? ''} validation to ${at(args, 'range_address')}`;
    case 'format_range': return `Format ${at(args, 'range_address')}`;
    case 'add_conditional_format': return `Add ${String(args.type ?? '').replace(/_/g, ' ')} conditional format to ${at(args, 'range_address')}`;
    case 'set_rows_columns': return `Adjust rows/columns ${at(args, 'range_address')}`;
    case 'freeze_panes': return `Freeze ${args.rows ?? 0} rows and ${args.columns ?? 0} columns on "${args.sheet_name}"`;
    case 'merge_cells': return `${args.unmerge ? 'Unmerge' : 'Merge'} ${at(args, 'range_address')}`;
    case 'add_comment': return `Comment on ${at(args, 'cell')}`;
    case 'create_chart': return `Create ${args.chart_type ?? 'ColumnClustered'} chart from ${at(args, 'data_range')}`;
    case 'update_chart': return `Update chart "${args.chart_name}"`;
    case 'delete_chart': return `Delete chart "${args.chart_name}"`;
    case 'create_pivot_table': return `Create PivotTable from ${args.source} on "${args.destination_sheet}" (rows: ${list(args.rows) || '—'}; values: ${list(args.values) || '—'})`;
    case 'update_pivot_table': return `Update PivotTable "${args.name}"`;
    case 'delete_pivot_table': return `Delete PivotTable "${args.name}"`;
    default: return `Unknown tool "${name}"`;
  }
};
