import type { ToolDefinition } from './llmClient';

export type ToolName =
  | 'get_workbook_context'
  | 'read_range'
  | 'activate_worksheet'
  | 'create_worksheet'
  | 'write_table'
  | 'set_range_values_or_formulas'
  | 'format_range';

/** Tools that change workbook content. They require user approval unless auto-approve is on, and are undoable. */
export const MUTATING_TOOLS: ReadonlySet<ToolName> = new Set<ToolName>([
  'create_worksheet',
  'write_table',
  'set_range_values_or_formulas',
  'format_range',
]);

// Cells are typed as strings for maximum provider compatibility (some reject union types).
// The executors also accept numbers, booleans and null.
const CELL = {
  type: 'string',
  description: 'Cell content. Numbers ("42", "0.15"), dates ("2024-01-31") and formulas ("=SUM(B2:B10)") are parsed as if typed into Excel.',
};
const MATRIX = { type: 'array', items: { type: 'array', items: CELL } };
const SHEET_NAME = { type: 'string', description: 'Exact name of the target worksheet.' };

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'get_workbook_context',
      description: 'Inspect the workbook structure: all worksheet names, the active sheet, the current selection, used ranges with sizes and header rows, existing Excel tables, and a few sample rows of the active sheet.',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_range',
      description: 'Read cell values (and optionally formulas) from a range. Large ranges are truncated to the per-call cell limit (see the system prompt); the result says so and which part was returned.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: SHEET_NAME,
          range_address: { type: 'string', description: 'A1-style range, e.g. "A1:F50" or "C:C".' },
          include_formulas: { type: 'boolean', description: 'Also return the formulas of the cells.' },
        },
        required: ['sheet_name', 'range_address'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'activate_worksheet',
      description: 'Switch the visible worksheet tab so the user can see it.',
      parameters: { type: 'object', properties: { name: SHEET_NAME }, required: ['name'] },
    },
  },
  {
    type: 'function',
    function: {
      name: 'create_worksheet',
      description: 'Create a new worksheet and activate it. Fails if a sheet with that name already exists.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'New sheet name: max 31 characters, none of \\ / ? * [ ] :' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'write_table',
      description: 'Write tabular data and convert it into a native Excel table (with header row, table style and auto-fitted columns). Body cells may contain formulas, including structured references like "=[@Price]*[@Qty]".',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: SHEET_NAME,
          start_cell: { type: 'string', description: 'Top-left cell of the table, e.g. "A1".' },
          headers: { type: 'array', items: { type: 'string' }, description: 'Unique, non-empty column names.' },
          rows: { ...MATRIX, description: 'Data rows; every row must have exactly one value per header.' },
          table_name: { type: 'string', description: 'Optional table name (letters, digits, underscores; no spaces).' },
          table_style: { type: 'string', description: 'Optional built-in style, e.g. "TableStyleMedium2" (default).' },
        },
        required: ['sheet_name', 'start_cell', 'headers', 'rows'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_range_values_or_formulas',
      description: 'Write values or formulas into a cell or range. Use English function names and comma separators (e.g. "=XLOOKUP(A2,Data!A:A,Data!B:B)"). If range_address is a single cell, the range is expanded from it to fit the values. The result reports formula errors such as #NAME? or #REF!.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: SHEET_NAME,
          range_address: { type: 'string', description: 'A1-style range or top-left cell, e.g. "B2" or "B2:D10".' },
          values: { ...MATRIX, description: 'Rectangular 2D array (rows of cells).' },
        },
        required: ['sheet_name', 'range_address', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'format_range',
      description: 'Apply visual and number formatting to a range. Only the provided properties are changed.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: SHEET_NAME,
          range_address: { type: 'string', description: 'A1-style range, e.g. "A1:D1".' },
          font_bold: { type: 'boolean' },
          font_italic: { type: 'boolean' },
          font_color: { type: 'string', description: 'Hex color, e.g. "#FFFFFF".' },
          fill_color: { type: 'string', description: 'Hex background color, e.g. "#FFFF00", or "none" to clear it.' },
          num_format: { type: 'string', description: 'Excel number format code, e.g. "$#,##0.00", "0.0%", "yyyy-mm-dd", "@" for text.' },
          horizontal_alignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
          autofit_columns: { type: 'boolean', description: 'Auto-fit the width of the range columns.' },
        },
        required: ['sheet_name', 'range_address'],
      },
    },
  },
];

const TOOL_NAMES = new Set(TOOL_DEFINITIONS.map(t => t.function.name));
export const isToolName = (name: string): name is ToolName => TOOL_NAMES.has(name);

const at = (args: Record<string, unknown>, key: 'range_address' | 'start_cell') =>
  `${args.sheet_name ?? '?'}!${args[key] ?? '?'}`;

const dims = (matrix: unknown) =>
  Array.isArray(matrix) ? `${matrix.length}×${Array.isArray(matrix[0]) ? matrix[0].length : 0}` : '?';

/** Short human-readable description shown in the chat and approval cards. */
export const describeToolCall = (name: string, args: Record<string, unknown>): string => {
  switch (name) {
    case 'get_workbook_context': return 'Inspect workbook structure';
    case 'read_range': return `Read ${at(args, 'range_address')}`;
    case 'activate_worksheet': return `Switch to sheet "${args.name}"`;
    case 'create_worksheet': return `Create sheet "${args.name}"`;
    case 'write_table': {
      const cols = Array.isArray(args.headers) ? args.headers.length : '?';
      const rows = Array.isArray(args.rows) ? args.rows.length : '?';
      return `Write table (${rows} rows × ${cols} columns) at ${at(args, 'start_cell')}`;
    }
    case 'set_range_values_or_formulas': return `Write ${dims(args.values)} cells at ${at(args, 'range_address')}`;
    case 'format_range': return `Format ${at(args, 'range_address')}`;
    default: return `Unknown tool "${name}"`;
  }
};
