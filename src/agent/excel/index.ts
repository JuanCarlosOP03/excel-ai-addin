import type { ToolName } from '../tools';
import type { Args, Executor, ToolOptions } from './common';
import { createChart, createPivotTable } from './analysisTools';
import {
  addDataValidation,
  clearRange,
  convertRangeToTable,
  filterTable,
  setRangeValuesOrFormulas,
  sortRange,
  writeTable,
} from './dataTools';
import { addConditionalFormat, formatRange } from './formatTools';
import {
  activateWorksheet,
  createNamedRange,
  createWorksheet,
  getWorkbookContext,
  readRange,
  renameWorksheet,
  searchWorkbook,
} from './workbookTools';

export { ToolError, type ToolOptions } from './common';
export { CHART_TYPES } from './analysisTools';
export { beginUndoGroup, canUndo, endUndoGroup, undoLastGroup } from './undo';
export { getSelectedRangeAddress, getWorkbookOverview, type WorkbookOverview } from './workbookTools';

const EXECUTORS: Record<ToolName, Executor> = {
  get_workbook_context: getWorkbookContext,
  read_range: readRange,
  search_workbook: searchWorkbook,
  activate_worksheet: activateWorksheet,
  create_worksheet: createWorksheet,
  rename_worksheet: renameWorksheet,
  create_named_range: createNamedRange,
  write_table: writeTable,
  convert_range_to_table: convertRangeToTable,
  set_range_values_or_formulas: setRangeValuesOrFormulas,
  clear_range: clearRange,
  sort_range: sortRange,
  filter_table: filterTable,
  add_data_validation: addDataValidation,
  format_range: formatRange,
  add_conditional_format: addConditionalFormat,
  create_chart: createChart,
  create_pivot_table: createPivotTable,
};

/** Runs one tool in its own Excel batch. Throws ToolError or Office errors on failure. */
export const executeTool = (name: ToolName, args: Args, options: ToolOptions = {}): Promise<unknown> =>
  Excel.run(context => EXECUTORS[name](context, args, options));
