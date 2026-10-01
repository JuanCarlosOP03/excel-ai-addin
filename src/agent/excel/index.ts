import type { ToolName } from '../tools';
import type { Args, Executor, ToolOptions } from './common';
import { createChart, createPivotTable, deleteChart, deletePivotTable, updateChart, updatePivotTable } from './analysisTools';
import {
  addDataValidation,
  clearRange,
  convertRangeToTable,
  filterTable,
  setRangeValuesOrFormulas,
  sortRange,
  writeTable,
} from './dataTools';
import { copyRange, fillRange, findReplace, removeDuplicates } from './editTools';
import { addConditionalFormat, formatRange } from './formatTools';
import { findFormulaErrors, traceFormula } from './formulaTools';
import { importAttachment, profileData, readAttachmentTool } from './insightTools';
import {
  addComment,
  deleteRange,
  deleteWorksheet,
  freezePanes,
  insertRange,
  mergeCells,
  setRowsColumns,
  setWorksheetVisibility,
} from './structureTools';
import {
  activateWorksheet,
  createNamedRange,
  createWorksheet,
  getWorkbookContext,
  readRange,
  renameWorksheet,
  searchWorkbook,
} from './workbookTools';

export { ToolError, isApiSupported, type ChangeLocation, type ToolOptions } from './common';
export { CHART_TYPES } from './analysisTools';
export { previewToolCall, selectReference, type GridPreview } from './preview';
export { beginUndoGroup, canUndo, cleanupOrphanBackups, endUndoGroup, initUndo, undoLastGroup, type UndoStep } from './undo';
export { getSelectedRangeAddress, getWorkbookOverview, type WorkbookOverview } from './workbookTools';

// use_skill is handled by the agent loop: it doesn't touch the workbook.
export type ExcelToolName = Exclude<ToolName, 'use_skill'>;

const EXECUTORS: Record<ExcelToolName, Executor> = {
  get_workbook_context: getWorkbookContext,
  read_range: readRange,
  search_workbook: searchWorkbook,
  profile_data: profileData,
  trace_formula: traceFormula,
  find_formula_errors: findFormulaErrors,
  read_attachment: readAttachmentTool,
  activate_worksheet: activateWorksheet,
  create_worksheet: createWorksheet,
  rename_worksheet: renameWorksheet,
  delete_worksheet: deleteWorksheet,
  set_worksheet_visibility: setWorksheetVisibility,
  create_named_range: createNamedRange,
  write_table: writeTable,
  convert_range_to_table: convertRangeToTable,
  import_attachment: importAttachment,
  set_range_values_or_formulas: setRangeValuesOrFormulas,
  clear_range: clearRange,
  copy_range: copyRange,
  fill_range: fillRange,
  find_replace: findReplace,
  remove_duplicates: removeDuplicates,
  insert_range: insertRange,
  delete_range: deleteRange,
  sort_range: sortRange,
  filter_table: filterTable,
  add_data_validation: addDataValidation,
  format_range: formatRange,
  add_conditional_format: addConditionalFormat,
  set_rows_columns: setRowsColumns,
  freeze_panes: freezePanes,
  merge_cells: mergeCells,
  add_comment: addComment,
  create_chart: createChart,
  update_chart: updateChart,
  delete_chart: deleteChart,
  create_pivot_table: createPivotTable,
  update_pivot_table: updatePivotTable,
  delete_pivot_table: deletePivotTable,
};

/** Runs one tool in its own Excel batch. Throws ToolError or Office errors on failure. */
export const executeTool = (name: ExcelToolName, args: Args, options: ToolOptions = {}): Promise<unknown> =>
  Excel.run(context => EXECUTORS[name](context, args, options));
