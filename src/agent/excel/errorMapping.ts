import { ToolError } from './common';

/**
 * Turns native Office.js exceptions into semantic errors the model can act on:
 * the original message plus a concrete hint about what to check or try instead.
 */
const HINTS: Record<string, string> = {
  InvalidArgument: 'Check the range address and sheet name syntax (A1 style, sheet exists, address within the sheet).',
  ItemNotFound: 'The referenced object does not exist. Check the exact name with get_workbook_context.',
  RangeOutOfBounds: 'The range goes beyond the sheet limits (max 16,384 columns × 1,048,576 rows). Use a smaller range.',
  InvalidOrUnsupportedRange: 'The operation does not support that range shape (whole rows/columns or multiple areas). Use a bounded rectangular range.',
  InvalidSelection: 'The selection is invalid for this operation.',
  AccessDenied: 'The workbook or range is protected (workbook structure, sheet protection or cell locking). Ask the user to remove the protection.',
  ActivityLimitReached: 'The account/license limit was reached. Ask the user to close other workbooks or instances.',
  GeneralException: 'Excel rejected the operation. Simplify the change (fewer cells, one operation at a time) or check that the target is a valid editable range.',
  NotSupported: 'This Excel host does not support the operation.',
  APINotAvailable: 'This Excel version does not support the required API set.',
  PropertyNotLoaded: 'Internal state error; retry the call once. If it repeats, split the work into smaller calls.',
  UnsupportedOperation: 'The operation is not allowed in the current state (e.g. editing part of an array formula, or a protected object).',
  ResourceNotFound: 'The referenced object (sheet, table, chart) was not found. List the existing ones first.',
  InvalidOperation: 'The operation is not valid on that object (e.g. resizing a table row/column of a chart axis). Check the object type.',
  VbaNotAvailable: 'The operation needs VBA support, which is not available in this host.',
  ContentSdlNotSupported: 'The operation is not supported for this content type.',
  CellDataMissing: 'The cell has no data for this operation.',
  FormulaTooLong: 'The formula exceeds the 8,192-character limit. Split it or simplify it.',
  RefreshRequired: 'The object needs a refresh before this operation (e.g. a PivotTable after changing the source).',
};

export const mapOfficeError = (e: unknown): Error => {
  if (e instanceof ToolError) return e;
  if (e instanceof DOMException && e.name === 'AbortError') return e;
  const error = e as { code?: unknown; message?: unknown; name?: unknown };
  const code = typeof error?.code === 'string' ? error.code : '';
  const message = typeof error?.message === 'string' ? error.message : String(e);

  // Failed syncs often surface as RichApi.Error with "The range address is invalid"-style text and no code.
  const hint = HINTS[code] ?? (code ? undefined : guessHintFromMessage(message));
  if (!hint) return e instanceof Error ? e : new Error(message);
  return new ToolError(`Excel error${code ? ` ${code}` : ''}: ${message} — ${hint}`);
};

const guessHintFromMessage = (message: string): string | undefined => {
  if (/cannot be (accessed|modified)|protected/i.test(message)) return 'The sheet or workbook may be protected; ask the user to unprotect it.';
  if (/address.*invalid|invalid.*range|not a valid range/i.test(message)) return HINTS.InvalidArgument;
  if (/out of bounds|too large|exceed/i.test(message)) return HINTS.RangeOutOfBounds;
  if (/table.*not|pivot.*not|chart.*not|sheet.*not|name.*not|wasn't found|not found/i.test(message)) return HINTS.ItemNotFound;
  return undefined;
};
