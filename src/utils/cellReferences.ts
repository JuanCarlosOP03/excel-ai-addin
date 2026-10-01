export const CELL_LINK_PREFIX = '#xlref=';

const SHEET = String.raw`(?:'(?:[^']|'')+'|[\p{L}\p{N}_.]+)`;
const CELL = String.raw`\$?[A-Z]{1,3}\$?\d{1,7}`;
const COLUMN = String.raw`\$?[A-Z]{1,3}`;
const ROW = String.raw`\$?\d{1,7}`;
/**
 * Sheet-qualified references (Sales!B5, 'Q1 Data'!A1:D9, Data!C:C, Data!2:5) and unqualified
 * ranges with a colon (A1:C10). Single unqualified cells are not linked: "Q4" or "B2B" would
 * become false positives.
 */
const REFERENCE = new RegExp(
  String.raw`(?<![\p{L}\p{N}_.!'$])(${SHEET}!(?:${CELL}(?::${CELL})?|${COLUMN}:${COLUMN}|${ROW}:${ROW})|${CELL}:${CELL})(?![\p{L}\p{N}_(])`,
  'gu'
);

/** Turns cell references outside code spans/blocks and existing links into Markdown links. */
export const linkifyCellReferences = (markdown: string): string =>
  markdown
    .split(/(```[\s\S]*?```|`[^`\n]*`|\[[^\]\n]*\]\([^)\n]*\))/)
    .map((segment, i) => (i % 2 === 1 ? segment : segment.replace(REFERENCE, ref => `[${ref}](${CELL_LINK_PREFIX}${encodeURIComponent(ref)})`)))
    .join('');

/** "Sheet 1" → "'Sheet 1'"; plain names stay as they are. */
export const quoteSheetName = (name: string) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
