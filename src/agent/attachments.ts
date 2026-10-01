import type { ContentPart } from './llmClient';

export type AttachmentCell = string | number | boolean;

export interface Attachment {
  id: string;
  name: string;
  kind: 'table' | 'text' | 'image' | 'pdf';
  /** Table attachments: all rows, the first one being the header row. */
  rows?: AttachmentCell[][];
  text?: string;
  /** Images and PDFs, sent to the model as a data URL. */
  dataUrl?: string;
}

const MAX_FILE_BYTES = 20 * 1024 * 1024;
const MAX_PDF_BYTES = 10 * 1024 * 1024;
const MAX_IMAGE_SIDE = 1568;
const MAX_TEXT_IN_PROMPT = 30000;
const PREVIEW_ROWS = 10;

// Attachments live for the session; the conversation only keeps their ids and summaries.
const registry = new Map<string, Attachment>();
let lastId = 0;

export const getAttachment = (id: string) => registry.get(id);

const register = (attachment: Omit<Attachment, 'id'>): Attachment => {
  const full = { ...attachment, id: `att_${++lastId}` };
  registry.set(full.id, full);
  return full;
};

/** Parses delimited text, handling quoted fields with embedded delimiters, quotes and newlines. */
export const parseDelimited = (text: string, delimiter?: string): string[][] => {
  const source = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text; // Strip the byte order mark.
  const firstLine = source.slice(0, source.indexOf('\n') === -1 ? undefined : source.indexOf('\n'));
  const sep = delimiter ?? [',', ';', '\t', '|'].reduce((best, d) => (firstLine.split(d).length > firstLine.split(best).length ? d : best), ',');

  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < source.length; i++) {
    const ch = source[i];
    if (quoted) {
      if (ch === '"' && source[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"' && field === '') {
      quoted = true;
    } else if (ch === sep) {
      row.push(field); field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && source[i + 1] === '\n') i++;
      row.push(field); field = '';
      rows.push(row); row = [];
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(c => c.trim() !== ''));
};

/** Converts numeric text to numbers so Excel stores real numbers; keeps everything else as text. */
const toCell = (value: string): AttachmentCell => {
  const trimmed = value.trim();
  // Leading zeros (ids, zip codes) stay text.
  if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(trimmed) && trimmed.length < 16) return Number(trimmed);
  return value;
};

const normalizeRows = (rows: unknown[][]): AttachmentCell[][] => {
  const width = Math.max(...rows.map(r => r.length));
  return rows.map(r => Array.from({ length: width }, (_, i) => {
    const v = r[i];
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' || typeof v === 'boolean') return v;
    return typeof v === 'string' ? toCell(v) : JSON.stringify(v);
  }));
};

const jsonToRows = (data: unknown): AttachmentCell[][] | null => {
  if (!Array.isArray(data) || data.length === 0) return null;
  if (data.every(Array.isArray)) return normalizeRows(data as unknown[][]);
  if (!data.every(item => item && typeof item === 'object' && !Array.isArray(item))) return null;
  const headers = [...new Set(data.flatMap(item => Object.keys(item as object)))];
  return normalizeRows([headers, ...data.map(item => headers.map(h => (item as Record<string, unknown>)[h]))]);
};

const readAsDataUrl = (file: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

/** Downscales large images so they fit what vision models accept, keeping the request small. */
const imageDataUrl = async (file: File): Promise<string> => {
  const original = await readAsDataUrl(file);
  const image = new Image();
  image.src = original;
  await image.decode();
  const scale = Math.min(1, MAX_IMAGE_SIDE / Math.max(image.width, image.height));
  if (scale === 1 && file.size < 4 * 1024 * 1024) return original;
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(image.width * scale);
  canvas.height = Math.round(image.height * scale);
  canvas.getContext('2d')?.drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/jpeg', 0.85);
};

export const readAttachment = async (file: File): Promise<Attachment> => {
  if (file.size > MAX_FILE_BYTES) throw new Error(`${file.name} is larger than ${MAX_FILE_BYTES / 1024 / 1024} MB.`);
  const name = file.name || 'pasted-image.png';
  const ext = name.toLowerCase().split('.').pop() ?? '';

  if (file.type.startsWith('image/')) return register({ name, kind: 'image', dataUrl: await imageDataUrl(file) });
  if (ext === 'pdf' || file.type === 'application/pdf') {
    if (file.size > MAX_PDF_BYTES) throw new Error(`PDFs can be at most ${MAX_PDF_BYTES / 1024 / 1024} MB.`);
    return register({ name, kind: 'pdf', dataUrl: await readAsDataUrl(file) });
  }
  if (['xlsx', 'xlsm', 'xls'].includes(ext)) {
    throw new Error('Excel files can\'t be attached. Open the file in Excel and copy the data, or save it as CSV.');
  }

  const text = await file.text();
  if (ext === 'csv' || ext === 'tsv') {
    const rows = parseDelimited(text, ext === 'tsv' ? '\t' : undefined);
    if (rows.length === 0) throw new Error(`${name} is empty.`);
    return register({ name, kind: 'table', rows: normalizeRows(rows) });
  }
  if (ext === 'json') {
    try {
      const rows = jsonToRows(JSON.parse(text));
      if (rows) return register({ name, kind: 'table', rows });
    } catch {
      // Not valid JSON; attached as text.
    }
  }
  return register({ name, kind: 'text', text });
};

/** Text describing an attachment for the model, plus the content part for images and PDFs. */
export const describeAttachment = (attachment: Attachment): { text: string; part?: ContentPart } => {
  switch (attachment.kind) {
    case 'table': {
      const rows = attachment.rows ?? [];
      return {
        text: `Attachment ${attachment.id} "${attachment.name}": table with ${rows.length - 1} data rows × ${rows[0]?.length ?? 0} columns.\n` +
          `First rows (header first): ${JSON.stringify(rows.slice(0, PREVIEW_ROWS + 1))}\n` +
          'Use import_attachment to write it into the workbook, or read_attachment to inspect more rows.',
      };
    }
    case 'text': {
      const text = attachment.text ?? '';
      const truncated = text.length > MAX_TEXT_IN_PROMPT;
      return {
        text: `Attachment ${attachment.id} "${attachment.name}" (text${truncated ? `, first ${MAX_TEXT_IN_PROMPT} of ${text.length} characters` : ''}):\n` +
          `<attachment_content>\n${text.slice(0, MAX_TEXT_IN_PROMPT)}\n</attachment_content>`,
      };
    }
    case 'image':
      return { text: `Attachment ${attachment.id} "${attachment.name}": image (below).`, part: { type: 'image_url', image_url: { url: attachment.dataUrl! } } };
    case 'pdf':
      return { text: `Attachment ${attachment.id} "${attachment.name}": PDF document (below).`, part: { type: 'file', file: { filename: attachment.name, file_data: attachment.dataUrl! } } };
  }
};
