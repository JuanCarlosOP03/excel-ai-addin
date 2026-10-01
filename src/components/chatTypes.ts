import type { ChangeLocation, GridPreview } from '../agent/excel';

export type ToolStatus = 'running' | 'awaiting_approval' | 'ok' | 'error' | 'rejected';

export interface ToolItem {
  kind: 'tool';
  id: string;
  name: string;
  summary: string;
  status: ToolStatus;
  mutating: boolean;
  detail?: string;
  location?: ChangeLocation;
  /** Only while awaiting approval. */
  args?: Record<string, unknown>;
  preview?: GridPreview | null;
  irreversible?: boolean;
}

export interface ChangesItem {
  kind: 'changes';
  id: string;
  entries: { summary: string; location?: ChangeLocation }[];
}

export type ChatItem =
  | { kind: 'user'; id: string; text: string; attachments?: string[]; skills?: string[] }
  | { kind: 'assistant'; id: string; text: string; streaming?: boolean }
  /** The model's reasoning, shown collapsed. */
  | { kind: 'thinking'; id: string; text: string; streaming?: boolean }
  | { kind: 'error' | 'info'; id: string; text: string }
  | ToolItem
  | ChangesItem;
