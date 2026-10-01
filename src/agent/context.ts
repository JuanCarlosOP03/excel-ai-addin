import type { AppSettings } from '../utils/storage';
import { createChatCompletion, type ChatMessage, type ContentPart } from './llmClient';

const CHARS_PER_TOKEN = 3.5;
const BINARY_PART_TOKENS = 1500;
/** Tool outputs older than this many user turns are shortened. */
const RECENT_USER_TURNS = 3;
const MAX_OLD_TOOL_OUTPUT = 2000;
/** User turns kept verbatim when older ones are summarized. */
const KEEP_USER_TURNS = 2;
const MAX_SUMMARY_INPUT_CHARS = 200000;
const SUMMARY_TOOL_OUTPUT_CHARS = 1500;

const contentLength = (content: string | ContentPart[] | null): number => {
  if (!content) return 0;
  if (typeof content === 'string') return content.length;
  return content.reduce((sum, part) => sum + (part.type === 'text' ? part.text.length : BINARY_PART_TOKENS * CHARS_PER_TOKEN), 0);
};

export const estimateTokens = (messages: ChatMessage[]): number =>
  Math.round(messages.reduce((sum, m) => {
    const calls = m.role === 'assistant' ? (m.tool_calls ?? []).reduce((s, c) => s + c.function.arguments.length + c.function.name.length, 0) : 0;
    return sum + contentLength(m.content) + calls + 10;
  }, 0) / CHARS_PER_TOKEN);

const placeholder = (part: ContentPart): ContentPart =>
  part.type === 'image_url' ? { type: 'text', text: '[image attachment omitted]' }
    : part.type === 'file' ? { type: 'text', text: `[file attachment "${part.file.filename}" omitted]` }
    : part;

/** Replaces images and files with placeholders (for storage, and for old turns). */
export const stripBinaryContent = (messages: ChatMessage[]): ChatMessage[] =>
  messages.map(m => (m.role === 'user' && Array.isArray(m.content) && m.content.some(p => p.type !== 'text')
    ? { ...m, content: m.content.map(placeholder) }
    : m));

/** Index of the user message that starts the n-th most recent user turn (or -1). */
const userTurnStart = (messages: ChatMessage[], n: number) => {
  let seen = 0;
  return messages.findLastIndex(m => m.role === 'user' && ++seen === n);
};

/** Cheap reductions: shorten old tool outputs and drop attachments from old turns. */
export const trimOldContent = (messages: ChatMessage[]): ChatMessage[] => {
  const cutoff = userTurnStart(messages, RECENT_USER_TURNS);
  const lastUser = messages.findLastIndex(m => m.role === 'user');
  return messages.map((m, i) => {
    if (i < cutoff && m.role === 'tool' && m.content.length > MAX_OLD_TOOL_OUTPUT) {
      return { ...m, content: `${m.content.slice(0, MAX_OLD_TOOL_OUTPUT)}… (truncated old output)` };
    }
    if (i < lastUser && m.role === 'user' && Array.isArray(m.content)) return stripBinaryContent([m])[0];
    return m;
  });
};

const textOf = (content: string | ContentPart[] | null) =>
  typeof content === 'string' ? content : (content ?? []).map(p => {
    const part = placeholder(p);
    return part.type === 'text' ? part.text : '';
  }).join('\n');

const transcript = (messages: ChatMessage[]): string => {
  const text = messages.map(m => {
    switch (m.role) {
      case 'user': return `USER: ${textOf(m.content)}`;
      case 'assistant': {
        const calls = (m.tool_calls ?? []).map(c => `  [tool call ${c.function.name} ${c.function.arguments.slice(0, 500)}]`).join('\n');
        return `ASSISTANT: ${m.content ?? ''}${calls ? `\n${calls}` : ''}`;
      }
      case 'tool': return `TOOL RESULT: ${m.content.slice(0, SUMMARY_TOOL_OUTPUT_CHARS)}`;
      default: return '';
    }
  }).join('\n\n');
  // Keep the most recent part if the old conversation is enormous.
  return text.length > MAX_SUMMARY_INPUT_CHARS ? text.slice(-MAX_SUMMARY_INPUT_CHARS) : text;
};

const SUMMARY_PROMPT = `You summarize a conversation between a user and an AI agent that edits a Microsoft Excel workbook, so the agent can continue the work with less context.
Write a concise summary (max ~400 words) in the user's language that preserves: the user's goals and preferences, what the agent changed (sheets, tables, ranges, formulas, charts, PivotTables, with exact names and addresses), important facts learned about the data, decisions made, and anything still pending. Do not invent details.`;

export interface FitResult {
  history: ChatMessage[];
  /** How the history was reduced, for an info message in the chat. */
  action?: 'summarized' | 'dropped';
}

/**
 * Keeps the conversation under `budget` tokens: first trims old content, then summarizes
 * older turns with the model (or drops them if summarizing fails).
 */
export const fitHistory = async (
  settings: AppSettings,
  history: ChatMessage[],
  budget: number,
  signal?: AbortSignal
): Promise<FitResult> => {
  const trimmed = trimOldContent(history);
  if (estimateTokens(trimmed) <= budget) return { history: trimmed };

  const cut = userTurnStart(trimmed, KEEP_USER_TURNS);
  if (cut <= 0) return { history: trimmed };
  const older = trimmed.slice(0, cut);
  const recent = trimmed.slice(cut);

  try {
    const { content } = await createChatCompletion(settings, [
      { role: 'system', content: SUMMARY_PROMPT },
      { role: 'user', content: `Conversation to summarize:\n\n${transcript(older)}` },
    ], [], { signal });
    if (!content?.trim()) throw new Error('Empty summary.');
    return {
      action: 'summarized',
      history: [
        { role: 'user', content: `[Summary of the earlier conversation]\n${content.trim()}` },
        { role: 'assistant', content: 'Understood. I will continue from this summary.' },
        ...recent,
      ],
    };
  } catch (e) {
    if (signal?.aborted) throw e;
    console.error('Failed to summarize the conversation', e);
    return { action: 'dropped', history: recent };
  }
};
