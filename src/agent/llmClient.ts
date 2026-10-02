import { PROVIDERS, type AppSettings, type ModelInfo } from '../utils/storage';

export interface ToolCall {
  id: string;
  type: 'function';
  function: { name: string; arguments: string };
}

export type ContentPart =
  | { type: 'text'; text: string; cache_control?: { type: 'ephemeral' } }
  | { type: 'image_url'; image_url: { url: string } }
  | { type: 'file'; file: { filename: string; file_data: string } };

export type ChatMessage =
  | { role: 'system'; content: string | ContentPart[] }
  | { role: 'user'; content: string | ContentPart[] }
  /** reasoning_details (OpenRouter) must be sent back unchanged while a tool-calling turn continues. */
  | { role: 'assistant'; content: string | null; tool_calls?: ToolCall[]; reasoning_details?: ReasoningDetail[] }
  | { role: 'tool'; tool_call_id: string; content: string };

export interface ToolDefinition {
  type: 'function';
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

export type ReasoningDetail = Record<string, unknown> & { type?: string; index?: number };

export interface TokenUsage {
  /** Input tokens billed (including the cached ones). */
  promptTokens?: number;
  /** Input tokens served from the provider's prompt cache (billed at a fraction). */
  cachedTokens?: number;
  completionTokens?: number;
}

export interface ChatCompletionResult {
  content: string | null;
  toolCalls: ToolCall[];
  /** Readable reasoning text, when the model exposes it. */
  reasoning?: string;
  reasoningDetails?: ReasoningDetail[];
  /** `length` means the model hit its output limit and the response is truncated. */
  finishReason: string | null;
  usage?: TokenUsage;
}

/** Sums the usage of several steps of one run. */
export const addUsage = (a: TokenUsage | undefined, b: TokenUsage | undefined): TokenUsage => ({
  promptTokens: (a?.promptTokens ?? 0) + (b?.promptTokens ?? 0) || undefined,
  cachedTokens: (a?.cachedTokens ?? 0) + (b?.cachedTokens ?? 0) || undefined,
  completionTokens: (a?.completionTokens ?? 0) + (b?.completionTokens ?? 0) || undefined,
});

const readUsage = (usage?: { prompt_tokens?: number; completion_tokens?: number; prompt_tokens_details?: { cached_tokens?: number }; cache_read_input_tokens?: number }): TokenUsage | undefined => {
  if (!usage) return undefined;
  // OpenAI-compatible providers report prompt_tokens_details.cached_tokens; Anthropic-style
  // responses use cache_read_input_tokens.
  const cachedTokens = usage.prompt_tokens_details?.cached_tokens ?? usage.cache_read_input_tokens;
  if (usage.prompt_tokens === undefined && cachedTokens === undefined && usage.completion_tokens === undefined) return undefined;
  return { promptTokens: usage.prompt_tokens, cachedTokens, completionTokens: usage.completion_tokens };
};

export interface CompletionOptions {
  signal?: AbortSignal;
  /** Receives text as it is generated (streaming only). */
  onTextDelta?: (delta: string) => void;
  /** Receives reasoning ("thinking") text as it is generated. */
  onReasoningDelta?: (delta: string) => void;
  /** Called before waiting to retry a failed request. */
  onRetry?: (info: { attempt: number; delaySeconds: number; reason: string }) => void;
  stream?: boolean;
}

const CORS_PROXY = 'https://corsproxy.io/?';
const MAX_RETRIES = 3;
const MAX_RETRY_DELAY_SECONDS = 30;
const RETRYABLE_STATUS = new Set([408, 425, 429, 500, 502, 503, 504, 529]);

const activeConfig = (settings: AppSettings) => settings.providers[settings.provider];

/** Accepts either a base URL (`.../v1`) or a full `.../chat/completions` URL. */
const resolveBaseUrl = (settings: AppSettings): string => {
  const baseUrl = activeConfig(settings).baseUrl.trim().replace(/\/+$/, '').replace(/\/chat\/completions$/, '');
  if (!baseUrl) throw new Error(`Configure a base URL for ${PROVIDERS[settings.provider].label} in Settings.`);
  return baseUrl;
};

const buildUrl = (settings: AppSettings, path: string): string => {
  const url = `${resolveBaseUrl(settings)}${path}`;
  const useProxy = PROVIDERS[settings.provider].supportsProxy && activeConfig(settings).useProxy;
  return useProxy ? `${CORS_PROXY}${encodeURIComponent(url)}` : url;
};

const buildHeaders = (settings: AppSettings): Record<string, string> => {
  const info = PROVIDERS[settings.provider];
  const { apiKey } = activeConfig(settings);
  if (info.requiresKey && !apiKey) throw new Error(`Add your ${info.label} API key in Settings.`);
  // Provider-specific headers are only sent to that provider: unknown headers make other
  // providers fail the CORS preflight.
  return {
    'Content-Type': 'application/json',
    ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {}),
    ...info.extraHeaders,
  };
};

const sleep = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(timer);
      reject(signal.reason ?? new DOMException('Aborted', 'AbortError'));
    }, { once: true });
  });

const isAbortError = (e: unknown) => e instanceof DOMException && e.name === 'AbortError';

const errorDetail = (text: string, statusText: string): string => {
  try {
    const error = (JSON.parse(text) as { error?: { message?: string; metadata?: { raw?: string } } | string }).error;
    if (typeof error === 'string') return error;
    if (error?.message) return error.metadata?.raw ? `${error.message} (${error.metadata.raw.slice(0, 300)})` : error.message;
  } catch {
    // Not JSON.
  }
  return text.slice(0, 500) || statusText;
};

/** Turns provider errors that mean "this model can't use tools" into an actionable message. */
const describeApiError = (settings: AppSettings, status: number, detail: string): Error => {
  if (/tool|function/i.test(detail) && /support|not available|no endpoints/i.test(detail)) {
    return new Error(`The model "${activeConfig(settings).model}" does not support tool calling, which the agent needs. Choose another model in Settings. (${detail})`);
  }
  return new Error(`API error (${status}): ${detail}`);
};

/** Performs the request, retrying network errors, rate limits and temporary server errors. */
const fetchWithRetry = async (settings: AppSettings, url: string, init: RequestInit, options: CompletionOptions): Promise<Response> => {
  const label = PROVIDERS[settings.provider].label;
  for (let attempt = 0; ; attempt++) {
    let res: Response | null = null;
    let reason: string;
    try {
      res = await fetch(url, init);
      if (res.ok) return res;
      if (!RETRYABLE_STATUS.has(res.status) || attempt >= MAX_RETRIES) {
        throw describeApiError(settings, res.status, errorDetail(await res.text(), res.statusText));
      }
      reason = res.status === 429 ? 'rate limited' : `server error ${res.status}`;
    } catch (e) {
      if (isAbortError(e) || options.signal?.aborted) throw e;
      if (res) throw e; // An API error built above.
      if (attempt >= MAX_RETRIES) {
        const hint = PROVIDERS[settings.provider].supportsProxy && !activeConfig(settings).useProxy
          ? " If the provider doesn't allow browser requests (CORS), you can enable the public CORS proxy in Settings."
          : '';
        throw new Error(`Could not reach ${label}. The server may be offline, the URL may be wrong, or the request was blocked by CORS.${hint}`, { cause: e });
      }
      reason = 'connection failed';
    }

    const retryAfter = Number(res?.headers.get('retry-after'));
    const delaySeconds = Math.min(
      MAX_RETRY_DELAY_SECONDS,
      Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter : 2 ** attempt + Math.random()
    );
    options.onRetry?.({ attempt: attempt + 1, delaySeconds: Math.round(delaySeconds), reason });
    await sleep(delaySeconds * 1000, options.signal);
  }
};

const parseJson = (text: string): unknown => {
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Unexpected non-JSON response from the provider: ${text.slice(0, 200)}`);
  }
};

interface ReasoningFields {
  /** OpenRouter. */
  reasoning?: string | null;
  /** DeepSeek, LM Studio and others. */
  reasoning_content?: string | null;
  reasoning_details?: ReasoningDetail[];
}

interface RawUsage {
  prompt_tokens?: number;
  completion_tokens?: number;
  prompt_tokens_details?: { cached_tokens?: number };
  cache_read_input_tokens?: number;
}

interface CompletionChunk {
  error?: { message?: string } | string;
  choices?: {
    message?: { content?: string | null; tool_calls?: Partial<ToolCall>[] } & ReasoningFields;
    delta?: { content?: string | null; tool_calls?: (Partial<ToolCall> & { index?: number; function?: { name?: string; arguments?: string } })[] } & ReasoningFields;
    finish_reason?: string | null;
  }[];
  usage?: RawUsage;
}

const reasoningText = (fields?: ReasoningFields) => fields?.reasoning || fields?.reasoning_content || '';

/** Streamed reasoning_details arrive in pieces; pieces with the same index are concatenated. */
const mergeReasoningDetails = (target: ReasoningDetail[], pieces: ReasoningDetail[]) => {
  for (const piece of pieces) {
    const index = typeof piece.index === 'number' ? piece.index : target.length ? target.length - 1 : 0;
    const existing = target.find(d => (d.index ?? 0) === index && d.type === piece.type);
    if (!existing) {
      target.push({ ...piece, index });
      continue;
    }
    for (const [key, value] of Object.entries(piece)) {
      if (['text', 'summary', 'data'].includes(key) && typeof value === 'string') existing[key] = `${existing[key] ?? ''}${value}`;
      else if (value !== null && value !== undefined) existing[key] = value;
    }
  }
};

const throwChunkError = (chunk: CompletionChunk) => {
  if (!chunk.error) return;
  const message = typeof chunk.error === 'string' ? chunk.error : chunk.error.message ?? JSON.stringify(chunk.error);
  throw new Error(`API error: ${message}`);
};

const normalizeToolCalls = (calls: Partial<ToolCall>[]): ToolCall[] =>
  calls.map((call, i) => ({
    id: call.id || `call_${Date.now()}_${i}`,
    type: 'function',
    function: { name: call.function?.name ?? '', arguments: call.function?.arguments || '{}' },
  }));

/** Reads an OpenAI-style server-sent event stream, accumulating text and tool call fragments. */
export const readCompletionStream = async (
  body: ReadableStream<Uint8Array>,
  onTextDelta?: (delta: string) => void,
  onReasoningDelta?: (delta: string) => void
): Promise<ChatCompletionResult> => {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let content = '';
  let reasoning = '';
  let usage: TokenUsage | undefined;
  const reasoningDetails: ReasoningDetail[] = [];
  let finishReason: string | null = null;
  const calls: { id?: string; name: string; arguments: string }[] = [];

  const fragmentIndex = (fragment: { index?: number; id?: string }) => {
    if (typeof fragment.index === 'number') return fragment.index;
    // Some providers omit the index and send each call whole, or continue the last one.
    if (fragment.id) {
      const existing = calls.findIndex(c => c?.id === fragment.id);
      return existing === -1 ? calls.length : existing;
    }
    return Math.max(calls.length - 1, 0);
  };

  const handle = (payload: string) => {
    if (payload === '[DONE]') return;
    const chunk = parseJson(payload) as CompletionChunk;
    throwChunkError(chunk);
    const reported = readUsage(chunk.usage);
    // Usage arrives in its own final chunk when streaming.
    if (reported) usage = reported;
    const choice = chunk.choices?.[0];
    if (!choice) return;
    if (choice.finish_reason) finishReason = choice.finish_reason;
    const delta = choice.delta ?? choice.message;
    const thinking = reasoningText(delta);
    if (thinking) {
      reasoning += thinking;
      onReasoningDelta?.(thinking);
    }
    if (delta?.reasoning_details?.length) mergeReasoningDetails(reasoningDetails, delta.reasoning_details);
    if (delta?.content) {
      content += delta.content;
      onTextDelta?.(delta.content);
    }
    for (const fragment of delta?.tool_calls ?? []) {
      const call = (calls[fragmentIndex(fragment)] ??= { name: '', arguments: '' });
      if (fragment.id) call.id = fragment.id;
      if (fragment.function?.name) call.name += fragment.function.name;
      if (fragment.function?.arguments) call.arguments += fragment.function.arguments;
    }
  };

  for (;;) {
    const { value, done } = await reader.read();
    buffer += decoder.decode(value, { stream: !done });
    const lines = buffer.split(/\r?\n/);
    buffer = done ? '' : lines.pop() ?? '';
    for (const line of lines) {
      // Lines starting with ":" are comments (keep-alives such as ": OPENROUTER PROCESSING").
      if (line.startsWith('data:')) handle(line.slice(5).trim());
    }
    if (done) break;
  }

  return {
    content: content || null,
    toolCalls: normalizeToolCalls(calls.filter(Boolean).map(c => ({ id: c.id, function: { name: c.name, arguments: c.arguments } }))),
    finishReason,
    ...(usage ? { usage } : {}),
    ...(reasoning ? { reasoning } : {}),
    ...(reasoningDetails.length ? { reasoningDetails } : {}),
  };
};

export const withoutReasoningDetails = (message: ChatMessage): ChatMessage => {
  if (message.role !== 'assistant' || !message.reasoning_details) return message;
  const copy = { ...message };
  delete copy.reasoning_details;
  return copy;
};

/** Provider-specific request fields for the chosen reasoning effort. */
const reasoningParams = (settings: AppSettings): Record<string, unknown> => {
  const effort = settings.reasoningEffort;
  if (effort === 'default') return {};
  if (activeConfig(settings).modelInfo?.supportsReasoning === false) return {};
  // OpenRouter has a unified `reasoning` object; other OpenAI-compatible APIs use `reasoning_effort`.
  return settings.provider === 'openrouter' ? { reasoning: { effort } } : { reasoning_effort: effort };
};

const usesPromptCaching = (settings: AppSettings) =>
  settings.provider === 'openrouter' && activeConfig(settings).model.startsWith('anthropic/');

const withCacheBreakpoint = (message: ChatMessage): ChatMessage => {
  if (message.role !== 'system' && message.role !== 'user') return message;
  const parts: ContentPart[] = typeof message.content === 'string' ? [{ type: 'text', text: message.content }] : [...message.content];
  const last = parts.findLastIndex(p => p.type === 'text');
  if (last === -1) return message;
  parts[last] = { ...(parts[last] as Extract<ContentPart, { type: 'text' }>), cache_control: { type: 'ephemeral' } };
  return { ...message, content: parts };
};

/**
 * Anthropic models need explicit cache breakpoints (other providers cache stable prefixes
 * automatically). Breakpoints are placed on the system prompt and on the two most recent user
 * messages: the last one covers the current turn, and the previous one lets the next request
 * reuse everything up to the end of the previous turn. Up to four are allowed per request, and
 * the tool definitions are cached as part of the prefix.
 */
const applyPromptCaching = (settings: AppSettings, messages: ChatMessage[]): ChatMessage[] => {
  if (!usesPromptCaching(settings)) return messages;
  const lastUser = messages.findLastIndex(m => m.role === 'user');
  const previousUser = messages.findLastIndex((m, i) => m.role === 'user' && i < lastUser);
  const breakpoints = new Set([lastUser, previousUser].filter(i => i !== -1));
  return messages.map((m, i) => (m.role === 'system' || breakpoints.has(i) ? withCacheBreakpoint(m) : m));
};

export const createChatCompletion = async (
  settings: AppSettings,
  messages: ChatMessage[],
  tools: ToolDefinition[],
  options: CompletionOptions = {}
): Promise<ChatCompletionResult> => {
  const model = activeConfig(settings).model.trim();
  if (!model) throw new Error('Select a model in Settings.');
  const stream = options.stream ?? false;
  const reasoning = reasoningParams(settings);
  // reasoning_details are only understood by OpenRouter; other providers may reject unknown fields.
  const outgoing = settings.provider === 'openrouter'
    ? messages
    : messages.map(withoutReasoningDetails);

  const res = await fetchWithRetry(settings, buildUrl(settings, '/chat/completions'), {
    method: 'POST',
    headers: buildHeaders(settings),
    body: JSON.stringify({
      model,
      messages: applyPromptCaching(settings, outgoing),
      ...(tools.length ? { tools, tool_choice: 'auto' } : {}),
      // Models that think with a reasoning budget (e.g. Claude) don't accept a custom temperature.
      ...(Object.keys(reasoning).length && settings.reasoningEffort !== 'none' ? {} : { temperature: 0.2 }),
      ...reasoning,
      // OpenRouter always reports usage, but streaming needs it requested explicitly.
      ...(stream ? { stream: true, ...(settings.provider === 'openrouter' ? { stream_options: { include_usage: true } } : {}) } : {}),
    }),
    signal: options.signal,
  }, options);

  // Some providers ignore `stream` and answer with plain JSON.
  if (stream && res.body && (res.headers.get('content-type') ?? '').includes('text/event-stream')) {
    return readCompletionStream(res.body, options.onTextDelta, options.onReasoningDelta);
  }

  const data = parseJson(await res.text()) as CompletionChunk;
  throwChunkError(data);
  const choice = data.choices?.[0];
  if (!choice?.message) throw new Error(`The provider returned no message: ${JSON.stringify(data).slice(0, 300)}`);
  const thinking = reasoningText(choice.message);
  if (thinking) options.onReasoningDelta?.(thinking);
  if (choice.message.content) options.onTextDelta?.(choice.message.content);
  return {
    content: choice.message.content ?? null,
    toolCalls: normalizeToolCalls(choice.message.tool_calls ?? []),
    finishReason: choice.finish_reason ?? null,
    ...(readUsage(data.usage) ? { usage: readUsage(data.usage) } : {}),
    ...(thinking ? { reasoning: thinking } : {}),
    ...(choice.message.reasoning_details?.length ? { reasoningDetails: choice.message.reasoning_details } : {}),
  };
};

interface RawModel {
  id?: string;
  context_length?: number;
  supported_parameters?: string[];
  architecture?: { input_modalities?: string[] };
}

export const listModels = async (settings: AppSettings): Promise<ModelInfo[]> => {
  const headers = buildHeaders(settings);
  delete headers['Content-Type'];
  const res = await fetchWithRetry(settings, buildUrl(settings, '/models'), { headers }, {});
  const data = parseJson(await res.text()) as { data?: RawModel[] };
  if (!Array.isArray(data.data)) throw new Error('Unexpected response format when listing models.');
  return data.data
    .filter(m => m.id)
    .map(m => ({
      id: m.id!.replace(/^models\//, ''),
      // OpenRouter reports capabilities; other providers leave them unknown.
      ...(m.supported_parameters ? {
        supportsTools: m.supported_parameters.includes('tools'),
        supportsReasoning: m.supported_parameters.includes('reasoning') || m.supported_parameters.includes('include_reasoning'),
      } : {}),
      ...(m.architecture?.input_modalities ? {
        supportsImages: m.architecture.input_modalities.includes('image'),
        supportsFiles: m.architecture.input_modalities.includes('file'),
      } : {}),
      ...(m.context_length ? { contextLength: m.context_length } : {}),
    }))
    .sort((a, b) => a.id.localeCompare(b.id));
};
