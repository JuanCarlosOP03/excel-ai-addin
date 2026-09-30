import { PROVIDERS, type AppSettings } from '../utils/storage';

export interface ToolCall {
  id: string;
  type: 'function';
  function: { name: string; arguments: string };
}

export type ChatMessage =
  | { role: 'system'; content: string }
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string | null; tool_calls?: ToolCall[] }
  | { role: 'tool'; tool_call_id: string; content: string };

export interface ToolDefinition {
  type: 'function';
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

export interface ChatCompletionResult {
  content: string | null;
  toolCalls: ToolCall[];
  /** `length` means the model hit its output limit and the response is truncated. */
  finishReason: string | null;
}

const CORS_PROXY = 'https://corsproxy.io/?';

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

const request = async (settings: AppSettings, url: string, init: RequestInit): Promise<unknown> => {
  let res: Response;
  try {
    res = await fetch(url, init);
  } catch (e) {
    if (e instanceof DOMException && e.name === 'AbortError') throw e;
    const hint = PROVIDERS[settings.provider].supportsProxy && !activeConfig(settings).useProxy
      ? " If the provider doesn't allow browser requests (CORS), you can enable the public CORS proxy in Settings."
      : '';
    throw new Error(`Could not reach ${PROVIDERS[settings.provider].label}. The server may be offline, the URL may be wrong, or the request was blocked by CORS.${hint}`, { cause: e });
  }

  const text = await res.text();
  let data: unknown = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    // Non-JSON body; reported below.
  }

  const apiError = (data as { error?: { message?: string } | string } | null)?.error;
  if (!res.ok || apiError) {
    const detail = typeof apiError === 'string' ? apiError : apiError?.message || text.slice(0, 500) || res.statusText;
    throw new Error(`API error (${res.status}): ${detail}`);
  }
  if (data === null) throw new Error(`Unexpected non-JSON response from the provider: ${text.slice(0, 200)}`);
  return data;
};

export const createChatCompletion = async (
  settings: AppSettings,
  messages: ChatMessage[],
  tools: ToolDefinition[],
  signal?: AbortSignal
): Promise<ChatCompletionResult> => {
  const model = activeConfig(settings).model.trim();
  if (!model) throw new Error('Select a model in Settings.');

  const data = await request(settings, buildUrl(settings, '/chat/completions'), {
    method: 'POST',
    headers: buildHeaders(settings),
    body: JSON.stringify({ model, messages, tools, tool_choice: 'auto', temperature: 0.2 }),
    signal,
  }) as { choices?: { message?: { content?: string | null; tool_calls?: ToolCall[] }; finish_reason?: string | null }[] };

  const choice = data.choices?.[0];
  if (!choice?.message) throw new Error(`The provider returned no message: ${JSON.stringify(data).slice(0, 300)}`);

  return {
    content: choice.message.content ?? null,
    toolCalls: (choice.message.tool_calls ?? []).map((call, i) => ({
      id: call.id || `call_${Date.now()}_${i}`,
      type: 'function',
      function: { name: call.function?.name ?? '', arguments: call.function?.arguments || '{}' },
    })),
    finishReason: choice.finish_reason ?? null,
  };
};

export const listModels = async (settings: AppSettings): Promise<string[]> => {
  const headers = buildHeaders(settings);
  delete headers['Content-Type'];
  const data = await request(settings, buildUrl(settings, '/models'), { headers }) as { data?: { id?: string }[] };
  if (!Array.isArray(data.data)) throw new Error('Unexpected response format when listing models.');
  return data.data
    .map(m => (m.id ?? '').replace(/^models\//, ''))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
};
