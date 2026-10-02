export type ProviderId = 'openrouter' | 'gemini' | 'lmstudio' | 'opencode' | 'custom';

import type { EncryptedBox } from './secretBox';
import { decryptJson, encryptJson } from './secretBox';

/** Capabilities reported by the provider's model list (unknown when undefined). */
export interface ModelInfo {
  id: string;
  supportsTools?: boolean;
  supportsImages?: boolean;
  supportsFiles?: boolean;
  supportsReasoning?: boolean;
  contextLength?: number;
}

/** "default" sends nothing and lets the model decide; "none" turns reasoning off. */
export type ReasoningEffort = 'default' | 'none' | 'low' | 'medium' | 'high';
export const REASONING_EFFORTS: { value: ReasoningEffort; label: string }[] = [
  { value: 'default', label: 'Default' },
  { value: 'none', label: 'Off' },
  { value: 'low', label: 'Low' },
  { value: 'medium', label: 'Medium' },
  { value: 'high', label: 'High' },
];

export interface FavoriteModel {
  provider: ProviderId;
  model: string;
  /** Short name shown in the model picker. */
  label: string;
  modelInfo?: ModelInfo;
}

export const MAX_FAVORITES = 5;

export interface CustomSkill {
  id: string;
  name: string;
  /** When the agent should use it (shown to the model in the skill list). */
  description: string;
  instructions: string;
}

export const MAX_CUSTOM_SKILLS = 20;
export const MAX_SKILL_INSTRUCTIONS = 8000;

export interface ProviderConfig {
  baseUrl: string;
  apiKey: string;
  model: string;
  useProxy: boolean;
  /** Capabilities of `model`, when it was picked from a loaded model list. */
  modelInfo?: ModelInfo;
}

export interface AppSettings {
  provider: ProviderId;
  providers: Record<ProviderId, ProviderConfig>;
  /** When false, every tool call that modifies the workbook needs explicit approval in the chat. */
  autoApprove: boolean;
  /** Maximum number of cells returned by one read_range call. */
  maxReadCells: number;
  /** Show the answer while it is being generated. */
  streaming: boolean;
  /** Extra instructions added to the system prompt (language, conventions, domain rules). */
  customInstructions: string;
  /** Approximate conversation size (tokens) above which older turns are summarized. */
  maxContextTokens: number;
  reasoningEffort: ReasoningEffort;
  /** Up to MAX_FAVORITES models the user switches between from the chat. */
  favorites: FavoriteModel[];
  customSkills: CustomSkill[];
  /**
   * When true, API keys are stored encrypted with a passphrase and memory-only until
   * unlocked; `keysLocked` says whether they are readable right now.
   */
  protectKeys: boolean;
  keysLocked: boolean;
}

const unprotect = (settings: AppSettings): AppSettings => ({
  ...settings,
  protectKeys: false,
  keysLocked: false,
});

const clamp = (limits: { min: number; max: number; default: number }) => (value: unknown): number => {
  const n = Math.round(Number(value));
  if (!Number.isFinite(n) || n <= 0) return limits.default;
  return Math.min(limits.max, Math.max(limits.min, n));
};

export const READ_CELLS_LIMITS = { min: 100, max: 50000, default: 2000 };
export const clampReadCells = clamp(READ_CELLS_LIMITS);

export const CONTEXT_TOKEN_LIMITS = { min: 8000, max: 1000000, default: 64000 };
export const clampContextTokens = clamp(CONTEXT_TOKEN_LIMITS);

export const MAX_CUSTOM_INSTRUCTIONS = 4000;

export interface ProviderInfo {
  label: string;
  /** OpenAI-compatible base URL (the part before `/chat/completions`). */
  defaultBaseUrl: string;
  defaultModel: string;
  editableBaseUrl: boolean;
  requiresKey: boolean;
  /** Whether routing through the public CORS proxy makes sense (it can never reach localhost). */
  supportsProxy: boolean;
  suggestedModels: string[];
  extraHeaders?: Record<string, string>;
}

export const PROVIDERS: Record<ProviderId, ProviderInfo> = {
  openrouter: {
    label: 'OpenRouter',
    defaultBaseUrl: 'https://openrouter.ai/api/v1',
    defaultModel: 'deepseek/deepseek-chat',
    editableBaseUrl: true,
    requiresKey: true,
    supportsProxy: true,
    suggestedModels: ['deepseek/deepseek-chat', 'anthropic/claude-3.5-sonnet', 'meta-llama/llama-3.3-70b-instruct'],
    extraHeaders: { 'HTTP-Referer': 'https://localhost', 'X-Title': 'Excel-Agent-Pro' },
  },
  gemini: {
    label: 'Google Gemini',
    // Gemini's OpenAI-compatible endpoint supports tool calling and Bearer auth (no key in the URL).
    defaultBaseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai',
    defaultModel: 'gemini-2.5-flash',
    editableBaseUrl: false,
    requiresKey: true,
    supportsProxy: false,
    suggestedModels: ['gemini-2.5-flash', 'gemini-2.5-pro'],
  },
  lmstudio: {
    label: 'LM Studio (local)',
    defaultBaseUrl: 'http://localhost:1234/v1',
    defaultModel: 'local-model',
    editableBaseUrl: true,
    requiresKey: false,
    supportsProxy: false,
    suggestedModels: [],
  },
  opencode: {
    label: 'OpenCode Zen',
    defaultBaseUrl: 'https://opencode.ai/zen/v1',
    defaultModel: 'minimax-m2.5-free',
    editableBaseUrl: false,
    requiresKey: false,
    supportsProxy: true,
    suggestedModels: ['minimax-m2.5-free'],
  },
  custom: {
    label: 'Custom (OpenAI compatible)',
    defaultBaseUrl: '',
    defaultModel: '',
    editableBaseUrl: true,
    requiresKey: false,
    supportsProxy: true,
    suggestedModels: [],
  },
};

export const PROVIDER_IDS = Object.keys(PROVIDERS) as ProviderId[];

const STORAGE_KEY = 'excel_ai_settings';

// Session state: the decrypted keys live here, never back in localStorage.
let sessionSettings: AppSettings | null = null;
let keyBox: EncryptedBox | null = null;

const collectKeys = (settings: AppSettings): Record<ProviderId, string> =>
  Object.fromEntries(PROVIDER_IDS.map(id => [id, settings.providers[id].apiKey])) as Record<ProviderId, string>;

const applyKeys = (settings: AppSettings, keys: Partial<Record<ProviderId, string>>) => {
  const merged: AppSettings = {
    ...settings,
    providers: Object.fromEntries(PROVIDER_IDS.map(id => [
      id,
      { ...settings.providers[id], apiKey: keys[id] ?? settings.providers[id].apiKey },
    ])) as Record<ProviderId, ProviderConfig>,
  };
  return { ...merged, keysLocked: false };
};
const LEGACY_STORAGE_KEY = 'antigravity_settings';

const defaultProviderConfig = (id: ProviderId): ProviderConfig => ({
  baseUrl: PROVIDERS[id].defaultBaseUrl,
  apiKey: '',
  model: PROVIDERS[id].defaultModel,
  useProxy: false,
});

export const createDefaultSettings = (): AppSettings => ({
  provider: 'openrouter',
  providers: Object.fromEntries(PROVIDER_IDS.map(id => [id, defaultProviderConfig(id)])) as Record<ProviderId, ProviderConfig>,
  autoApprove: false,
  maxReadCells: READ_CELLS_LIMITS.default,
  streaming: true,
  customInstructions: '',
  maxContextTokens: CONTEXT_TOKEN_LIMITS.default,
  reasoningEffort: 'default',
  favorites: [],
  customSkills: [],
  protectKeys: false,
  keysLocked: false,
});

const isProviderId = (value: unknown): value is ProviderId =>
  typeof value === 'string' && (PROVIDER_IDS as string[]).includes(value);

const str = (value: unknown, fallback: string): string => (typeof value === 'string' ? value : fallback);

const mergeSettings = (stored: Partial<AppSettings>): AppSettings => {
  const settings = createDefaultSettings();
  if (isProviderId(stored.provider)) settings.provider = stored.provider;
  settings.autoApprove = stored.autoApprove === true;
  settings.maxReadCells = clampReadCells(stored.maxReadCells);
  settings.streaming = stored.streaming !== false;
  settings.customInstructions = str(stored.customInstructions, '').slice(0, MAX_CUSTOM_INSTRUCTIONS);
  settings.maxContextTokens = clampContextTokens(stored.maxContextTokens);
  if (REASONING_EFFORTS.some(e => e.value === stored.reasoningEffort)) settings.reasoningEffort = stored.reasoningEffort!;
  settings.favorites = (Array.isArray(stored.favorites) ? stored.favorites : [])
    .filter(f => f && isProviderId(f.provider) && typeof f.model === 'string' && f.model.trim())
    .slice(0, MAX_FAVORITES)
    .map(f => ({ provider: f.provider, model: f.model.trim(), label: str(f.label, '').trim() || f.model.trim(), ...(f.modelInfo?.id === f.model ? { modelInfo: f.modelInfo } : {}) }));
  settings.customSkills = (Array.isArray(stored.customSkills) ? stored.customSkills : [])
    .filter(s => s && typeof s.id === 'string' && typeof s.name === 'string' && typeof s.instructions === 'string')
    .slice(0, MAX_CUSTOM_SKILLS)
    .map(s => ({ id: s.id, name: s.name, description: str(s.description, ''), instructions: s.instructions.slice(0, MAX_SKILL_INSTRUCTIONS) }));
  // keysLocked is decided by loadSettings (whether an encrypted box exists), not persisted.
  settings.protectKeys = stored.protectKeys === true;
  for (const id of PROVIDER_IDS) {
    const saved = stored.providers?.[id];
    if (saved) settings.providers[id] = { ...settings.providers[id], ...saved };
    if (!PROVIDERS[id].editableBaseUrl) settings.providers[id].baseUrl = PROVIDERS[id].defaultBaseUrl;
    if (settings.providers[id].modelInfo?.id !== settings.providers[id].model) delete settings.providers[id].modelInfo;
  }
  return settings;
};

/** Converts the flat settings format used before the agent refactor. */
const migrateLegacySettings = (legacy: Record<string, unknown>): AppSettings => {
  const settings = createDefaultSettings();
  const p = settings.providers;
  if (isProviderId(legacy.provider)) settings.provider = legacy.provider;

  p.gemini.apiKey = str(legacy.geminiApiKey, '');
  const geminiModel = str(legacy.geminiModel, '');
  // Gemini 1.x models have been retired.
  if (geminiModel && !geminiModel.startsWith('gemini-1.')) p.gemini.model = geminiModel;

  p.lmstudio.baseUrl = str(legacy.lmStudioUrl, p.lmstudio.baseUrl);
  p.lmstudio.model = str(legacy.lmStudioModel, p.lmstudio.model) || p.lmstudio.model;

  p.opencode.apiKey = str(legacy.opencodeApiKey, '');
  p.opencode.model = str(legacy.opencodeModel, p.opencode.model) || p.opencode.model;
  p.opencode.useProxy = legacy.opencodeUseProxy === true;

  p.custom.baseUrl = str(legacy.customBaseUrl, '');
  p.custom.apiKey = str(legacy.customApiKey, '');
  p.custom.model = str(legacy.customModel, '');
  p.custom.useProxy = legacy.customUseProxy === true;
  return settings;
};

export const loadSettings = (): AppSettings => {
  // Already unlocked this session: serve the decrypted copy.
  if (sessionSettings?.keysLocked === false && sessionSettings.protectKeys) return sessionSettings;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      const legacy = localStorage.getItem(LEGACY_STORAGE_KEY);
      if (!legacy) return createDefaultSettings();
      const migrated = migrateLegacySettings(JSON.parse(legacy));
      saveSettings(migrated);
      localStorage.removeItem(LEGACY_STORAGE_KEY);
      return migrated;
    }
    const stored = JSON.parse(raw) as Partial<AppSettings> & { encryptedKeys?: EncryptedBox };
    keyBox = stored.encryptedKeys ?? null;
    const settings = mergeSettings(stored);
    if (settings.protectKeys && keyBox) {
      // Locked: no plaintext keys available until unlockKeys() succeeds.
      sessionSettings = { ...settings, keysLocked: true };
      return sessionSettings;
    }
    sessionSettings = unprotect(settings);
    return sessionSettings;
  } catch (e) {
    console.error('Failed to load settings', e);
    return createDefaultSettings();
  }
};

/** Whether the keys are protected and still need a passphrase this session. */
export const keysAreLocked = () => loadSettings().keysLocked;

/** Decrypts the stored key box into the session; throws on a wrong passphrase. */
export const unlockKeys = async (passphrase: string): Promise<void> => {
  const settings = loadSettings();
  if (!keyBox) throw new Error('There are no encrypted keys to unlock.');
  const keys = await decryptJson<Record<ProviderId, string>>(keyBox, passphrase);
  sessionSettings = applyKeys(settings, keys);
};

/** Re-locks: the decrypted keys are dropped from memory. */
export const lockKeys = () => {
  if (!sessionSettings?.protectKeys) return;
  sessionSettings = { ...loadSettings(), keysLocked: true };
};

/** Turns key protection on or off, encrypting or decrypting the stored keys. */
export const setKeyProtection = async (settings: AppSettings, enable: boolean, passphrase: string): Promise<AppSettings> => {
  if (enable) {
    if (!passphrase) throw new Error('Enter a passphrase.');
    keyBox = await encryptJson(collectKeys(settings), passphrase);
    sessionSettings = { ...settings, protectKeys: true, keysLocked: false };
    // Persisted without plaintext keys: they live only inside the encrypted box.
    persistSettings({ ...sessionSettings, providers: stripKeys(sessionSettings.providers), encryptedKeys: keyBox });
    return sessionSettings;
  }
  keyBox = null;
  sessionSettings = { ...settings, protectKeys: false, keysLocked: false };
  persistSettings(sessionSettings);
  return sessionSettings;
};

/** Makes a favorite the active model. */
export const applyFavorite = (settings: AppSettings, favorite: FavoriteModel): AppSettings => ({
  ...settings,
  provider: favorite.provider,
  providers: {
    ...settings.providers,
    [favorite.provider]: { ...settings.providers[favorite.provider], model: favorite.model, modelInfo: favorite.modelInfo },
  },
});

/** Short display name of the active model: its favorite label, or the last part of its id. */
export const activeModelLabel = (settings: AppSettings): string => {
  const { model } = settings.providers[settings.provider];
  const favorite = settings.favorites.find(f => f.provider === settings.provider && f.model === model);
  return favorite?.label || model.split('/').pop() || model || 'No model';
};

/** Parses an exported settings file (keys stay as they are in the file). */
export const parseImportedSettings = (text: string): AppSettings => {
  const imported = mergeSettings(JSON.parse(text));
  // Importing replaces the session copy; keys in the file are honored.
  sessionSettings = imported;
  return imported;
};

export const saveSettings = (settings: AppSettings) => {
  sessionSettings = settings;
  // With protection on, keys are stored only inside the encrypted box.
  persistSettings(settings.protectKeys ? { ...settings, providers: stripKeys(settings.providers) } : settings);
};

const stripKeys = (providers: Record<ProviderId, ProviderConfig>) =>
  Object.fromEntries(PROVIDER_IDS.map(id => [id, { ...providers[id], apiKey: '' }])) as Record<ProviderId, ProviderConfig>;

const persistSettings = (settings: Partial<AppSettings> & { encryptedKeys?: EncryptedBox }) => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch (e) {
    console.error('Failed to save settings', e);
  }
};

export const importEncryptedKeys = async (settings: AppSettings, box: unknown, passphrase: string): Promise<AppSettings> => {
  const keys = await decryptJson<Record<ProviderId, string>>(box as EncryptedBox, passphrase);
  keyBox = box as EncryptedBox;
  sessionSettings = applyKeys(settings, keys);
  persistSettings({ ...sessionSettings, encryptedKeys: keyBox });
  return sessionSettings;
};

export const exportEncryptedKeys = () => keyBox;

export const keysAreEncrypted = () => keyBox !== null;
