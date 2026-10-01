import { useState } from 'react';
import { Button, Checkbox, Field, Input, Select, Textarea, makeStyles, tokens } from '@fluentui/react-components';
import {
  CONTEXT_TOKEN_LIMITS,
  MAX_CUSTOM_INSTRUCTIONS,
  MAX_CUSTOM_SKILLS,
  MAX_FAVORITES,
  MAX_SKILL_INSTRUCTIONS,
  PROVIDERS,
  PROVIDER_IDS,
  READ_CELLS_LIMITS,
  clampContextTokens,
  clampReadCells,
  loadSettings,
  saveSettings,
  activeModelLabel,
  type AppSettings,
  type CustomSkill,
  type FavoriteModel,
  type ModelInfo,
  type ProviderConfig,
  type ProviderId,
} from '../utils/storage';
import { listModels } from '../agent/llmClient';
import { BUILT_IN_SKILLS, getSkills, skillIdFromName } from '../agent/skills';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    padding: '16px',
  },
  header: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
  section: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginTop: '8px',
  },
  row: {
    display: 'flex',
    gap: '8px',
    alignItems: 'flex-end',
  },
  grow: {
    flex: 1,
  },
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  capabilities: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  listItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  warning: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
});

interface SkillDraft {
  id: string | null;
  name: string;
  description: string;
  instructions: string;
}

interface SettingsProps {
  onBack: () => void;
}

const formatContext = (tokensCount: number) => (tokensCount >= 1000 ? `${Math.round(tokensCount / 1000)}k` : String(tokensCount));

export const Settings: React.FC<SettingsProps> = ({ onBack }) => {
  const styles = useStyles();
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [loadedModels, setLoadedModels] = useState<Partial<Record<ProviderId, ModelInfo[]>>>({});
  const [isLoadingModels, setIsLoadingModels] = useState(false);
  const [modelError, setModelError] = useState('');
  const [onlyToolModels, setOnlyToolModels] = useState(true);
  // Kept as text while editing; validated and clamped on save.
  const [maxReadCellsText, setMaxReadCellsText] = useState(() => String(settings.maxReadCells));
  const [maxContextText, setMaxContextText] = useState(() => String(settings.maxContextTokens));

  const provider = settings.provider;
  const info = PROVIDERS[provider];
  const config = settings.providers[provider];
  const loaded = loadedModels[provider];
  const knowsTools = loaded?.some(m => m.supportsTools !== undefined) ?? false;
  const visibleModels = (loaded ?? []).filter(m => !onlyToolModels || m.supportsTools !== false);
  const modelOptions = [...new Set([...info.suggestedModels, ...visibleModels.map(m => m.id)])];
  const modelInfo = config.modelInfo?.id === config.model ? config.modelInfo : undefined;

  const updateConfig = (patch: Partial<ProviderConfig>) =>
    setSettings(prev => ({
      ...prev,
      providers: { ...prev.providers, [prev.provider]: { ...prev.providers[prev.provider], ...patch } },
    }));

  const setModel = (model: string) => {
    const found = loadedModels[provider]?.find(m => m.id === model.trim());
    // Keep known capabilities when the model is unchanged and the list wasn't reloaded.
    const keep = config.modelInfo?.id === model.trim() ? config.modelInfo : undefined;
    updateConfig({ model, modelInfo: found ?? keep });
  };

  const loadModelsFor = async (providerId: ProviderId) => {
    setIsLoadingModels(true);
    setModelError('');
    try {
      const models = await listModels({ ...settings, provider: providerId });
      setLoadedModels(prev => ({ ...prev, [providerId]: models }));
      const currentModel = settings.providers[providerId].model;
      const current = models.find(m => m.id === currentModel);
      if (current) {
        setSettings(prev => ({ ...prev, providers: { ...prev.providers, [providerId]: { ...prev.providers[providerId], modelInfo: current } } }));
      }
      // Refresh the capabilities of this provider's favorites too.
      setSettings(prev => ({
        ...prev,
        favorites: prev.favorites.map(f => (f.provider === providerId ? { ...f, modelInfo: models.find(m => m.id === f.model) ?? f.modelInfo } : f)),
      }));
      if (models.length === 0) setModelError('The provider returned no models.');
    } catch (e) {
      setModelError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsLoadingModels(false);
    }
  };
  const loadModels = () => loadModelsFor(provider);

  const currentIsFavorite = settings.favorites.some(f => f.provider === provider && f.model === config.model);

  // Add any model as a favorite directly, independently of the model configured above.
  const [favoriteDraft, setFavoriteDraft] = useState<{ provider: ProviderId; model: string; label: string }>({ provider, model: '', label: '' });
  const favoriteDraftInfo = loadedModels[favoriteDraft.provider]?.find(m => m.id === favoriteDraft.model.trim());
  const favoriteDraftModels = [...new Set([...PROVIDERS[favoriteDraft.provider].suggestedModels, ...(loadedModels[favoriteDraft.provider] ?? []).map(m => m.id)])];
  const favoriteDraftDuplicate = settings.favorites.some(f => f.provider === favoriteDraft.provider && f.model === favoriteDraft.model.trim());

  const addFavorite = (entry: { provider: ProviderId; model: string; label: string; modelInfo?: ModelInfo }) => {
    const model = entry.model.trim();
    if (!model || settings.favorites.length >= MAX_FAVORITES) return;
    if (settings.favorites.some(f => f.provider === entry.provider && f.model === model)) return;
    setSettings(prev => ({
      ...prev,
      favorites: [...prev.favorites, { provider: entry.provider, model, label: entry.label.trim() || model.split('/').pop() || model, modelInfo: entry.modelInfo }],
    }));
  };
  const addFavoriteFromDraft = () => {
    addFavorite({ ...favoriteDraft, modelInfo: favoriteDraftInfo });
    setFavoriteDraft({ provider: favoriteDraft.provider, model: '', label: '' });
  };
  const addCurrentAsFavorite = () => addFavorite({ provider, model: config.model, label: activeModelLabel(settings), modelInfo });
  const updateFavorite = (index: number, patch: Partial<FavoriteModel>) =>
    setSettings(prev => ({ ...prev, favorites: prev.favorites.map((f, i) => (i === index ? { ...f, ...patch } : f)) }));
  const removeFavorite = (index: number) =>
    setSettings(prev => ({ ...prev, favorites: prev.favorites.filter((_, i) => i !== index) }));
  const moveFavorite = (index: number) =>
    setSettings(prev => {
      const favorites = [...prev.favorites];
      [favorites[index - 1], favorites[index]] = [favorites[index], favorites[index - 1]];
      return { ...prev, favorites };
    });

  const [skillDraft, setSkillDraft] = useState<SkillDraft | null>(null);
  const saveSkillDraft = () => {
    if (!skillDraft || !skillDraft.name.trim() || !skillDraft.instructions.trim()) return;
    const id = skillDraft.id ?? skillIdFromName(skillDraft.name);
    const skill: CustomSkill = {
      id,
      name: skillDraft.name.trim(),
      description: skillDraft.description.trim(),
      instructions: skillDraft.instructions.slice(0, MAX_SKILL_INSTRUCTIONS),
    };
    setSettings(prev => {
      const exists = prev.customSkills.some(s => s.id === id);
      return { ...prev, customSkills: exists ? prev.customSkills.map(s => (s.id === id ? skill : s)) : [...prev.customSkills, skill] };
    });
    setSkillDraft(null);
  };
  const removeSkill = (id: string) => setSettings(prev => ({ ...prev, customSkills: prev.customSkills.filter(s => s.id !== id) }));
  const editSkill = (skill: { id: string; name: string; description: string; instructions: string }) =>
    setSkillDraft({ id: skill.id, name: skill.name, description: skill.description, instructions: skill.instructions });

  const handleSave = () => {
    saveSettings({
      ...settings,
      maxReadCells: clampReadCells(maxReadCellsText),
      maxContextTokens: clampContextTokens(maxContextText),
    });
    onBack();
  };

  const capabilities = modelInfo && [
    modelInfo.supportsTools === undefined ? null : modelInfo.supportsTools ? '✓ tool calling' : null,
    modelInfo.supportsImages ? '✓ images' : null,
    modelInfo.supportsFiles ? '✓ PDFs' : null,
    modelInfo.contextLength ? `${formatContext(modelInfo.contextLength)} context` : null,
  ].filter(Boolean).join(' · ');

  return (
    <div className={styles.container}>
      <div className={styles.header}>Settings</div>

      <Field label="AI provider">
        <Select
          value={provider}
          onChange={(_, data) => {
            setSettings(prev => ({ ...prev, provider: data.value as ProviderId }));
            setModelError('');
          }}
        >
          {PROVIDER_IDS.map(id => <option key={id} value={id}>{PROVIDERS[id].label}</option>)}
        </Select>
      </Field>

      {provider === 'custom' && (
        <div className={styles.hint}>
          Any OpenAI-compatible API with tool calling, e.g.<br />
          • Groq: https://api.groq.com/openai/v1<br />
          • Together AI: https://api.together.xyz/v1
        </div>
      )}

      {info.editableBaseUrl && (
        <Field label="Base URL" hint="Base URL (…/v1) or the full …/chat/completions URL.">
          <Input
            value={config.baseUrl}
            placeholder={info.defaultBaseUrl || 'https://api.example.com/v1'}
            onChange={(_, data) => updateConfig({ baseUrl: data.value })}
          />
        </Field>
      )}

      <Field label={info.requiresKey ? 'API key' : 'API key (optional)'}>
        <Input type="password" value={config.apiKey} onChange={(_, data) => updateConfig({ apiKey: data.value })} />
      </Field>

      <div className={styles.row}>
        <Field
          label="Model"
          className={styles.grow}
          hint={loaded ? `${visibleModels.length} models available — type to filter.` : 'Type a model id or load the list.'}
        >
          <Input list="model-options" value={config.model} placeholder={info.defaultModel || 'model-id'} onChange={(_, data) => setModel(data.value)} />
        </Field>
        <Button onClick={loadModels} disabled={isLoadingModels || (info.requiresKey && !config.apiKey)}>
          {isLoadingModels ? 'Loading…' : 'Load models'}
        </Button>
      </div>
      <datalist id="model-options">
        {modelOptions.map(m => <option key={m} value={m} />)}
      </datalist>
      {knowsTools && (
        <Checkbox label="Only show models that support tool calling" checked={onlyToolModels} onChange={(_, data) => setOnlyToolModels(data.checked === true)} />
      )}
      {modelError && <Field validationState="error" validationMessage={modelError} />}
      {modelInfo?.supportsTools === false && (
        <div className={styles.warning}>⚠ This model does not support tool calling, so the agent can't use it. Choose another model.</div>
      )}
      {capabilities && <div className={styles.capabilities}>{capabilities}</div>}
      {!modelInfo && <div className={styles.hint}>The model must support tool (function) calling. Load the list to see each model's capabilities.</div>}

      {info.supportsProxy && (
        <Checkbox
          label="Use the public CORS proxy (corsproxy.io). Only enable it if direct requests are blocked: your API key and workbook data will pass through that third-party service."
          checked={config.useProxy}
          onChange={(_, data) => updateConfig({ useProxy: data.checked === true })}
        />
      )}

      <div className={styles.section}>Favorite models ({settings.favorites.length}/{MAX_FAVORITES})</div>
      <div className={styles.hint}>Add any model you want, from any provider. Switch between them from the model button in the chat. Each favorite uses that provider's own API key, configured above when you select it there.</div>
      {settings.favorites.map((f, i) => (
        <div key={`${f.provider}::${f.model}`} className={styles.listItem}>
          <div className={styles.row}>
            <Input className={styles.grow} size="small" value={f.label} aria-label="Display name" onChange={(_, data) => updateFavorite(i, { label: data.value })} />
            <Button size="small" appearance="subtle" disabled={i === 0} onClick={() => moveFavorite(i)} aria-label="Move up" title="Move up">↑</Button>
            <Button size="small" appearance="subtle" onClick={() => removeFavorite(i)} aria-label="Remove" title="Remove">✕</Button>
          </div>
          <div className={styles.hint}>{PROVIDERS[f.provider].label} · {f.model}</div>
        </div>
      ))}

      {settings.favorites.length < MAX_FAVORITES && (
        <div className={styles.listItem}>
          <div className={styles.row}>
            <Field label="Provider">
              <Select value={favoriteDraft.provider} onChange={(_, data) => setFavoriteDraft({ provider: data.value as ProviderId, model: '', label: '' })}>
                {PROVIDER_IDS.map(id => <option key={id} value={id}>{PROVIDERS[id].label}</option>)}
              </Select>
            </Field>
          </div>
          <div className={styles.row}>
            <Field className={styles.grow} label="Model" hint={loadedModels[favoriteDraft.provider] ? `${loadedModels[favoriteDraft.provider]!.length} models loaded` : 'Type a model id or load the list.'}>
              <Input
                list="favorite-model-options"
                value={favoriteDraft.model}
                placeholder={PROVIDERS[favoriteDraft.provider].defaultModel || 'model-id'}
                onChange={(_, data) => setFavoriteDraft({ ...favoriteDraft, model: data.value })}
              />
            </Field>
            <Button
              onClick={() => void loadModelsFor(favoriteDraft.provider)}
              disabled={isLoadingModels || (PROVIDERS[favoriteDraft.provider].requiresKey && !settings.providers[favoriteDraft.provider].apiKey)}
            >
              {isLoadingModels ? 'Loading…' : 'Load models'}
            </Button>
          </div>
          <datalist id="favorite-model-options">
            {favoriteDraftModels.map(m => <option key={m} value={m} />)}
          </datalist>
          {PROVIDERS[favoriteDraft.provider].requiresKey && !settings.providers[favoriteDraft.provider].apiKey && (
            <div className={styles.hint}>
              You can type a model id and add it directly. To use "Load models", first switch "AI provider" above to {PROVIDERS[favoriteDraft.provider].label} and set its API key.
            </div>
          )}
          {favoriteDraftDuplicate && <div className={styles.hint}>Already a favorite.</div>}
          <div className={styles.row}>
            <Field className={styles.grow} label="Display name (optional)">
              <Input value={favoriteDraft.label} placeholder={favoriteDraft.model.split('/').pop() || 'name'} onChange={(_, data) => setFavoriteDraft({ ...favoriteDraft, label: data.value })} />
            </Field>
            <Button appearance="primary" onClick={addFavoriteFromDraft} disabled={!favoriteDraft.model.trim() || favoriteDraftDuplicate}>
              + Add to favorites
            </Button>
          </div>
        </div>
      )}
      {!currentIsFavorite && config.model.trim() && settings.favorites.length < MAX_FAVORITES && (
        <Button appearance="subtle" onClick={addCurrentAsFavorite}>☆ Also add the model configured above ({info.label} · {config.model})</Button>
      )}

      <div className={styles.section}>Skills</div>
      <div className={styles.hint}>
        Expert playbooks the agent loads when a request needs them, or that you attach from the Skills button or with /skill-id at the start of a message.
        Built-in skills can be customized: your version replaces the original.
      </div>
      {getSkills(settings).map(skill => (
        <div key={skill.id} className={styles.listItem}>
          <div className={styles.row}>
            <div className={styles.grow}>
              <b>{skill.name}</b> <span className={styles.hint}>/{skill.id}{skill.builtIn ? '' : BUILT_IN_SKILLS.some(b => b.id === skill.id) ? ' · customized' : ' · custom'}</span>
            </div>
            <Button size="small" appearance="subtle" onClick={() => editSkill(skill)}>{skill.builtIn ? 'Customize' : 'Edit'}</Button>
            {!skill.builtIn && (
              <Button size="small" appearance="subtle" onClick={() => removeSkill(skill.id)} title={BUILT_IN_SKILLS.some(b => b.id === skill.id) ? 'Restore the built-in version' : 'Delete'}>
                {BUILT_IN_SKILLS.some(b => b.id === skill.id) ? 'Reset' : '✕'}
              </Button>
            )}
          </div>
          <div className={styles.hint}>{skill.description}</div>
        </div>
      ))}
      {skillDraft ? (
        <div className={styles.listItem}>
          <Field label="Name" required>
            <Input value={skillDraft.name} placeholder="e.g. Monthly sales report" onChange={(_, data) => setSkillDraft({ ...skillDraft, name: data.value })} />
          </Field>
          <Field label="When to use it" hint="Shown to the agent so it knows when to load the skill.">
            <Input value={skillDraft.description} placeholder="e.g. Building the monthly sales report for management." onChange={(_, data) => setSkillDraft({ ...skillDraft, description: data.value })} />
          </Field>
          <Field label="Instructions" required hint={`Steps, conventions, layout, formats… (max ${MAX_SKILL_INSTRUCTIONS} characters)`}>
            <Textarea
              value={skillDraft.instructions}
              maxLength={MAX_SKILL_INSTRUCTIONS}
              rows={8}
              resize="vertical"
              onChange={(_, data) => setSkillDraft({ ...skillDraft, instructions: data.value })}
            />
          </Field>
          <div className={styles.row}>
            <Button appearance="primary" size="small" onClick={saveSkillDraft} disabled={!skillDraft.name.trim() || !skillDraft.instructions.trim()}>Save skill</Button>
            <Button size="small" onClick={() => setSkillDraft(null)}>Cancel</Button>
          </div>
        </div>
      ) : (
        <Button
          onClick={() => setSkillDraft({ id: null, name: '', description: '', instructions: '' })}
          disabled={settings.customSkills.length >= MAX_CUSTOM_SKILLS}
        >
          + New skill
        </Button>
      )}

      <div className={styles.section}>Agent</div>

      <Field label="Custom instructions" hint="Added to every request: language, conventions, domain rules.">
        <Textarea
          value={settings.customInstructions}
          maxLength={MAX_CUSTOM_INSTRUCTIONS}
          resize="vertical"
          rows={4}
          placeholder="e.g. Respond in Spanish. Amounts are in MXN. Our fiscal year starts in April."
          onChange={(_, data) => setSettings(prev => ({ ...prev, customInstructions: data.value }))}
        />
      </Field>

      <Checkbox
        label="Apply workbook changes without asking for approval (irreversible actions always ask)"
        checked={settings.autoApprove}
        onChange={(_, data) => setSettings(prev => ({ ...prev, autoApprove: data.checked === true }))}
      />
      <Checkbox
        label="Stream answers as they are generated"
        checked={settings.streaming}
        onChange={(_, data) => setSettings(prev => ({ ...prev, streaming: data.checked === true }))}
      />

      <Field
        label="Max cells per read"
        hint={`Cells the agent can read in one read_range call (${READ_CELLS_LIMITS.min}–${READ_CELLS_LIMITS.max}, default ${READ_CELLS_LIMITS.default}). Higher values send more tokens (slower, more expensive).`}
      >
        <Input
          type="number"
          min={READ_CELLS_LIMITS.min}
          max={READ_CELLS_LIMITS.max}
          step={500}
          value={maxReadCellsText}
          onChange={(_, data) => setMaxReadCellsText(data.value)}
          onBlur={() => setMaxReadCellsText(String(clampReadCells(maxReadCellsText)))}
        />
      </Field>

      <Field
        label="Conversation size before summarizing (tokens)"
        hint={`Older messages are summarized above this size (${formatContext(CONTEXT_TOKEN_LIMITS.min)}–${formatContext(CONTEXT_TOKEN_LIMITS.max)}, default ${formatContext(CONTEXT_TOKEN_LIMITS.default)}). It is also capped by the model's context window when known.`}
      >
        <Input
          type="number"
          min={CONTEXT_TOKEN_LIMITS.min}
          max={CONTEXT_TOKEN_LIMITS.max}
          step={8000}
          value={maxContextText}
          onChange={(_, data) => setMaxContextText(data.value)}
          onBlur={() => setMaxContextText(String(clampContextTokens(maxContextText)))}
        />
      </Field>

      <div className={styles.hint}>
        Settings and API keys are stored in this browser's local storage for the add-in's web address. They are only sent to the provider you choose (and to the CORS proxy, if enabled). Chat history and undo data are stored in this browser per workbook.
      </div>

      <Button appearance="primary" onClick={handleSave}>Save & close</Button>
      <Button appearance="subtle" onClick={onBack}>Cancel</Button>
    </div>
  );
};
