import { useState } from 'react';
import { Button, Checkbox, Field, Input, Select, Textarea, makeStyles, tokens } from '@fluentui/react-components';
import {
  CONTEXT_TOKEN_LIMITS,
  MAX_CUSTOM_INSTRUCTIONS,
  PROVIDERS,
  PROVIDER_IDS,
  READ_CELLS_LIMITS,
  clampContextTokens,
  clampReadCells,
  loadSettings,
  saveSettings,
  type AppSettings,
  type ModelInfo,
  type ProviderConfig,
  type ProviderId,
} from '../utils/storage';
import { listModels } from '../agent/llmClient';

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
  warning: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
});

/** Starting points for the custom instructions, appended to what the user already wrote. */
const TEMPLATES: Record<string, string> = {
  'Spanish answers': 'Respond in Spanish. Use the decimal and date conventions of Latin America when explaining values, but keep formulas in English with comma separators.',
  'Financial modeling': [
    'Follow financial modeling conventions: blue font (#0000FF) for hard-coded inputs, black for formulas, green (#008000) for links to other sheets.',
    'Keep assumptions in a separate, clearly labeled inputs section and reference them instead of hard-coding numbers in formulas.',
    'Show years as columns and line items as rows. Use number formats like #,##0;(#,##0) for amounts and 0.0% for rates.',
    'For valuations (DCF): project free cash flow, discount with WACC using XNPV/NPV, compute terminal value with the Gordon growth or exit multiple method, and add a sensitivity table.',
  ].join('\n'),
  Accounting: 'Use accounting number format, keep debits and credits balanced, flag differences with conditional formatting, and never overwrite source ledgers: put results in new sheets.',
  'Data cleaning': 'Before changing data, profile it and report issues (blanks, duplicates, inconsistent text, numbers stored as text). Make cleaned copies in new sheets instead of editing the original data.',
  'Dashboards': 'When building reports, put summary KPIs at the top, use PivotTables or SUMIFS for aggregates, add clear charts with titles and axis labels, and keep a consistent color palette.',
};

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

  const loadModels = async () => {
    setIsLoadingModels(true);
    setModelError('');
    try {
      const models = await listModels(settings);
      setLoadedModels(prev => ({ ...prev, [provider]: models }));
      const current = models.find(m => m.id === config.model);
      if (current) updateConfig({ modelInfo: current });
      if (models.length === 0) setModelError('The provider returned no models.');
    } catch (e) {
      setModelError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsLoadingModels(false);
    }
  };

  const addTemplate = (name: string) => {
    const template = TEMPLATES[name];
    if (!template) return;
    setSettings(prev => ({
      ...prev,
      customInstructions: (prev.customInstructions.trim() ? `${prev.customInstructions.trim()}\n\n${template}` : template).slice(0, MAX_CUSTOM_INSTRUCTIONS),
    }));
  };

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
      <Field label="Add a template">
        <Select value="" onChange={(_, data) => addTemplate(data.value)}>
          <option value="">Choose a template…</option>
          {Object.keys(TEMPLATES).map(name => <option key={name} value={name}>{name}</option>)}
        </Select>
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
