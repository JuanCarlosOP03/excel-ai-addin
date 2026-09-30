import { useState } from 'react';
import { Button, Checkbox, Field, Input, Select, makeStyles, tokens } from '@fluentui/react-components';
import {
  PROVIDERS,
  PROVIDER_IDS,
  READ_CELLS_LIMITS,
  clampReadCells,
  loadSettings,
  saveSettings,
  type AppSettings,
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
});

interface SettingsProps {
  onBack: () => void;
}

export const Settings: React.FC<SettingsProps> = ({ onBack }) => {
  const styles = useStyles();
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [loadedModels, setLoadedModels] = useState<Partial<Record<ProviderId, string[]>>>({});
  const [isLoadingModels, setIsLoadingModels] = useState(false);
  const [modelError, setModelError] = useState('');
  // Kept as text while editing; validated and clamped on save.
  const [maxReadCellsText, setMaxReadCellsText] = useState(() => String(settings.maxReadCells));

  const provider = settings.provider;
  const info = PROVIDERS[provider];
  const config = settings.providers[provider];
  const modelOptions = [...new Set([...info.suggestedModels, ...(loadedModels[provider] ?? [])])];

  const updateConfig = (patch: Partial<ProviderConfig>) =>
    setSettings(prev => ({
      ...prev,
      providers: { ...prev.providers, [prev.provider]: { ...prev.providers[prev.provider], ...patch } },
    }));

  const loadModels = async () => {
    setIsLoadingModels(true);
    setModelError('');
    try {
      const models = await listModels(settings);
      setLoadedModels(prev => ({ ...prev, [provider]: models }));
      if (models.length === 0) setModelError('The provider returned no models.');
    } catch (e) {
      setModelError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsLoadingModels(false);
    }
  };

  const handleSave = () => {
    saveSettings({ ...settings, maxReadCells: clampReadCells(maxReadCellsText) });
    onBack();
  };

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
          hint={loadedModels[provider] ? `${loadedModels[provider].length} models loaded — type to filter.` : 'Type a model id or load the list.'}
        >
          <Input
            list="model-options"
            value={config.model}
            placeholder={info.defaultModel || 'model-id'}
            onChange={(_, data) => updateConfig({ model: data.value })}
          />
        </Field>
        <Button onClick={loadModels} disabled={isLoadingModels || (info.requiresKey && !config.apiKey)}>
          {isLoadingModels ? 'Loading…' : 'Load models'}
        </Button>
      </div>
      <datalist id="model-options">
        {modelOptions.map(m => <option key={m} value={m} />)}
      </datalist>
      {modelError && <Field validationState="error" validationMessage={modelError} />}
      <div className={styles.hint}>The model must support tool (function) calling.</div>

      {info.supportsProxy && (
        <Checkbox
          label="Use the public CORS proxy (corsproxy.io). Only enable it if direct requests are blocked: your API key and workbook data will pass through that third-party service."
          checked={config.useProxy}
          onChange={(_, data) => updateConfig({ useProxy: data.checked === true })}
        />
      )}

      <Field
        label="Max cells per read"
        hint={`Cells the agent can read in one read_range call (${READ_CELLS_LIMITS.min}–${READ_CELLS_LIMITS.max}, default ${READ_CELLS_LIMITS.default}). Higher values let it see more data at once but send more tokens to the provider (slower, more expensive, may exceed the model's context).`}
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

      <Checkbox
        label="Apply workbook changes without asking for approval"
        checked={settings.autoApprove}
        onChange={(_, data) => setSettings(prev => ({ ...prev, autoApprove: data.checked === true }))}
      />

      <div className={styles.hint}>
        Settings and API keys are stored in this browser's local storage for the add-in's web address. They are only sent to the provider you choose (and to the CORS proxy, if enabled).
      </div>

      <Button appearance="primary" onClick={handleSave}>Save & close</Button>
      <Button appearance="subtle" onClick={onBack}>Cancel</Button>
    </div>
  );
};
