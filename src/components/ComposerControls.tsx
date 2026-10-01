import {
  Button,
  Menu,
  MenuButton,
  MenuDivider,
  MenuGroup,
  MenuGroupHeader,
  MenuItem,
  MenuItemCheckbox,
  MenuItemRadio,
  MenuList,
  MenuPopover,
  MenuTrigger,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import { Attach24Regular, Clock24Regular, FastForward20Regular } from '@fluentui/react-icons';
import {
  MAX_FAVORITES,
  PROVIDERS,
  REASONING_EFFORTS,
  activeModelLabel,
  type AppSettings,
  type FavoriteModel,
  type ReasoningEffort,
} from '../utils/storage';
import type { Skill } from '../agent/skills';

const useStyles = makeStyles({
  modelButton: {
    minWidth: 0,
    maxWidth: '150px',
    paddingLeft: '8px',
    paddingRight: '6px',
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground2,
    '& > span:first-child': {
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      whiteSpace: 'nowrap',
    },
  },
  iconButton: {
    color: tokens.colorNeutralForeground3,
  },
  iconButtonActive: {
    color: tokens.colorBrandForeground1,
  },
  note: {
    padding: '4px 8px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    maxWidth: '240px',
  },
});

const favoriteKey = (f: Pick<FavoriteModel, 'provider' | 'model'>) => `${f.provider}::${f.model}`;

interface ModelMenuProps {
  settings: AppSettings;
  disabled?: boolean;
  onSelect: (favorite: FavoriteModel) => void;
  onAddCurrent: () => void;
  onManage: () => void;
}

export const ModelMenu: React.FC<ModelMenuProps> = ({ settings, disabled, onSelect, onAddCurrent, onManage }) => {
  const styles = useStyles();
  const config = settings.providers[settings.provider];
  const current = favoriteKey({ provider: settings.provider, model: config.model });
  const isFavorite = settings.favorites.some(f => favoriteKey(f) === current);
  const label = activeModelLabel(settings);

  return (
    <Menu
      checkedValues={{ model: [current] }}
      onCheckedValueChange={(_, data) => {
        const favorite = settings.favorites.find(f => favoriteKey(f) === data.checkedItems[0]);
        if (favorite) onSelect(favorite);
      }}
    >
      <MenuTrigger disableButtonEnhancement>
        <MenuButton size="small" appearance="subtle" className={styles.modelButton} disabled={disabled} title={`${PROVIDERS[settings.provider].label}: ${config.model}`}>
          {label}
        </MenuButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuGroup>
            <MenuGroupHeader>Favorite models</MenuGroupHeader>
            {settings.favorites.map(f => (
              <MenuItemRadio key={favoriteKey(f)} name="model" value={favoriteKey(f)} secondaryContent={PROVIDERS[f.provider].label}>
                {f.label}
              </MenuItemRadio>
            ))}
            {!isFavorite && (
              <MenuItemRadio name="model" value={current} secondaryContent="current">
                {label}
              </MenuItemRadio>
            )}
          </MenuGroup>
          {settings.favorites.length === 0 && <div className={styles.note}>Add up to {MAX_FAVORITES} favorite models to switch between them here.</div>}
          <MenuDivider />
          {!isFavorite && config.model && (
            <MenuItem onClick={onAddCurrent} disabled={settings.favorites.length >= MAX_FAVORITES}>
              ☆ Add "{label}" to favorites
            </MenuItem>
          )}
          <MenuItem onClick={onManage}>Manage models…</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

interface EffortMenuProps {
  settings: AppSettings;
  disabled?: boolean;
  onChange: (effort: ReasoningEffort) => void;
}

export const EffortMenu: React.FC<EffortMenuProps> = ({ settings, disabled, onChange }) => {
  const styles = useStyles();
  const config = settings.providers[settings.provider];
  const unsupported = config.modelInfo?.id === config.model && config.modelInfo.supportsReasoning === false;
  const current = REASONING_EFFORTS.find(e => e.value === settings.reasoningEffort) ?? REASONING_EFFORTS[0];
  const active = !unsupported && settings.reasoningEffort !== 'default';

  return (
    <Menu checkedValues={{ effort: [settings.reasoningEffort] }} onCheckedValueChange={(_, data) => onChange(data.checkedItems[0] as ReasoningEffort)}>
      <MenuTrigger disableButtonEnhancement>
        <Button
          size="small"
          appearance="subtle"
          shape="circular"
          className={mergeClasses(styles.iconButton, active && styles.iconButtonActive)}
          icon={<Clock24Regular />}
          disabled={disabled}
          aria-label={`Reasoning effort: ${current.label}`}
          title={unsupported ? "Reasoning effort (the selected model doesn't support it)" : `Reasoning effort: ${current.label}`}
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuGroup>
            <MenuGroupHeader>Reasoning effort</MenuGroupHeader>
            {REASONING_EFFORTS.map(e => (
              <MenuItemRadio key={e.value} name="effort" value={e.value} disabled={unsupported && e.value !== 'default'}>
                {e.label}
              </MenuItemRadio>
            ))}
          </MenuGroup>
          <div className={styles.note}>
            {unsupported
              ? "The selected model doesn't support reasoning effort."
              : 'Higher effort thinks longer: better for complex models and analyses, but slower and more expensive. "Default" lets the model decide.'}
          </div>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

interface SkillsMenuProps {
  skills: Skill[];
  selected: string[];
  disabled?: boolean;
  onChange: (ids: string[]) => void;
  onManage: () => void;
}

export const SkillsMenu: React.FC<SkillsMenuProps> = ({ skills, selected, disabled, onChange, onManage }) => {
  const styles = useStyles();
  return (
    <Menu checkedValues={{ skills: selected }} onCheckedValueChange={(_, data) => onChange(data.checkedItems)}>
      <MenuTrigger disableButtonEnhancement>
        <Button
          size="small"
          appearance="subtle"
          shape="circular"
          className={mergeClasses(styles.iconButton, selected.length > 0 && styles.iconButtonActive)}
          icon={<FastForward20Regular />}
          disabled={disabled}
          aria-label={selected.length ? `Skills (${selected.length} selected)` : 'Skills'}
          title={selected.length ? `Skills (${selected.length} selected)` : 'Attach a skill (expert playbook) to your next message'}
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuGroup>
            <MenuGroupHeader>Use with the next message</MenuGroupHeader>
            {skills.map(s => (
              <MenuItemCheckbox key={s.id} name="skills" value={s.id} secondaryContent={s.builtIn ? undefined : 'custom'}>
                <span title={s.description}>{s.name}</span>
              </MenuItemCheckbox>
            ))}
          </MenuGroup>
          <div className={styles.note}>The agent also loads skills on its own when a request needs them. Tip: type /{skills[0]?.id ?? 'skill'} at the start of a message.</div>
          <MenuDivider />
          <MenuItem onClick={onManage}>Manage skills…</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

interface AttachButtonProps {
  disabled?: boolean;
  onClick: () => void;
}

export const AttachButton: React.FC<AttachButtonProps> = ({ disabled, onClick }) => {
  const styles = useStyles();
  return (
    <Button
      size="small"
      appearance="subtle"
      shape="circular"
      className={styles.iconButton}
      icon={<Attach24Regular />}
      disabled={disabled}
      aria-label="Attach files"
      title="Attach CSV, JSON, text, PDF or images"
      onClick={onClick}
    />
  );
};
