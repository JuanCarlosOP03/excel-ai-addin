import { useRef, useState } from 'react';
import {
  Button,
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  AddCircle24Regular,
  ArrowDownload24Regular,
  ArrowUpload24Regular,
  History24Regular,
  MoreHorizontal24Regular,
  Settings24Regular,
} from '@fluentui/react-icons';
import { Settings } from './components/Settings';
import { Chat, type ChatHandle } from './components/Chat';
import { loadSettings, parseImportedSettings, saveSettings } from './utils/storage';
import type { ChatSession } from './utils/persistence';

const useStyles = makeStyles({
  appContainer: {
    height: '100vh',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  header: {
    padding: '6px 12px',
    backgroundColor: '#0f703b',
    color: tokens.colorNeutralForegroundOnBrand,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexShrink: 0,
  },
  title: {
    fontWeight: 600,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  headerButton: {
    color: tokens.colorNeutralForegroundOnBrand,
    minWidth: '32px',
  },
  menuButton: {
    color: tokens.colorNeutralForegroundOnBrand,
  },
  content: {
    flex: 1,
    overflowY: 'auto',
  },
  sessionLabel: {
    maxWidth: '260px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    display: 'inline-block',
    verticalAlign: 'bottom',
  },
  hiddenInput: {
    display: 'none',
  },
});

const downloadJson = (data: unknown, name: string) => {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
};

function App() {
  const styles = useStyles();
  const [view, setView] = useState<'chat' | 'settings'>('chat');
  // Remounts the chat so re-imported settings are picked up.
  const [chatKey, setChatKey] = useState(0);
  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const chatRef = useRef<ChatHandle>(null);
  const importInputRef = useRef<HTMLInputElement>(null);

  const openHistory = () => {
    chatRef.current?.listHistory().then(list => setSessions(list)).catch(e => console.error('Failed to load chat history', e));
  };

  const exportConfig = () => downloadJson(loadSettings(), 'excel-ai-config.json');

  const importConfig = async (file: File) => {
    try {
      const settings = parseImportedSettings(await file.text());
      saveSettings(settings);
      setChatKey(key => key + 1);
    } catch (e) {
      alert(`Couldn't import the configuration: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  const chatView = view === 'chat';

  return (
    <div className={styles.appContainer}>
      <div className={styles.header}>
        <span className={styles.title}>Excel AI</span>
        {chatView && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
            <Menu onOpenChange={(_, data) => data.open && openHistory()}>
              <MenuTrigger disableButtonEnhancement>
                <Button
                  appearance="transparent"
                  className={styles.headerButton}
                  icon={<History24Regular />}
                  aria-label="Chat history"
                  title="Chat history"
                />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem disabled>Previous chats</MenuItem>
                  {sessions.length === 0 && <MenuItem disabled>No previous chats</MenuItem>}
                  {sessions.map(s => (
                    <MenuItem key={s.id} onClick={() => chatRef.current?.openChat(s.id)} title={s.title}>
                      <span className={styles.sessionLabel}>{s.title}</span>
                      <span style={{ opacity: 0.7, fontSize: 11, marginLeft: 6 }}>
                        {new Date(s.updatedAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}
                      </span>
                    </MenuItem>
                  ))}
                </MenuList>
              </MenuPopover>
            </Menu>
            <Button
              appearance="transparent"
              className={styles.headerButton}
              icon={<AddCircle24Regular />}
              aria-label="New chat"
              title="New chat"
              onClick={() => chatRef.current?.newChat()}
            />
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <Button
                  appearance="transparent"
                  className={styles.headerButton}
                  icon={<MoreHorizontal24Regular />}
                  aria-label="More options"
                  title="More options"
                />
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<Settings24Regular />} onClick={() => setView('settings')}>Settings</MenuItem>
                  <MenuDivider />
                  <MenuItem icon={<ArrowDownload24Regular />} onClick={() => chatRef.current?.downloadTranscript()}>Download transcript</MenuItem>
                  <MenuItem icon={<ArrowDownload24Regular />} onClick={exportConfig}>Export config</MenuItem>
                  <MenuItem icon={<ArrowUpload24Regular />} onClick={() => importInputRef.current?.click()}>Import config</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        )}
      </div>

      <input
        ref={importInputRef}
        className={styles.hiddenInput}
        type="file"
        accept=".json,application/json"
        onChange={e => {
          const file = e.target.files?.[0];
          if (file) void importConfig(file);
          e.target.value = '';
        }}
      />

      <div className={styles.content}>
        {view === 'settings' ? (
          <Settings onBack={() => setView('chat')} />
        ) : (
          <Chat key={chatKey} ref={chatRef} onOpenSettings={() => setView('settings')} />
        )}
      </div>
    </div>
  );
}

export default App;
