import { useEffect, useRef, useState } from 'react';
import { Button, ProgressBar, Textarea, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { MAX_FAVORITES, applyFavorite, loadSettings, saveSettings, activeModelLabel, type AppSettings } from '../utils/storage';
import { debounce, getWorkbookId, kvDelete, kvGet, kvSet } from '../utils/persistence';
import { runAgentLoop, type AgentEvent, type ApprovalDecision, type ApprovalRequest } from '../agent/agentLoop';
import {
  canUndo,
  cleanupOrphanBackups,
  getSelectedRangeAddress,
  initUndo,
  selectReference,
  undoLastGroup,
  type UndoStep,
} from '../agent/excel';
import { readAttachment, type Attachment } from '../agent/attachments';
import { stripBinaryContent } from '../agent/context';
import type { ChatMessage } from '../agent/llmClient';
import { findSkill, getSkills, type Skill } from '../agent/skills';
import { Markdown } from './Markdown';
import { ChangesCard, ToolCard } from './ToolCard';
import { EffortMenu, ModelMenu, SkillsMenu } from './ComposerControls';
import type { ChangesItem, ChatItem, ToolItem } from './chatTypes';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    padding: '12px',
    boxSizing: 'border-box',
    gap: '8px',
  },
  history: {
    flex: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  empty: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
  },
  bubble: {
    padding: '8px 12px',
    borderRadius: tokens.borderRadiusLarge,
    maxWidth: '92%',
    wordBreak: 'break-word',
    fontSize: tokens.fontSizeBase300,
  },
  user: {
    alignSelf: 'flex-end',
    whiteSpace: 'pre-wrap',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  userAttachments: {
    fontSize: tokens.fontSizeBase200,
    opacity: 0.85,
    marginTop: '4px',
  },
  assistant: {
    alignSelf: 'flex-start',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
  },
  error: {
    alignSelf: 'stretch',
    whiteSpace: 'pre-wrap',
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    maxWidth: '100%',
  },
  info: {
    alignSelf: 'center',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
  },
  composer: {
    display: 'flex',
    flexDirection: 'column',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusXLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: 'hidden',
  },
  chipRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '4px',
    padding: '6px 8px 0',
  },
  chip: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    maxWidth: '100%',
    padding: '0 2px 0 8px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  chipLabel: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  warning: {
    padding: '4px 10px 0',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteDarkOrangeForeground1,
  },
  selectionChip: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
    padding: '4px 4px 4px 12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  inputRow: {
    display: 'flex',
    gap: '6px',
    alignItems: 'flex-end',
    padding: '8px 8px 4px',
  },
  input: {
    flex: 1,
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '2px',
    padding: '0 4px 4px',
  },
  spacer: {
    flex: 1,
  },
  hiddenInput: {
    display: 'none',
  },
  activity: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  activityAwaiting: {
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
    color: tokens.colorPaletteMarigoldForeground2,
  },
  activityLine: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  activityText: {
    flex: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  pulse: {
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    flexShrink: 0,
    backgroundColor: 'currentColor',
    animationName: {
      '0%': { opacity: 1, transform: 'scale(1)' },
      '50%': { opacity: 0.3, transform: 'scale(0.7)' },
      '100%': { opacity: 1, transform: 'scale(1)' },
    },
    animationDuration: '1.2s',
    animationIterationCount: 'infinite',
  },
  elapsed: {
    fontVariantNumeric: 'tabular-nums',
    opacity: 0.8,
  },
  thinking: {
    alignSelf: 'stretch',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    '& summary': { cursor: 'pointer' },
  },
  thinkingText: {
    whiteSpace: 'pre-wrap',
    maxHeight: '200px',
    overflowY: 'auto',
    margin: '4px 0 0',
    paddingLeft: '8px',
    borderLeft: `2px solid ${tokens.colorNeutralStroke2}`,
  },
});

/** Splits leading "/skill" commands off the message: "/dashboard sales by region". */
const parseSkillCommands = (text: string, settings: AppSettings): { text: string; skills: Skill[] } => {
  const skills: Skill[] = [];
  let rest = text;
  for (let match = /^\/(\S+)\s*/.exec(rest); match; match = /^\/(\S+)\s*/.exec(rest)) {
    const skill = findSkill(settings, match[1]);
    if (!skill) break;
    if (!skills.includes(skill)) skills.push(skill);
    rest = rest.slice(match[0].length);
  }
  return { text: rest.trim() || text, skills };
};

interface StoredChat {
  items: ChatItem[];
  history: ChatMessage[];
}

const ATTACH_ACCEPT = '.csv,.tsv,.txt,.json,.md,.pdf,image/*';

const saveChat = debounce(({ key, chat }: { key: string; chat: StoredChat }) => void kvSet(`chat:${key}`, chat));

let lastItemId = 0;
const newId = () => `${Date.now().toString(36)}-${++lastItemId}`;

const formatElapsed = (ms: number) => {
  const seconds = Math.max(0, Math.floor(ms / 1000));
  return seconds < 60 ? `${seconds}s` : `${Math.floor(seconds / 60)}m ${String(seconds % 60).padStart(2, '0')}s`;
};

const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));

/** Items restored from storage can't still be running. */
const settleItems = (items: ChatItem[]): ChatItem[] =>
  items.map(item => {
    if ((item.kind === 'assistant' || item.kind === 'thinking') && item.streaming) return { ...item, streaming: false };
    if (item.kind === 'tool' && (item.status === 'running' || item.status === 'awaiting_approval')) {
      return { ...item, status: 'error', detail: 'Not executed.', args: undefined, preview: undefined };
    }
    return item;
  });

interface ChatProps {
  onOpenSettings: () => void;
}

export const Chat: React.FC<ChatProps> = ({ onOpenSettings }) => {
  const styles = useStyles();
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [selectedSkills, setSelectedSkills] = useState<string[]>([]);
  // What the agent is doing right now, shown in the activity bar while it runs.
  const [activity, setActivity] = useState('');
  const [startedAt, setStartedAt] = useState(0);
  const [now, setNow] = useState(0);
  const [input, setInput] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [status, setStatus] = useState('');
  const [undoAvailable, setUndoAvailable] = useState(false);
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  // Sheet-qualified address of the current selection, shown above the input and given to the agent.
  const [selection, setSelection] = useState<string | null>(null);
  const [selectionDismissed, setSelectionDismissed] = useState(false);
  // Conversation in API format, including tool calls and results. Kept apart from the
  // displayed items so UI-only messages never reach the model.
  const historyRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const approvalsRef = useRef(new Map<string, (decision: ApprovalDecision) => void>());
  const bottomRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const storageKeyRef = useRef<string | null>(null);

  // Restore this workbook's conversation and undo history.
  useEffect(() => {
    let active = true;
    void (async () => {
      const id = await getWorkbookId();
      const [chat, undo] = await Promise.all([kvGet<StoredChat>(`chat:${id}`), kvGet<UndoStep[][]>(`undo:${id}`)]);
      if (!active) return;
      storageKeyRef.current = id;
      if (chat) {
        historyRef.current = chat.history ?? [];
        setItems(settleItems(chat.items ?? []));
      }
      initUndo(undo ?? [], stack => void kvSet(`undo:${id}`, stack));
      setUndoAvailable(canUndo());
      if (typeof Excel !== 'undefined') void cleanupOrphanBackups();
    })();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    // Nothing is saved until the stored conversation has been restored.
    if (storageKeyRef.current) saveChat({ key: storageKeyRef.current, chat: { items, history: stripBinaryContent(historyRef.current) } });
  }, [items]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ block: 'end' });
  }, [items, status]);

  useEffect(() => {
    // Not running inside Excel (e.g. the page opened in a regular browser).
    if (!Office.context?.document) return;
    let active = true;
    const refresh = () => {
      getSelectedRangeAddress()
        .then(address => {
          if (!active) return;
          setSelection(address);
          // A new selection brings the chip back after it was dismissed.
          setSelectionDismissed(false);
        })
        .catch(e => console.error('Failed to read the selection', e));
    };
    refresh();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refresh);
    return () => {
      active = false;
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: refresh });
    };
  }, []);

  // Ticks once per second while running so the elapsed time shows the agent is alive.
  useEffect(() => {
    if (!isRunning) return;
    const timer = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(timer);
  }, [isRunning]);

  const addItem = (item: ChatItem) => setItems(prev => [...prev, item]);
  const updateItem = (id: string | undefined | null, patch: Partial<ChatItem>) =>
    setItems(prev => prev.map(item => (item.id === id ? ({ ...item, ...patch } as ChatItem) : item)));

  const goToCell = (reference: string) => {
    selectReference(reference).catch(e => addItem({ kind: 'error', id: newId(), text: `Couldn't go to ${reference}: ${errorText(e)}` }));
  };

  const addFiles = async (files: File[]) => {
    for (const file of files) {
      try {
        const attachment = await readAttachment(file);
        setAttachments(prev => [...prev, attachment]);
      } catch (e) {
        addItem({ kind: 'error', id: newId(), text: `Couldn't attach ${file.name || 'the file'}: ${errorText(e)}` });
      }
    }
  };

  const updateSettings = (next: AppSettings) => {
    setSettings(next);
    saveSettings(next);
  };

  const addCurrentToFavorites = () => {
    const config = settings.providers[settings.provider];
    if (settings.favorites.length >= MAX_FAVORITES || !config.model) return;
    updateSettings({
      ...settings,
      favorites: [...settings.favorites, { provider: settings.provider, model: config.model, label: activeModelLabel(settings), modelInfo: config.modelInfo }],
    });
  };

  const handleSend = async () => {
    const raw = input.trim();
    if (!raw || isRunning) return;
    const command = parseSkillCommands(raw, settings);
    const text = command.text;
    const skills = [...new Map([
      ...selectedSkills.flatMap(id => findSkill(settings, id) ?? []),
      ...command.skills,
    ].map(s => [s.id, s])).values()];
    const sentAttachments = attachments;
    setInput('');
    setAttachments([]);
    setSelectedSkills([]);
    addItem({
      kind: 'user',
      id: newId(),
      text,
      ...(sentAttachments.length ? { attachments: sentAttachments.map(a => a.name) } : {}),
      ...(skills.length ? { skills: skills.map(s => s.name) } : {}),
    });
    setIsRunning(true);
    setActivity('Reading the workbook…');
    setStartedAt(Date.now());
    setNow(Date.now());

    const controller = new AbortController();
    abortRef.current = controller;
    // Providers may reuse tool call ids across runs, so they are mapped to item ids per run.
    const toolItems = new Map<string, string>();
    const toolInfo = new Map<string, { summary: string; mutating: boolean }>();
    const changes: ChangesItem['entries'] = [];
    let streamingId: string | null = null;
    let thinkingId: string | null = null;

    const onEvent = (event: AgentEvent) => {
      if (event.type !== 'status') setStatus('');
      // Any output other than more reasoning ends the current thinking block.
      if (event.type !== 'reasoning_delta' && event.type !== 'status' && thinkingId) {
        updateItem(thinkingId, { streaming: false });
        thinkingId = null;
      }
      if (event.type === 'reasoning_delta') setActivity('Thinking…');
      if (event.type === 'assistant_delta') setActivity('Writing the answer…');
      if (event.type === 'tool_call') setActivity(`${event.summary}…`);
      if (event.type === 'tool_result') setActivity('Thinking about the next step…');
      switch (event.type) {
        case 'reasoning_delta':
          if (thinkingId) {
            const id = thinkingId;
            setItems(prev => prev.map(item => (item.id === id && item.kind === 'thinking' ? { ...item, text: item.text + event.delta } : item)));
          } else {
            thinkingId = newId();
            addItem({ kind: 'thinking', id: thinkingId, text: event.delta, streaming: true });
          }
          break;
        case 'assistant_delta':
          if (streamingId) {
            const id = streamingId;
            setItems(prev => prev.map(item => (item.id === id && item.kind === 'assistant' ? { ...item, text: item.text + event.delta } : item)));
          } else {
            streamingId = newId();
            addItem({ kind: 'assistant', id: streamingId, text: event.delta, streaming: true });
          }
          break;
        case 'assistant_text':
          if (streamingId) updateItem(streamingId, { text: event.text, streaming: false });
          else addItem({ kind: 'assistant', id: newId(), text: event.text });
          streamingId = null;
          break;
        case 'tool_call': {
          const id = newId();
          toolItems.set(event.callId, id);
          toolInfo.set(event.callId, { summary: event.summary, mutating: event.mutating });
          addItem({ kind: 'tool', id, name: event.name, summary: event.summary, status: 'running', mutating: event.mutating });
          break;
        }
        case 'tool_result': {
          updateItem(toolItems.get(event.callId), { status: event.status, detail: event.detail || undefined, location: event.location, args: undefined, preview: undefined });
          const info = toolInfo.get(event.callId);
          if (event.status === 'ok' && info?.mutating) changes.push({ summary: info.summary, location: event.location });
          break;
        }
        case 'status':
          setStatus(event.text);
          break;
        case 'info':
          addItem({ kind: 'info', id: newId(), text: event.text });
          break;
      }
    };

    const requestApproval = (request: ApprovalRequest) =>
      new Promise<ApprovalDecision>(resolve => {
        const id = toolItems.get(request.callId) ?? newId();
        updateItem(id, { status: 'awaiting_approval', args: request.args, preview: request.preview, irreversible: request.irreversible } as Partial<ToolItem>);
        setActivity('Waiting for your approval…');
        approvalsRef.current.set(id, resolve);
      });

    try {
      const result = await runAgentLoop({
        settings,
        history: historyRef.current,
        userInput: text,
        attachments: sentAttachments,
        skills,
        includeSelection: !selectionDismissed,
        onEvent,
        requestApproval,
        signal: controller.signal,
      });
      historyRef.current = result.history;
      if (changes.length > 1) addItem({ kind: 'changes', id: newId(), entries: changes });
      if (result.status === 'error') addItem({ kind: 'error', id: newId(), text: result.error ?? 'Unknown error' });
      if (result.status === 'aborted') addItem({ kind: 'info', id: newId(), text: 'Stopped.' });
    } finally {
      setItems(prev => settleItems(prev));
      approvalsRef.current.clear();
      abortRef.current = null;
      setStatus('');
      setIsRunning(false);
      setUndoAvailable(canUndo());
    }
  };

  const decide = (id: string, decision: ApprovalDecision) => {
    const resolve = approvalsRef.current.get(id);
    if (!resolve) return;
    approvalsRef.current.delete(id);
    updateItem(id, { status: decision === 'reject' ? 'rejected' : 'running' });
    setActivity(decision === 'reject' ? 'Thinking about the next step…' : 'Applying the change…');
    resolve(decision);
  };

  const handleStop = () => {
    abortRef.current?.abort();
    approvalsRef.current.forEach(resolve => resolve('reject'));
    approvalsRef.current.clear();
  };

  const handleUndo = async () => {
    try {
      await undoLastGroup();
      historyRef.current = [
        ...historyRef.current,
        { role: 'user', content: '[Note: the user undid the most recent set of workbook changes you made.]' },
        { role: 'assistant', content: 'Noted: those changes were undone.' },
      ];
      addItem({ kind: 'info', id: newId(), text: 'Last AI changes undone.' });
    } catch (e) {
      addItem({ kind: 'error', id: newId(), text: `Undo failed: ${errorText(e)}` });
    } finally {
      setUndoAvailable(canUndo());
    }
  };

  const handleClear = () => {
    historyRef.current = [];
    setItems([]);
    if (storageKeyRef.current) void kvDelete(`chat:${storageKeyRef.current}`);
  };

  const handlePaste = (e: React.ClipboardEvent) => {
    const images = [...e.clipboardData.files].filter(f => f.type.startsWith('image/'));
    if (images.length) {
      e.preventDefault();
      void addFiles(images);
    }
  };

  const awaitingApproval = items.some(item => item.kind === 'tool' && item.status === 'awaiting_approval');
  const activeConfig = settings.providers[settings.provider];
  const modelInfo = activeConfig.modelInfo?.id === activeConfig.model ? activeConfig.modelInfo : undefined;
  const allSkills = getSkills(settings);
  const attachmentWarning =
    attachments.some(a => a.kind === 'image') && modelInfo?.supportsImages === false ? 'The selected model does not accept images.'
    : attachments.some(a => a.kind === 'pdf') && modelInfo?.supportsFiles === false && modelInfo?.supportsImages === false ? 'The selected model may not accept PDF files.'
    : '';

  const renderItem = (item: ChatItem) => {
    switch (item.kind) {
      case 'tool':
        return <ToolCard key={item.id} item={item} onDecide={decide} onCellClick={goToCell} />;
      case 'changes':
        return <ChangesCard key={item.id} item={item} onCellClick={goToCell} />;
      case 'assistant':
        return (
          <div key={item.id} className={mergeClasses(styles.bubble, styles.assistant)}>
            <Markdown text={item.text} onCellClick={goToCell} />
          </div>
        );
      case 'thinking':
        return (
          <details key={item.id} className={styles.thinking} open={item.streaming}>
            <summary>{item.streaming ? '💭 Thinking…' : '💭 Thought process'}</summary>
            <div className={styles.thinkingText}>{item.text}</div>
          </details>
        );
      case 'user':
        return (
          <div key={item.id} className={mergeClasses(styles.bubble, styles.user)}>
            {item.text}
            {item.skills && <div className={styles.userAttachments}>🧩 {item.skills.join(', ')}</div>}
            {item.attachments && <div className={styles.userAttachments}>📎 {item.attachments.join(', ')}</div>}
          </div>
        );
      case 'error':
        return <div key={item.id} className={mergeClasses(styles.bubble, styles.error)}>{item.text}</div>;
      case 'info':
        return <div key={item.id} className={styles.info}>{item.text}</div>;
    }
  };


  return (
    <div className={styles.container}>
      <div className={styles.history}>
        {items.length === 0 && (
          <div className={styles.empty}>
            Ask the agent to analyze or change your workbook, for example:
            <br />• "Create a Summary sheet with total sales per region using SUMIFS"
            <br />• "Make a PivotTable and a chart of revenue by product"
            <br />• "Why does F20 show #N/A?"
            <br />• "Import the attached CSV as a table and highlight values over 1000"
          </div>
        )}
        {items.map(renderItem)}
        <div ref={bottomRef} />
      </div>

      <div className={styles.actions}>
        <Button size="small" onClick={handleUndo} disabled={!undoAvailable || isRunning}>Undo last AI changes</Button>
        <Button size="small" appearance="subtle" onClick={handleClear} disabled={isRunning || items.length === 0}>New chat</Button>
      </div>

      {isRunning && (
        <div className={mergeClasses(styles.activity, awaitingApproval && styles.activityAwaiting)} role="status" aria-live="polite">
          <div className={styles.activityLine}>
            <span className={styles.pulse} aria-hidden />
            <span className={styles.activityText} title={status || activity}>{status || activity || 'Working…'}</span>
            <span className={styles.elapsed}>{formatElapsed(now - startedAt)}</span>
          </div>
          {!awaitingApproval && <ProgressBar thickness="medium" shape="rounded" />}
        </div>
      )}

      <div className={styles.composer}>
        {selection && !selectionDismissed && (
          <div className={styles.selectionChip}>
            <span className={styles.chipLabel} title={selection}>
              {selection.slice(selection.lastIndexOf('!') + 1)} selected
            </span>
            <Button
              size="small"
              appearance="transparent"
              icon={<span aria-hidden>✕</span>}
              aria-label="Don't include the selection"
              title="Don't include the selection"
              onClick={() => setSelectionDismissed(true)}
            />
          </div>
        )}
        {(attachments.length > 0 || selectedSkills.length > 0) && (
          <div className={styles.chipRow}>
            {selectedSkills.map(id => {
              const skill = allSkills.find(s => s.id === id);
              return (
                <span key={id} className={styles.chip}>
                  <span className={styles.chipLabel} title={skill?.description}>🧩 {skill?.name ?? id}</span>
                  <Button
                    size="small"
                    appearance="transparent"
                    icon={<span aria-hidden>✕</span>}
                    aria-label={`Remove skill ${skill?.name ?? id}`}
                    onClick={() => setSelectedSkills(prev => prev.filter(x => x !== id))}
                  />
                </span>
              );
            })}
            {attachments.map(a => (
              <span key={a.id} className={styles.chip}>
                <span className={styles.chipLabel} title={a.name}>📎 {a.name}</span>
                <Button
                  size="small"
                  appearance="transparent"
                  icon={<span aria-hidden>✕</span>}
                  aria-label={`Remove ${a.name}`}
                  onClick={() => setAttachments(prev => prev.filter(x => x.id !== a.id))}
                />
              </span>
            ))}
          </div>
        )}
        {attachmentWarning && <div className={styles.warning}>{attachmentWarning}</div>}
        <div className={styles.inputRow}>
          <input
            ref={fileInputRef}
            className={styles.hiddenInput}
            type="file"
            multiple
            accept={ATTACH_ACCEPT}
            onChange={e => {
              void addFiles([...(e.target.files ?? [])]);
              e.target.value = '';
            }}
          />
          <Textarea
            className={styles.input}
            appearance="filled-lighter"
            value={input}
            onChange={(_, data) => setInput(data.value)}
            onPaste={handlePaste}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault();
                void handleSend();
              }
            }}
            placeholder="Ask the agent… (/skill to use a skill, Shift+Enter for a new line)"
            resize="vertical"
            disabled={isRunning}
          />
          {isRunning
            ? <Button onClick={handleStop}>Stop</Button>
            : <Button appearance="primary" onClick={() => void handleSend()} disabled={!input.trim()}>Send</Button>}
        </div>
        <div className={styles.toolbar}>
          <Button
            size="small"
            appearance="subtle"
            icon={<span aria-hidden>📎</span>}
            aria-label="Attach files"
            title="Attach CSV, JSON, text, PDF or images"
            disabled={isRunning}
            onClick={() => fileInputRef.current?.click()}
          />
          <SkillsMenu skills={allSkills} selected={selectedSkills} disabled={isRunning} onChange={setSelectedSkills} onManage={onOpenSettings} />
          <div className={styles.spacer} />
          <ModelMenu
            settings={settings}
            disabled={isRunning}
            onSelect={favorite => updateSettings(applyFavorite(settings, favorite))}
            onAddCurrent={addCurrentToFavorites}
            onManage={onOpenSettings}
          />
          <EffortMenu settings={settings} disabled={isRunning} onChange={effort => updateSettings({ ...settings, reasoningEffort: effort })} />
        </div>
      </div>
    </div>
  );
};
