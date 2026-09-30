import { useEffect, useRef, useState } from 'react';
import { Button, Spinner, Textarea, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { loadSettings } from '../utils/storage';
import { runAgentLoop, type AgentEvent, type ApprovalDecision, type ApprovalRequest } from '../agent/agentLoop';
import { canUndo, getSelectedRangeAddress, undoLastGroup } from '../agent/excelTools';
import type { ChatMessage } from '../agent/llmClient';

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
    maxWidth: '90%',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    fontSize: tokens.fontSizeBase300,
  },
  user: {
    alignSelf: 'flex-end',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  assistant: {
    alignSelf: 'flex-start',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
  },
  error: {
    alignSelf: 'stretch',
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    maxWidth: '100%',
  },
  info: {
    alignSelf: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  tool: {
    alignSelf: 'stretch',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  toolAwaiting: {
    border: `1px solid ${tokens.colorPaletteMarigoldBorder2}`,
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
  },
  toolLine: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  toolDetail: {
    color: tokens.colorNeutralForeground3,
    wordBreak: 'break-word',
  },
  toolError: {
    color: tokens.colorPaletteRedForeground1,
  },
  args: {
    maxHeight: '160px',
    overflow: 'auto',
    margin: 0,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    whiteSpace: 'pre',
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
  selectionLabel: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  inputRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'flex-end',
    padding: '8px',
  },
  input: {
    flex: 1,
  },
});

type ToolStatus = 'running' | 'awaiting_approval' | 'ok' | 'error' | 'rejected';

interface ToolItem {
  kind: 'tool';
  id: string;
  name: string;
  summary: string;
  status: ToolStatus;
  detail?: string;
  args?: string;
}

type ChatItem = { kind: 'user' | 'assistant' | 'error' | 'info'; id: string; text: string } | ToolItem;

const STATUS_ICON: Record<ToolStatus, string> = {
  running: '',
  awaiting_approval: '⚠',
  ok: '✓',
  error: '✗',
  rejected: '⊘',
};

let lastItemId = 0;
const newId = () => `item-${++lastItemId}`;

export const Chat: React.FC = () => {
  const styles = useStyles();
  const [input, setInput] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [undoAvailable, setUndoAvailable] = useState(canUndo);
  // Sheet-qualified address of the current selection, shown above the input and given to the agent.
  const [selection, setSelection] = useState<string | null>(null);
  const [selectionDismissed, setSelectionDismissed] = useState(false);
  // Conversation in API format, including tool calls and results. Kept apart from the
  // displayed items so UI-only messages never reach the model.
  const historyRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const approvalsRef = useRef(new Map<string, (decision: ApprovalDecision) => void>());
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ block: 'end' });
  }, [items]);

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

  const addItem = (item: ChatItem) => setItems(prev => [...prev, item]);
  const updateTool = (id: string | undefined, patch: Partial<ToolItem>) =>
    setItems(prev => prev.map(item => (item.kind === 'tool' && item.id === id ? { ...item, ...patch } : item)));

  const handleSend = async () => {
    const text = input.trim();
    if (!text || isRunning) return;
    setInput('');
    addItem({ kind: 'user', id: newId(), text });
    setIsRunning(true);

    const controller = new AbortController();
    abortRef.current = controller;
    // Providers may reuse tool call ids across runs, so they are mapped to item ids per run.
    const toolItems = new Map<string, string>();

    const onEvent = (event: AgentEvent) => {
      switch (event.type) {
        case 'assistant_text':
          addItem({ kind: 'assistant', id: newId(), text: event.text });
          break;
        case 'tool_call': {
          const id = newId();
          toolItems.set(event.callId, id);
          addItem({ kind: 'tool', id, name: event.name, summary: event.summary, status: 'running' });
          break;
        }
        case 'tool_result':
          updateTool(toolItems.get(event.callId), { status: event.status, detail: event.detail || undefined });
          break;
      }
    };

    const requestApproval = (request: ApprovalRequest) =>
      new Promise<ApprovalDecision>(resolve => {
        const id = toolItems.get(request.callId) ?? newId();
        updateTool(id, { status: 'awaiting_approval', args: JSON.stringify(request.args, null, 2) });
        approvalsRef.current.set(id, resolve);
      });

    try {
      const result = await runAgentLoop({
        settings: loadSettings(),
        history: historyRef.current,
        userInput: text,
        includeSelection: !selectionDismissed,
        onEvent,
        requestApproval,
        signal: controller.signal,
      });
      historyRef.current = result.history;
      if (result.status === 'error') addItem({ kind: 'error', id: newId(), text: result.error ?? 'Unknown error' });
      if (result.status === 'aborted') addItem({ kind: 'info', id: newId(), text: 'Stopped.' });
    } finally {
      setItems(prev => prev.map(item =>
        item.kind === 'tool' && (item.status === 'running' || item.status === 'awaiting_approval')
          ? { ...item, status: 'error', detail: 'Not executed.' }
          : item
      ));
      approvalsRef.current.clear();
      abortRef.current = null;
      setIsRunning(false);
      setUndoAvailable(canUndo());
    }
  };

  const decide = (id: string, decision: ApprovalDecision) => {
    const resolve = approvalsRef.current.get(id);
    if (!resolve) return;
    approvalsRef.current.delete(id);
    updateTool(id, { status: decision === 'reject' ? 'rejected' : 'running' });
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
      ];
      addItem({ kind: 'info', id: newId(), text: 'Last AI changes undone.' });
    } catch (e) {
      addItem({ kind: 'error', id: newId(), text: `Undo failed: ${e instanceof Error ? e.message : String(e)}` });
    } finally {
      setUndoAvailable(canUndo());
    }
  };

  const handleClear = () => {
    historyRef.current = [];
    setItems([]);
  };

  const renderTool = (item: ToolItem) => (
    <div key={item.id} className={mergeClasses(styles.tool, item.status === 'awaiting_approval' && styles.toolAwaiting)}>
      <div className={styles.toolLine}>
        {item.status === 'running' ? <Spinner size="extra-tiny" /> : <span>{STATUS_ICON[item.status]}</span>}
        <span>{item.summary}</span>
      </div>
      {item.detail && (
        <div className={mergeClasses(styles.toolDetail, item.status === 'error' && styles.toolError)}>{item.detail}</div>
      )}
      {item.status === 'awaiting_approval' && (
        <>
          {item.args && (
            <details>
              <summary>Details</summary>
              <pre className={styles.args}>{item.args}</pre>
            </details>
          )}
          <div className={styles.actions}>
            <Button size="small" appearance="primary" onClick={() => decide(item.id, 'approve')}>Approve</Button>
            <Button size="small" onClick={() => decide(item.id, 'approve_all')}>Approve all</Button>
            <Button size="small" onClick={() => decide(item.id, 'reject')}>Reject</Button>
          </div>
        </>
      )}
    </div>
  );

  const bubbleClass = { user: styles.user, assistant: styles.assistant, error: styles.error, info: styles.info };
  const awaitingApproval = items.some(item => item.kind === 'tool' && item.status === 'awaiting_approval');

  return (
    <div className={styles.container}>
      <div className={styles.history}>
        {items.length === 0 && (
          <div className={styles.empty}>
            Ask the agent to analyze or change your workbook, for example:
            <br />• "Create a Summary sheet with total sales per region using SUMIFS"
            <br />• "Turn the data in Sheet1 into a table and format the prices as currency"
            <br />• "Add a column that looks up each product's category with XLOOKUP"
          </div>
        )}
        {items.map(item =>
          item.kind === 'tool'
            ? renderTool(item)
            : <div key={item.id} className={mergeClasses(item.kind !== 'info' && styles.bubble, bubbleClass[item.kind])}>{item.text}</div>
        )}
        {isRunning && !awaitingApproval && <Spinner size="tiny" label="Working…" labelPosition="after" />}
        <div ref={bottomRef} />
      </div>

      <div className={styles.actions}>
        <Button size="small" onClick={handleUndo} disabled={!undoAvailable || isRunning}>Undo last AI changes</Button>
        <Button size="small" appearance="subtle" onClick={handleClear} disabled={isRunning || items.length === 0}>New chat</Button>
      </div>

      <div className={styles.composer}>
        {selection && !selectionDismissed && (
          <div className={styles.selectionChip}>
            <span className={styles.selectionLabel} title={selection}>
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
        <div className={styles.inputRow}>
          <Textarea
            className={styles.input}
            appearance="filled-lighter"
            value={input}
            onChange={(_, data) => setInput(data.value)}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault();
                void handleSend();
              }
            }}
            placeholder="Ask the agent… (Shift+Enter for a new line)"
            resize="vertical"
            disabled={isRunning}
          />
          {isRunning
            ? <Button onClick={handleStop}>Stop</Button>
            : <Button appearance="primary" onClick={() => void handleSend()} disabled={!input.trim()}>Send</Button>}
        </div>
      </div>
    </div>
  );
};
