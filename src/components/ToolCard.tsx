import { Button, Spinner, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import type { ApprovalDecision } from '../agent/agentLoop';
import type { GridPreview } from '../agent/excel';
import { quoteSheetName } from '../utils/cellReferences';
import type { ChangesItem, ToolItem, ToolStatus } from './chatTypes';

const useStyles = makeStyles({
  card: {
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
  awaiting: {
    border: `1px solid ${tokens.colorPaletteMarigoldBorder2}`,
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
  },
  line: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  detail: {
    color: tokens.colorNeutralForeground3,
    wordBreak: 'break-word',
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
  },
  warning: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  link: {
    color: tokens.colorBrandForegroundLink,
    background: 'none',
    border: 'none',
    padding: 0,
    font: 'inherit',
    textAlign: 'left',
    cursor: 'pointer',
    textDecorationLine: 'underline',
    textDecorationStyle: 'dotted',
  },
  gridLabel: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: '2px',
  },
  gridWrap: {
    overflowX: 'auto',
  },
  grid: {
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase100,
    fontFamily: tokens.fontFamilyMonospace,
    '& td': {
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      padding: '1px 4px',
      maxWidth: '110px',
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      whiteSpace: 'nowrap',
      backgroundColor: tokens.colorNeutralBackground1,
    },
  },
  changed: {
    '&&': { backgroundColor: tokens.colorPaletteLightGreenBackground2 },
  },
  params: {
    margin: 0,
    paddingLeft: '14px',
    wordBreak: 'break-word',
  },
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
  },
  changesTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
});

const STATUS_ICON: Record<ToolStatus, string> = {
  running: '',
  awaiting_approval: '⚠',
  ok: '✓',
  error: '✗',
  rejected: '⊘',
};

const display = (value: unknown) => (value === null || value === undefined ? '' : String(value));

const shorten = (value: unknown): string => {
  const text = typeof value === 'string' ? value : JSON.stringify(value);
  return text.length > 140 ? `${text.slice(0, 140)}…` : text;
};

const Grid: React.FC<{ rows: unknown[][]; compare?: unknown[][] }> = ({ rows, compare }) => {
  const styles = useStyles();
  return (
    <div className={styles.gridWrap}>
      <table className={styles.grid}>
        <tbody>
          {rows.map((row, r) => (
            <tr key={r}>
              {row.map((cell, c) => (
                <td key={c} title={display(cell)} className={compare && display(compare[r]?.[c]) !== display(cell) ? styles.changed : undefined}>
                  {display(cell)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

const Preview: React.FC<{ preview: GridPreview }> = ({ preview }) => {
  const styles = useStyles();
  const clipped = preview.rows > preview.before.length || preview.columns > (preview.before[0]?.length ?? 0);
  return (
    <>
      <div className={styles.gridLabel}>
        {quoteSheetName(preview.sheet)}!{preview.address} ({preview.rows}×{preview.columns}{clipped ? ', first cells shown' : ''})
      </div>
      <div>Before</div>
      <Grid rows={preview.before} />
      {preview.after ? (
        <>
          <div>After</div>
          <Grid rows={preview.after} compare={preview.before} />
        </>
      ) : (
        <div className={styles.warning}>These cells will be removed.</div>
      )}
    </>
  );
};

interface LinkProps {
  location?: { sheet: string; address: string };
  text: string;
  onCellClick: (reference: string) => void;
}

const LocationLink: React.FC<LinkProps> = ({ location, text, onCellClick }) => {
  const styles = useStyles();
  if (!location) return <span>{text}</span>;
  const reference = `${quoteSheetName(location.sheet)}!${location.address}`;
  return (
    <button type="button" className={styles.link} title={`Go to ${reference}`} onClick={() => onCellClick(reference)}>
      {text}
    </button>
  );
};

interface ToolCardProps {
  item: ToolItem;
  onDecide: (id: string, decision: ApprovalDecision) => void;
  onCellClick: (reference: string) => void;
}

export const ToolCard: React.FC<ToolCardProps> = ({ item, onDecide, onCellClick }) => {
  const styles = useStyles();
  const awaiting = item.status === 'awaiting_approval';
  return (
    <div className={mergeClasses(styles.card, awaiting && styles.awaiting)}>
      <div className={styles.line}>
        {item.status === 'running' ? <Spinner size="extra-tiny" /> : <span>{STATUS_ICON[item.status]}</span>}
        <LocationLink location={item.status === 'ok' ? item.location : undefined} text={item.summary} onCellClick={onCellClick} />
      </div>
      {item.detail && <div className={mergeClasses(styles.detail, item.status === 'error' && styles.error)}>{item.detail}</div>}
      {awaiting && (
        <>
          {item.irreversible && <div className={styles.warning}>This action cannot be undone.</div>}
          {item.preview && <Preview preview={item.preview} />}
          {item.args && (
            <details>
              <summary>Parameters</summary>
              <ul className={styles.params}>
                {Object.entries(item.args).map(([key, value]) => <li key={key}><b>{key}</b>: {shorten(value)}</li>)}
              </ul>
            </details>
          )}
          <div className={styles.actions}>
            <Button size="small" appearance="primary" onClick={() => onDecide(item.id, 'approve')}>Approve</Button>
            {!item.irreversible && <Button size="small" onClick={() => onDecide(item.id, 'approve_all')}>Approve all</Button>}
            <Button size="small" onClick={() => onDecide(item.id, 'reject')}>Reject</Button>
          </div>
        </>
      )}
    </div>
  );
};

export const ChangesCard: React.FC<{ item: ChangesItem; onCellClick: (reference: string) => void }> = ({ item, onCellClick }) => {
  const styles = useStyles();
  return (
    <div className={styles.card}>
      <div className={styles.changesTitle}>{item.entries.length} changes made</div>
      <ul className={styles.params}>
        {item.entries.map((entry, i) => (
          <li key={i}><LocationLink location={entry.location} text={entry.summary} onCellClick={onCellClick} /></li>
        ))}
      </ul>
    </div>
  );
};
