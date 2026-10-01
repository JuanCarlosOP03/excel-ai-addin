import ReactMarkdown, { defaultUrlTransform, type Components } from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { makeStyles, tokens } from '@fluentui/react-components';
import { CELL_LINK_PREFIX, linkifyCellReferences } from '../utils/cellReferences';

const useStyles = makeStyles({
  root: {
    '& p': { margin: '0 0 6px' },
    '& p:last-child': { marginBottom: 0 },
    '& ul, & ol': { margin: '0 0 6px', paddingLeft: '20px' },
    '& li': { marginBottom: '2px' },
    '& h1, & h2, & h3, & h4': { margin: '8px 0 4px', fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
    '& code': {
      fontFamily: tokens.fontFamilyMonospace,
      fontSize: tokens.fontSizeBase200,
      backgroundColor: tokens.colorNeutralBackground1,
      padding: '0 3px',
      borderRadius: tokens.borderRadiusSmall,
    },
    '& pre': {
      overflowX: 'auto',
      padding: '6px 8px',
      backgroundColor: tokens.colorNeutralBackground1,
      borderRadius: tokens.borderRadiusMedium,
      margin: '0 0 6px',
    },
    '& pre code': { padding: 0, backgroundColor: 'transparent' },
    '& table': { borderCollapse: 'collapse', margin: '0 0 6px', fontSize: tokens.fontSizeBase200, display: 'block', overflowX: 'auto' },
    '& th, & td': { border: `1px solid ${tokens.colorNeutralStroke2}`, padding: '2px 6px', textAlign: 'left' },
    '& th': { backgroundColor: tokens.colorNeutralBackground1, fontWeight: tokens.fontWeightSemibold },
  },
  cellLink: {
    color: tokens.colorBrandForegroundLink,
    textDecorationLine: 'underline',
    textDecorationStyle: 'dotted',
    cursor: 'pointer',
    background: 'none',
    border: 'none',
    padding: 0,
    font: 'inherit',
    ':hover': { color: tokens.colorBrandForegroundLinkHover },
  },
});

interface MarkdownProps {
  text: string;
  onCellClick: (reference: string) => void;
}

export const Markdown: React.FC<MarkdownProps> = ({ text, onCellClick }) => {
  const styles = useStyles();
  const components: Components = {
    a: ({ href, children }) => {
      if (href?.startsWith(CELL_LINK_PREFIX)) {
        const reference = decodeURIComponent(href.slice(CELL_LINK_PREFIX.length));
        return (
          <button type="button" className={styles.cellLink} title={`Go to ${reference}`} onClick={() => onCellClick(reference)}>
            {children}
          </button>
        );
      }
      return <a href={href} target="_blank" rel="noopener noreferrer">{children}</a>;
    },
  };

  return (
    <div className={styles.root}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={components}
        urlTransform={url => (url.startsWith(CELL_LINK_PREFIX) ? url : defaultUrlTransform(url))}
      >
        {linkifyCellReferences(text)}
      </ReactMarkdown>
    </div>
  );
};
