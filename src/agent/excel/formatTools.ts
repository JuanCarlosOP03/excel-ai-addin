import {
  ToolError,
  localAddress,
  optionalBoolean,
  optionalEnum,
  optionalNumber,
  optionalString,
  requireString,
  resolveRange,
  type Args,
} from './common';
import { recordUndo, snapshotRange } from './undo';

const MAX_NUMBER_FORMAT_CELLS = 200000;

const ALIGNMENTS = ['Left', 'Center', 'Right'] as const;
const BORDER_OPTIONS = ['all', 'outline', 'none'] as const;
const OUTLINE_EDGES = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'] as const;
const INSIDE_EDGES = ['InsideHorizontal', 'InsideVertical'] as const;

const isNoColor = (color: string) => ['none', 'transparent', 'no fill'].includes(color.toLowerCase());

export const formatRange = async (context: Excel.RequestContext, args: Args) => {
  const options = {
    bold: optionalBoolean(args, 'font_bold'),
    italic: optionalBoolean(args, 'font_italic'),
    fontColor: optionalString(args, 'font_color'),
    fontSize: optionalNumber(args, 'font_size'),
    fillColor: optionalString(args, 'fill_color'),
    numberFormat: optionalString(args, 'num_format'),
    alignment: optionalEnum(args, 'horizontal_alignment', ALIGNMENTS),
    wrapText: optionalBoolean(args, 'wrap_text'),
    borders: optionalEnum(args, 'borders', BORDER_OPTIONS),
    borderColor: optionalString(args, 'border_color'),
    autofit: optionalBoolean(args, 'autofit_columns'),
  };
  const applied = Object.entries(options).filter(([, value]) => value !== undefined).map(([key]) => key);
  if (applied.length === 0) throw new ToolError('Provide at least one formatting property.');

  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  if (options.numberFormat && range.cellCount > MAX_NUMBER_FORMAT_CELLS) {
    throw new ToolError(`"num_format" can be applied to at most ${MAX_NUMBER_FORMAT_CELLS} cells at once, but the range has ${range.cellCount}. Limit it to the rows that contain data.`);
  }

  const snapshot = await snapshotRange(context, sheetName, range);
  const { format } = range;
  if (options.bold !== undefined) format.font.bold = options.bold;
  if (options.italic !== undefined) format.font.italic = options.italic;
  if (options.fontColor) format.font.color = options.fontColor;
  if (options.fontSize) format.font.size = options.fontSize;
  if (options.fillColor) {
    if (isNoColor(options.fillColor)) format.fill.clear();
    else format.fill.color = options.fillColor;
  }
  if (options.alignment) format.horizontalAlignment = options.alignment;
  if (options.wrapText !== undefined) format.wrapText = options.wrapText;
  if (options.borders) {
    const edges = options.borders === 'outline' ? OUTLINE_EDGES : [...OUTLINE_EDGES, ...INSIDE_EDGES];
    for (const edge of edges) {
      const border = format.borders.getItem(edge);
      if (options.borders === 'none') {
        border.style = 'None';
      } else {
        border.style = 'Continuous';
        border.weight = 'Thin';
        border.color = options.borderColor ?? '#000000';
      }
    }
  }
  if (options.numberFormat) {
    range.numberFormat = Array.from({ length: range.rowCount }, () => new Array<string>(range.columnCount).fill(options.numberFormat!));
  }
  if (options.autofit) format.autofitColumns();
  await context.sync();

  return {
    sheet: sheetName,
    address: localAddress(range.address),
    applied,
    undoAvailable: snapshot !== null,
    ...(snapshot ? {} : { note: 'The range was too large to snapshot, so this change cannot be undone.' }),
  };
};

const CONDITIONAL_TYPES = ['color_scale', 'data_bar', 'icon_set', 'cell_value', 'text_contains', 'top_bottom', 'formula'] as const;
const CELL_OPERATORS = ['GreaterThan', 'LessThan', 'Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThanOrEqual', 'LessThanOrEqual'] as const;
const TOP_BOTTOM_TYPES = ['TopItems', 'BottomItems', 'TopPercent', 'BottomPercent'] as const;

/** Applies the optional highlight format shared by rule-based conditional formats. */
const applyHighlight = (format: Excel.ConditionalRangeFormat, args: Args) => {
  const fill = optionalString(args, 'fill_color');
  const font = optionalString(args, 'font_color');
  const bold = optionalBoolean(args, 'font_bold');
  if (!fill && !font && bold === undefined) {
    // Default: light red fill with dark red text, like Excel's preset.
    format.fill.color = '#FFC7CE';
    format.font.color = '#9C0006';
    return;
  }
  if (fill) format.fill.color = fill;
  if (font) format.font.color = font;
  if (bold !== undefined) format.font.bold = bold;
};

export const addConditionalFormat = async (context: Excel.RequestContext, args: Args) => {
  const type = optionalEnum(args, 'type', CONDITIONAL_TYPES);
  if (!type) throw new ToolError(`"type" is required: ${CONDITIONAL_TYPES.join(', ')}.`);
  const { sheetName, range } = await resolveRange(context, args, 'range_address');
  const formats = range.conditionalFormats;
  let conditional: Excel.ConditionalFormat;

  switch (type) {
    case 'color_scale': {
      conditional = formats.add('ColorScale');
      const mid = optionalString(args, 'mid_color');
      conditional.colorScale.criteria = {
        minimum: { type: 'LowestValue', color: optionalString(args, 'min_color') ?? '#F8696B' },
        ...(mid ? { midpoint: { formula: '50', type: 'Percentile', color: mid } } : {}),
        maximum: { type: 'HighestValue', color: optionalString(args, 'max_color') ?? '#63BE7B' },
      };
      break;
    }
    case 'data_bar':
      conditional = formats.add('DataBar');
      conditional.dataBar.positiveFormat.fillColor = optionalString(args, 'bar_color') ?? '#638EC6';
      break;
    case 'icon_set':
      conditional = formats.add('IconSet');
      conditional.iconSet.style = (optionalString(args, 'icon_style') ?? 'ThreeTrafficLights1') as Excel.IconSet;
      break;
    case 'cell_value': {
      const operator = optionalEnum(args, 'operator', CELL_OPERATORS);
      if (!operator) throw new ToolError(`"operator" is required: ${CELL_OPERATORS.join(', ')}.`);
      const formula1 = requireString(args, 'value1');
      const formula2 = operator === 'Between' || operator === 'NotBetween' ? requireString(args, 'value2') : undefined;
      conditional = formats.add('CellValue');
      conditional.cellValue.rule = { formula1, formula2, operator };
      applyHighlight(conditional.cellValue.format, args);
      break;
    }
    case 'text_contains':
      conditional = formats.add('ContainsText');
      conditional.textComparison.rule = { operator: 'Contains', text: requireString(args, 'text') };
      applyHighlight(conditional.textComparison.format, args);
      break;
    case 'top_bottom':
      conditional = formats.add('TopBottom');
      conditional.topBottom.rule = {
        rank: optionalNumber(args, 'rank') ?? 10,
        type: optionalEnum(args, 'top_bottom_type', TOP_BOTTOM_TYPES) ?? 'TopItems',
      };
      applyHighlight(conditional.topBottom.format, args);
      break;
    case 'formula': {
      const formula = requireString(args, 'formula');
      conditional = formats.add('Custom');
      conditional.custom.rule.formula = formula.startsWith('=') ? formula : `=${formula}`;
      applyHighlight(conditional.custom.format, args);
      break;
    }
  }

  conditional.load('id');
  await context.sync();
  const address = localAddress(range.address);
  recordUndo({ kind: 'conditionalFormat', sheetName, address, id: conditional.id });
  return { sheet: sheetName, address, conditionalFormat: type };
};
