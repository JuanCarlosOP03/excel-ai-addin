import type { AppSettings } from '../utils/storage';

/**
 * Skills are expert playbooks for common Excel deliverables. The system prompt lists their
 * names and descriptions; the model loads the full instructions with the use_skill tool only
 * when it needs them (keeping the prompt small), or the user attaches them to a message.
 */
export interface Skill {
  id: string;
  name: string;
  description: string;
  instructions: string;
  builtIn: boolean;
}

const BUILT_IN: Omit<Skill, 'builtIn'>[] = [
  {
    id: 'financial-model',
    name: 'Financial model',
    description: 'Projections, 3-statement models, DCF valuations, scenarios and sensitivity tables.',
    instructions: `# Financial model

Structure
- Use separate sheets in this order: "Inputs" (all assumptions), "Calc"/"Model" (projections), "Outputs" (summary, valuation, charts). Create missing sheets with create_worksheet.
- Time runs left to right: one column per period, periods in row 1 or 2 with a clear header ("FY2025", "FY2026"…). Line items run top to bottom with labels in column A, units in column B.
- Never hard-code a number inside a formula: every assumption lives in Inputs and is referenced. Use create_named_range for key drivers (Growth, WACC, TaxRate) so formulas read clearly.

Conventions
- Font colors: blue #0000FF for hard-coded inputs, black for formulas, green #008000 for links to other sheets. Apply them with format_range after writing.
- Number formats: amounts #,##0;(#,##0);"-" ; percentages 0.0% ; multiples 0.0"x" ; years as plain text or 0.
- Totals in bold with a top border. Freeze panes below the headers and right of the labels (freeze_panes).
- Signs: revenues and assets positive, costs shown as positive numbers and subtracted in formulas (be consistent and say which convention you use).

Building blocks
- Growth projections: =prior*(1+growth). Margins: =line/revenue.
- 3-statement: Income statement → Balance sheet → Cash flow. The balance sheet must balance: add a check row "Assets - Liabilities - Equity" that must be 0, and highlight non-zero values with add_conditional_format (cell_value NotEqualTo 0, red).
- DCF: Free cash flow = EBIT*(1-tax) + D&A - CapEx - ΔNWC. Discount factor =1/(1+WACC)^period (or mid-year). Terminal value = FCF_last*(1+g)/(WACC-g) (check WACC > g). Enterprise value = SUM(PV of FCF) + PV(TV). Equity value = EV - net debt; per share = equity / shares.
- Scenarios: a scenario selector cell with add_data_validation (list: Base, Upside, Downside) and CHOOSE/INDEX-MATCH to pick the active assumptions.
- Sensitivity: a 2-way grid (e.g. WACC vs g) where each cell recomputes the result. Excel Data Tables can't be created through the API, so write explicit formulas in each grid cell that rebuild the result with the row/column values, and add a color scale.

Before finishing
- Run find_formula_errors on the model sheets and fix every error.
- Summarize the structure, key assumptions (with sheet-qualified references) and results.`,
  },
  {
    id: 'dashboard',
    name: 'Dashboard',
    description: 'Executive dashboards and reports: KPI cards, summaries by category, PivotTables and charts on one sheet.',
    instructions: `# Dashboard

Plan first
- Inspect the data with get_workbook_context and profile_data: identify the date column, categories (region, product, channel…) and measures (sales, units, cost).
- Pick 3–6 KPIs (totals, growth vs prior period, averages, margins) and 2–4 views (trend over time, ranking by category, mix/share).

Data layer
- If the data isn't an Excel table yet, convert it (convert_range_to_table) so formulas and PivotTables grow with it.
- Build summaries either with create_pivot_table (fast, interactive) on a helper sheet, or with SUMIFS/COUNTIFS/AVERAGEIFS tables (fully live, better for fixed layouts). Prefer formulas referencing the table's structured references.

Layout (sheet "Dashboard")
- Row 1: title (bold, 16–18 pt) and a subtitle with the data period. Hide gridlines is not available, so use a white fill on the area for a clean look.
- Rows 3–6: KPI cards — a label cell above a big value cell (bold, 14–16 pt, number format), light fill (#F2F2F2), outline borders; leave a blank column between cards.
- Below: charts aligned to a grid. Use anchor_cell and width/height in create_chart so charts line up (e.g. 480×288 pt, two per row).
- Charts: line or column for trends over time, bar (sorted descending) for rankings, pie/doughnut only for ≤ 5 parts of a whole. Always set a title; remove the legend for single-series charts.
- Keep one consistent palette (e.g. #1F4E79, #2E75B6, #9DC3E6, #F4B183) and the same number formats everywhere.

Finish
- freeze_panes is usually not needed on dashboards; activate the dashboard sheet at the end.
- Explain where each KPI comes from (sheet-qualified references) and how to refresh PivotTables (update_pivot_table refresh) when data changes.`,
  },
  {
    id: 'data-cleaning',
    name: 'Data cleaning',
    description: 'Find and fix messy data: duplicates, blanks, inconsistent text, numbers or dates stored as text.',
    instructions: `# Data cleaning

1. Profile before touching anything: profile_data on the data. Report per column: type (mixed types are a red flag), blanks, distinct counts, suspicious top values (spelling variants like "Mexico"/"México"/"MEXICO"), min/max outliers.
2. Ask before destructive changes unless the user asked for them explicitly. Default to working on a copy: copy_range the data to a new sheet "Clean" (paste values) and clean there, leaving the source intact.
3. Fixes, using the right tool:
   - Exact duplicate rows: remove_duplicates (say which columns define a duplicate).
   - Extra spaces / casing: a helper column with =TRIM(), =PROPER()/=UPPER(), then copy_range paste values over the original column and clear the helper.
   - Spelling variants: find_replace per variant (whole_cell true) after confirming the mapping with the user.
   - Numbers stored as text: =VALUE(TRIM(x)) or =--x in a helper column; dates as text: =DATEVALUE() or DATE(RIGHT(),MID(),LEFT()) depending on the pattern you observed.
   - Blanks: leave them, fill down with fill_range when the blank means "same as above", or flag them — never invent values.
   - Split/merge columns with TEXTBEFORE/TEXTAFTER/TEXTSPLIT or concatenation with &.
4. Convert the clean result into a table (convert_range_to_table) and apply number/date formats.
5. Add a short "Cleaning log" (sheet or message): every rule applied and how many rows/cells changed.`,
  },
  {
    id: 'professional-formatting',
    name: 'Professional formatting',
    description: 'Make tables, reports and sheets look polished and consistent.',
    instructions: `# Professional formatting

- Data blocks become Excel tables (convert_range_to_table / write_table) with a subtle style (TableStyleMedium2 or TableStyleLight9). Don't mix tables with manual fills on the same range.
- Headers: bold, wrap_text if long, centered over numeric columns. Freeze the header row (freeze_panes rows: 1, or the header row number).
- Alignment: text left, numbers and dates right, short codes/flags centered.
- Number formats by meaning: currency $#,##0.00 (or the user's currency), large amounts #,##0, percentages 0.0%, dates yyyy-mm-dd or dd/mm/yyyy (follow the user's locale if known), negatives in parentheses for finance #,##0;(#,##0).
- Column widths: autofit_columns, then cap very wide text columns with set_rows_columns column_width (~40) plus wrap_text.
- Totals row: bold with a top border (format_range borders "outline" on the total row or font_bold + top border), clearly labeled "Total".
- Use color sparingly: one accent color for headers or highlights, light grey (#F2F2F2) for input or subtotal areas, red/green only for meaning (conditional formats for negatives, targets).
- Titles: a title in A1 (bold, 14 pt) and an optional subtitle/date in A2 above the data; merge_cells only for titles, never inside data.
- Consistency across sheets: same fonts sizes, same palette, same formats for the same kind of number.
- Finish with find_formula_errors if the sheet has formulas.`,
  },
  {
    id: 'advanced-formulas',
    name: 'Advanced formulas',
    description: 'Robust modern formulas: XLOOKUP, dynamic arrays (FILTER, UNIQUE, SORT), LET, LAMBDA, conditional aggregation and error handling.',
    instructions: `# Advanced formulas

Choosing functions
- Lookups: XLOOKUP(key, lookup_col, return_col, "Not found") instead of VLOOKUP; INDEX/MATCH for older Excel. Two-key lookups: XLOOKUP(1, (colA=a)*(colB=b), return_col).
- Conditional totals: SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS; arrays with SUMPRODUCT for complex conditions.
- Dynamic arrays (one formula spills a whole result): UNIQUE, FILTER(data, condition, "No results"), SORT/SORTBY, SEQUENCE, TAKE/DROP, VSTACK/HSTACK, TEXTSPLIT. Write the formula in ONE cell and leave room below/right for the spill; a #SPILL! error means something blocks the area.
- Readability: LET(name, value, …, result) for repeated sub-expressions; LAMBDA + create_named_range only if the user wants reusable custom functions.
- Dates: EOMONTH, EDATE, NETWORKDAYS, YEAR/MONTH, TEXT(date,"yyyy-mm") for grouping keys.
- Text: TEXTJOIN, CONCAT, TRIM, LEFT/MID/RIGHT, TEXTBEFORE/TEXTAFTER.

Robustness
- Reference whole table columns with structured references (Table[Column]) or bounded ranges; avoid full-column references inside heavy SUMPRODUCT.
- Use absolute references ($A$1) for fixed cells before filling (fill_range) and relative ones for the row.
- Handle expected errors explicitly (IFERROR only around the part that can fail, or XLOOKUP's if_not_found); never hide errors that indicate real problems.
- Avoid volatile functions (OFFSET, INDIRECT, TODAY/NOW in large ranges) unless needed.
- Always write formulas in English with comma separators.

Process
- Write one formula, check its result in the tool output, then fill it to the rest of the range (fill_range) instead of writing hundreds of formulas in the arguments.
- After writing, check formulaErrors in the result and fix them; use trace_formula to explain or debug a result.`,
  },
  {
    id: 'data-analysis',
    name: 'Data analysis',
    description: 'Explore a data set and answer questions with numbers: trends, rankings, comparisons, outliers and insights.',
    instructions: `# Data analysis

1. Understand the data: get_workbook_context, then profile_data on the relevant table. State the grain ("one row per order"), the period covered and the key columns.
2. Clarify the question if needed (which metric, which period, which segments).
3. Compute with live formulas or PivotTables in a new sheet "Analysis" — never by mental math. Typical analyses:
   - Totals and averages by segment (SUMIFS / PivotTable).
   - Trends: group by month (TEXT(date,"yyyy-mm") key or PivotTable on dates), period-over-period growth =(this-prior)/prior.
   - Rankings: SORTBY / LARGE / PivotTable sorted descending; top N and their share of total.
   - Distribution and outliers: MIN/MAX/MEDIAN/PERCENTILE, values beyond 1.5×IQR.
   - Contribution / mix: share of total = part/SUM(total).
4. Visualize the 1–3 most important results with create_chart (titles that state the insight, e.g. "North drives 45% of sales").
5. Answer in plain language: lead with the direct answer and the key numbers, then 3–5 insights, each with a sheet-qualified reference to where it's computed. Mention caveats (missing data, small samples). Distinguish facts from hypotheses.`,
  },
  {
    id: 'charts',
    name: 'Charts',
    description: 'Pick the right chart type and make clear, well-labeled charts.',
    instructions: `# Charts

Choosing the type
- Trend over time: Line (many periods) or ColumnClustered (few periods). Time always on the horizontal axis, in chronological order.
- Ranking / comparison of categories: BarClustered, sorted descending (sort the source first with sort_range), long labels read better on bars.
- Part of a whole: Pie or Doughnut only for 2–5 parts; otherwise a sorted bar or ColumnStacked100.
- Composition over time: ColumnStacked or AreaStacked.
- Relationship between two measures: XYScatter (add trend reasoning in the message).
- Distribution: Histogram or Boxwhisker. Bridges (start → changes → end): Waterfall. Hierarchies: Treemap or Sunburst.

Data preparation
- The source range must be a compact block: header row + category column + one column per series. Build it with formulas or a PivotTable if it doesn't exist, and keep it near the chart or on a helper sheet.
- Aggregate first: never chart thousands of raw rows.

Design
- Title states what the chart shows (or the insight). Axis titles with units when not obvious. Legend only with 2+ series (legend_position None otherwise, Bottom when needed).
- Size ~480×288 pt; align multiple charts to the same top/left cells.
- Place charts next to (not on top of) data: anchor_cell to the right of the data or on a dashboard sheet.
- To change a chart later use update_chart (type, data, titles, legend, position) instead of creating a new one.`,
  },
  {
    id: 'budget-forecast',
    name: 'Budget & forecast',
    description: 'Budgets, forecasts and actual-vs-budget variance analysis.',
    instructions: `# Budget & forecast

Layout
- Rows: accounts or cost centers (grouped with subtotals); columns: months (Jan…Dec) plus Total, or Budget | Actual | Variance | Variance % per period.
- Keep inputs (budget amounts, growth rates, seasonality %) separated from formulas and colored as inputs (blue font or light yellow fill).

Formulas
- Annual totals =SUM(Jan:Dec). Monthly spread of an annual budget: =Annual*Seasonality% (seasonality row must sum to 100%; add a check).
- Forecast methods: run-rate (=AVERAGE(last 3 actuals)), growth (=prior*(1+g)), or trend (=FORECAST.LINEAR / TREND on the history). Say which method you used.
- Year-to-date: =SUMIFS(actuals, month_col, "<="&current_month) or SUM over the elapsed months.
- Variance = Actual - Budget; Variance % = IF(Budget=0, "", Variance/Budget). For costs, an overspend is unfavorable: state the sign convention.

Highlighting
- add_conditional_format on Variance %: red when unfavorable beyond a threshold (e.g. < -5% for revenue, > 5% for costs), green when favorable; or icon_set.
- Totals and subtotals bold with a top border; number format #,##0;(#,##0).

Finish with a short commentary: the biggest variances (with sheet-qualified references), their likely drivers if visible in the data, and the forecast for the full year.`,
  },
  {
    id: 'trackers',
    name: 'Trackers & forms',
    description: 'Input sheets, trackers and simple databases with dropdowns, validation and status highlighting.',
    instructions: `# Trackers & forms

- Build the tracker as an Excel table (write_table) with clear columns, e.g. ID, Date, Owner, Category, Status, Due date, Amount, Notes. Include 1–3 example rows only if the user wants them.
- Put dropdown options in a "Lists" sheet (one column per list, with a header) and reference them with add_data_validation list_source (e.g. "=Lists!$A$2:$A$10"); for short fixed lists, list_values is fine. Hide the Lists sheet afterwards if the user prefers (set_worksheet_visibility).
- Validate inputs: dates (type date, GreaterThan a reasonable start), amounts (decimal ≥ 0), text length limits for codes; add input_message hints and error_message texts.
- Status highlighting with add_conditional_format: formula rules such as =$E2="Done" (green), =AND($F2<TODAY(),$E2<>"Done") (red, overdue), =$F2-TODAY()<=7 (amber, due soon). Apply rules to the whole table body.
- Helpful computed columns: Days open =IF([@Status]="Done","",TODAY()-[@Date]); summary counts per status with COUNTIFS on a small summary block above or on another sheet.
- Freeze the header row, autofit columns, and keep Notes wide with wrap_text.`,
  },
];

export const BUILT_IN_SKILLS: Skill[] = BUILT_IN.map(s => ({ ...s, builtIn: true }));

/** Built-in skills plus the user's own (a custom skill with the same id replaces the built-in one). */
export const getSkills = (settings: AppSettings): Skill[] => {
  const custom = settings.customSkills.map(s => ({ ...s, builtIn: false }));
  const overridden = new Set(custom.map(s => s.id));
  return [...BUILT_IN_SKILLS.filter(s => !overridden.has(s.id)), ...custom];
};

export const findSkill = (settings: AppSettings, idOrName: string): Skill | undefined => {
  const key = idOrName.trim().toLowerCase().replace(/^\//, '');
  return getSkills(settings).find(s => s.id.toLowerCase() === key || s.name.toLowerCase() === key);
};

/** "Sales report" → "sales-report". */
export const skillIdFromName = (name: string) =>
  name.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'skill';

/** Text block injected into the conversation when a skill is loaded or attached. */
export const formatSkill = (skill: Skill) => `<skill id="${skill.id}" name="${skill.name}">\n${skill.instructions}\n</skill>`;
