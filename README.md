# Excel AI

<div align="center">
  <img src="public/excel_ai_logo.png" alt="Excel AI Logo" width="128" />
</div>

**Excel AI** is a serverless AI agent that integrates directly into Microsoft Excel. It can inspect your whole workbook, create sheets, write native Excel tables and live formulas, and format data, using your favorite AI models without ever leaving your spreadsheet.

<div align="center">
  <img src="public/excel_ai_screenshot.png" alt="Excel AI Screenshot" width="800" />
</div>

## Features

- 🤖 **Workbook-level agent (tool calling)**: The model works through a set of Excel tools and gets the result of every call back, so it can fix its own mistakes (duplicate sheet names, invalid ranges, `#NAME?` formula errors…):
  - **Inspect & analyze:** `get_workbook_context` (sheets, selection, tables, charts, PivotTables, named ranges), `read_range` (2000 cells per call by default, configurable), `search_workbook`, `profile_data` (per-column statistics computed in the browser, for large data)
  - **Formulas:** `trace_formula` (precedents/dependents, ExcelApi 1.12+), `find_formula_errors`
  - **Sheets & names:** `activate_worksheet`, `create_worksheet`, `rename_worksheet`, `delete_worksheet`, `set_worksheet_visibility`, `create_named_range`
  - **Data:** `write_table`, `convert_range_to_table`, `import_attachment`, `set_range_values_or_formulas` (live formulas: SUM, XLOOKUP, FILTER…), `clear_range`, `copy_range` (copy/move, paste values/formats, transpose), `fill_range`, `find_replace`, `remove_duplicates`, `insert_range` / `delete_range` (rows, columns, cells), `sort_range`, `filter_table`, `add_data_validation`
  - **Formatting & layout:** `format_range`, `add_conditional_format`, `set_rows_columns` (widths, heights, hide/unhide), `freeze_panes`, `merge_cells`, `add_comment`
  - **Charts & PivotTables:** `create_chart`, `update_chart`, `delete_chart`, `create_pivot_table`, `update_pivot_table`, `delete_pivot_table`
  - Tools that need a newer Excel than the one running are hidden from the model automatically.
- 💬 **Chat**: answers stream in as they are generated and are rendered as Markdown; cell references like `Sales!B5` are clickable and select the range. Each request ends with a list of the changes made, linked to the affected ranges.
- 📎 **Attachments**: CSV, TSV, JSON and text files are parsed in the browser (large files are imported directly into the workbook without passing through the model); images and PDFs are sent to models that accept them. Images can also be pasted.
- ✅ **Approval & undo**: Changes are shown for approval with a before → after preview of the cells (can be turned off in Settings; irreversible actions such as deleting a chart always ask). All changes from one request can be undone together, including deleted sheets (kept as a hidden backup while they can be undone).
- 💾 **Per-workbook history**: the conversation and the undo history are saved in the browser per workbook and restored when the add-in is reopened.
- 🧩 **Robust agent**: long conversations are summarized automatically, failed requests (rate limits, server errors) are retried, models without tool support are flagged, prompt caching is used for Anthropic models on OpenRouter, and custom instructions (with templates for finance, accounting, data cleaning…) are added to every request.
- 🧠 **Bring Your Own AI**: Any OpenAI-compatible API with tool calling:
  - **OpenRouter** (default, model `deepseek/deepseek-chat`; also e.g. `anthropic/claude-3.5-sonnet`, `meta-llama/llama-3.3-70b-instruct`)
  - **Google Gemini**, **OpenCode Zen**, **LM Studio** (local models), or any **custom** endpoint (Groq, Together AI…)
- 🔒 **Serverless**: There is no backend. The add-in talks directly from your browser to the AI provider. Settings and API keys are kept in the browser's local storage, which is not encrypted and is shared by every page on the same origin — host the add-in on an origin you control.
- 🌐 **Optional CORS proxy**: Requests can be routed through `corsproxy.io` for providers that block browser requests. Your API key and data then pass through that third party, so it's off by default.

## Installation

You can sideload this add-in directly into Excel for Desktop or Excel on the Web:

1. Download the `manifest-prod.xml` from this repository.
2. **Excel on the Web:** 
   - Open Excel in your browser.
   - Go to `Insert` > `Add-ins`.
   - Click `Upload My Add-in` and select the `manifest-prod.xml` file.
3. **Excel Desktop (Windows/Mac):**
   - Place the `manifest-prod.xml` in a shared network folder.
   - Go to Excel Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.
   - Add the folder path and check "Show in Menu".
   - Go to `Insert` > `My Add-ins` > `Shared Folder` and select Excel AI.

## Supported Providers

You can configure your provider in the **Settings** menu of the add-in.

The selected model must support tool (function) calling.

- **OpenRouter:** Requires an OpenRouter API key. Base URL is configurable (default `https://openrouter.ai/api/v1`). Use **Load models** to pick from the live model list.
- **Google Gemini:** Requires a Gemini API key (uses Gemini's OpenAI-compatible endpoint).
- **OpenCode Zen:** API key optional.
- **LM Studio:** Requires LM Studio running locally on port `1234` with CORS enabled and a model that supports tool use.
- **Custom (OpenAI Compatible):** Any API that implements OpenAI chat completions with tools.
  - *Groq Base URL:* `https://api.groq.com/openai/v1`
  - *Together AI Base URL:* `https://api.together.xyz/v1`

Requires Excel with ExcelApi 1.9 or later (Microsoft 365, Excel 2021+, Excel on the web).

## Development

```bash
npm install
npm run dev     # https://localhost:5173 — sideload manifest.xml
npm run build
npm run lint
```

Code layout:

- `src/agent/tools.ts` – JSON schemas of the tools sent to the model
- `src/agent/excel/` – Office.js executors (`workbookTools`, `dataTools`, `editTools`, `structureTools`, `formatTools`, `formulaTools`, `analysisTools`, `insightTools`), approval previews (`preview.ts`) and the undo journal (`undo.ts`)
- `src/agent/context.ts` – conversation size control and summarization
- `src/agent/attachments.ts` – file parsing for attachments
- `src/agent/llmClient.ts` – OpenAI-compatible chat completions client
- `src/agent/agentLoop.ts` – `runAgentLoop`: inference → tool dispatch → feedback loop
- `src/components/` – chat and settings task pane UI

## Publishing to Microsoft AppSource

To publish this add-in to the Microsoft Store:
1. Ensure the code is hosted securely on HTTPS (e.g., GitHub Pages).
2. Create a developer account at the [Microsoft Partner Center](https://partner.microsoft.com/).
3. Create a new "Office Add-in" offer and upload `manifest-prod.xml`.
4. Provide the required screenshots, privacy policy, and terms of use links.
5. Submit for certification.

## Legal

- [Privacy Policy](https://juancarlosop03.github.io/excel-ai-addin/privacy.html)
- [Terms of Use](https://juancarlosop03.github.io/excel-ai-addin/terms.html)

---
*Built with React, Fluent UI, and Vite.*
