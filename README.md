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
  - `get_workbook_context` – sheets, active sheet, selection, used ranges, headers and tables
  - `read_range` – read values/formulas (limited per call; 2000 cells by default, configurable in Settings)
  - `activate_worksheet` / `create_worksheet`
  - `write_table` – write data as a native Excel table with auto-fitted columns
  - `set_range_values_or_formulas` – values or live formulas (SUM, XLOOKUP, FILTER…)
  - `format_range` – bold, colors, alignment, number formats
- ✅ **Approval & undo**: Changes to the workbook are shown for approval before they're applied (can be turned off in Settings), and all changes from one request can be undone together.
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
- `src/agent/excelTools.ts` – Office.js executors and the undo journal
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
