# Paper2Excel

Paper2Excel is a Windows desktop app for turning paper lists in Excel or CSV files into structured AI-generated spreadsheet outputs. It is designed for Zotero exports, literature screening tables, PDF attachment paths, and user-defined review fields.

![Paper2Excel icon](assets/Paper2Excel.png)

## Features

- Read `.xlsx` and `.csv` files.
- Keep selected source columns in the output workbook.
- Use all available row data as AI input by default, regardless of which source columns are retained in the output.
- Use `File Attachments` as the default PDF path column for Zotero-style exports.
- Extract PDF text and include it as optional context for the model.
- Define custom output fields with `string`, `number`, or `boolean` types.
- Move output fields up or down; the field order controls the exported column order.
- Save and load reusable field templates.
- Ask the model to turn plain-language field descriptions into stable English prompts.
- Call OpenAI-compatible `/chat/completions` APIs, including OpenAI, DeepSeek, Kimi, Qwen, GLM, Gemini, OpenRouter, Ollama, LM Studio, and similar providers.
- Save progress during long runs and record diagnostics for failed rows.
- Package as a portable Windows release that does not require Python on the target computer.

## Before And After Example

The Chinese README includes a visual workflow example showing how a Zotero CSV row becomes a structured Paper2Excel output row:

![Paper2Excel workflow example](docs/images/example-before-after.png)

In short, Paper2Excel keeps selected source columns, adds user-defined AI output columns, optimizes plain-language descriptions into stable English prompts, and writes the final spreadsheet as retained source columns plus AI fields plus diagnostic columns.

The public example row is available in [docs/example-row-after.json](docs/example-row-after.json).

## Download And Run

Download the Windows release zip from GitHub Releases:

```text
Paper2Excel-v0.1.1-windows.zip
```

Extract the zip and run:

```text
Paper2Excel\Paper2Excel.exe
```

Copy the whole `Paper2Excel` folder. Do not copy only the single EXE, because the app needs the bundled `_internal` runtime directory.

## Basic Workflow

1. Open `Paper2Excel.exe`.
2. Choose a model provider and confirm the Base URL.
3. Enter the model name and API key.
4. Fill a proxy URL only when your network requires one, for example `http://127.0.0.1:7897`.
5. Click `测试连接` to verify the model connection.
6. Choose an Excel or CSV input file.
7. Confirm the output file path.
8. Choose which original columns should be retained in the output.
9. Configure custom output fields.
10. Use `上移` and `下移` to reorder output fields.
11. Optionally save the field setup as a template for reuse.
12. Preview the first 3 rows before running the full batch.

## Model Base URLs

The Base URL can be either a base path:

```text
https://api.openai.com/v1
```

or a full endpoint:

```text
https://api.openai.com/v1/chat/completions
```

The app detects and builds the final `/chat/completions` URL automatically.

Common provider examples:

| Provider | Base URL |
|---|---|
| OpenAI | `https://api.openai.com/v1` |
| DeepSeek | `https://api.deepseek.com` |
| Kimi / Moonshot Intl | `https://api.moonshot.ai/v1` |
| Kimi / Moonshot CN | `https://api.moonshot.cn/v1` |
| Qwen / DashScope CN | `https://dashscope.aliyuncs.com/compatible-mode/v1` |
| Qwen / DashScope Intl | `https://dashscope-intl.aliyuncs.com/compatible-mode/v1` |
| GLM / BigModel | `https://open.bigmodel.cn/api/paas/v4` |
| Gemini OpenAI-compatible | `https://generativelanguage.googleapis.com/v1beta/openai` |
| OpenRouter | `https://openrouter.ai/api/v1` |
| SiliconFlow | `https://api.siliconflow.cn/v1` |
| Groq | `https://api.groq.com/openai/v1` |
| Mistral | `https://api.mistral.ai/v1` |
| xAI / Grok | `https://api.x.ai/v1` |
| Together AI | `https://api.together.xyz/v1` |
| Ollama Local | `http://localhost:11434/v1` |
| LM Studio Local | `http://localhost:1234/v1` |
| vLLM Local | `http://localhost:8000/v1` |

## Field Templates

Field templates save only reusable task configuration:

- `task_description`
- `output_fields`
- field names, field types, required flags, plain descriptions, and optimized prompts

Templates do not save API keys, proxy settings, input paths, output paths, or local user configuration.

## API Key Safety

By default, Paper2Excel does not save API keys. An API key is written to the local `user_config.json` file only if the user explicitly enables `保存 API Key 到本机配置文件`.

Before publishing source code or a release package, make sure these files are not included:

- `user_config.json`
- `.env`
- local Excel, CSV, or PDF files
- output workbooks
- logs
- build caches

The repository `.gitignore` excludes local configuration, build outputs, release folders, logs, outputs, and Python cache files.

## Run From Source

Create the Conda environment:

```powershell
conda env create -f environment.yml
conda activate paper2excel
python main.py
```

Or run with a specific Python executable:

```powershell
& "C:\Path\To\paper2excel\python.exe" ".\main.py"
```

## Build The Windows Release

Run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Specify Python and version if needed:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1 -Python "C:\Path\To\paper2excel\python.exe" -Version "v0.1.1"
```

The build script runs tests, builds the PyInstaller app, performs an EXE self-test, scans the release directory for obvious secret patterns, and creates:

```text
release\Paper2Excel-v0.1.1-windows.zip
```

## License

Paper2Excel is released under the MIT License. See [LICENSE](LICENSE).
