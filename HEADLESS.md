# Headless image analyzer

Agents: when asked to use this application, read AGENTS.md in this folder first. Install dependencies with `python -m pip install -r requirements.txt`.

```powershell
python analyze_cli.py "D:\photos" --excel "D:\reports\analysis.xlsx" --local-model gemma
```

Pass files or folders as inputs; folders are searched recursively. Local options are `--local-model qwen` (Qwen 3.6 35B A3B, default, port 8080) and `--local-model gemma` (Gemma 4 E4B, port 8081). The app reuses a responding server or calls the corresponding existing llama.cpp startup script if needed. It does not stop other servers; simultaneous models can compete for GPU memory. Pi may stop its managed servers when its model changes or its session ends. `--base-url` connects to a different running OpenAI-compatible vision server and does not start it.

For Gemini 3.1 Flash Lite, set `OPENROUTER_KEY` or `OPENROUTER_API_KEY` and use `--provider openrouter --model google/gemini-3.1-flash-lite`. This requires network access and OpenRouter billing. PDFs require OpenRouter. Translation is off by default; add `--translate` to enable it. `--geocode` looks up available GPS coordinates using Google Maps and requires `GEOCODING_KEY` and network access. Geocoding transmits coordinates to Google. GPS coordinates can be extracted even without geocoding. Run `python analyze_cli.py --help` for options.

An existing workbook uses its active sheet unless `--sheet` is specified. Row 1 must have a `Filename` header (case insensitive), or be empty. Recognized columns are matched by header; missing analyzer columns are added. Other sheets and columns are retained. A new .xlsx workbook is created if needed. Close the workbook in Excel before running. Back up complex workbooks: openpyxl may not preserve every Excel feature. Existing .xlsm files retain VBA where supported.

The command prints a JSON result with processed paths and per-file errors. Analysis data is written to the workbook, not included in that JSON; read the workbook to inspect it. Exit code 0 means all files succeeded; 1 means at least one failed or setup/save failed. Successful results are saved even when some inputs fail; retrying the entire batch will append duplicates. Credentials are read from environment variables and are not printed.
