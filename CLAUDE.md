# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Desktop automation tool (Windows-only) that creates SAP purchase requisitions in bulk (transaction ME51N) and updates logistics records in the Cargo Heroes web platform. Data source is a Google Sheets spreadsheet. Distributed as a standalone `.exe` via PyInstaller.

## Running

```bash
# Install dependencies
pip install -r requirements.txt

# Run the GUI application
python APP_UNIFICADO_SAP_CH.py
```

The app requires a `credentials.json` (Google service account) placed in the same directory, plus a `config.ini` that is auto-generated on first run.

## Building the Executable

```bash
python build.py
```

Output lands in `dist/Automacao_SAP_CH.exe`. The build uses `--onefile --noconsole`; `icone.ico` is bundled automatically if present.

## Architecture

All logic lives in a single class `SAPAutomationGUI` in `APP_UNIFICADO_SAP_CH.py`. The class is both the UI controller and the automation engine — there is no separation layer.

**Two independent automation flows, each runs in a `daemon` thread:**

1. **SAP flow** (`start_automation` → `run_sap_automation`):
   - Connects to a running SAP GUI instance via `win32com.client` (COM scripting)
   - Reads pending rows from Google Sheets (rows where `Status` column is empty)
   - Groups rows into batches of up to 10 by `(ORIGEM, DESTINO)` pair
   - For each batch: navigates to ME51N, validates items, then creates the purchase requisition
   - Writes back the RC number and status to the sheet via `gspread` batch updates with `tenacity` retry

2. **Cargo Heroes flow** (`start_ch_automation` → `run_ch_automation`):
   - Opens Chrome via Selenium and handles SSO login (Google → SAML → Microsoft)
   - After login, **minimizes the browser** and switches to API mode
   - Injects JavaScript (`execute_async_script`) that calls the Cargo Heroes internal BFF API directly using the session token from `sessionStorage`
   - Processes two sheet tabs: the main tab and an optional `MAPEAMENTO` tab
   - Writes `OK` or `ERRO` back to a `CH OK` column

**Key design points:**
- `self.session` (SAP COM object) is protected by `threading.Lock` via a Python property
- `self.running` flag is the cooperative stop mechanism for both flows
- `sys.stdout` is monkey-patched to `LogRedirector` during automation; restored on close or stop
- `LogRedirector` parses `<<COLOR>>` sentinel tags embedded in print strings to apply text color in the UI
- Passwords are stored via `keyring` (Windows Credential Manager) when available; fall back to plain text in `config.ini`
- `get_data_path` / `get_resource_path` handle the `sys.frozen` (PyInstaller) vs dev path difference

**`DEPOSITO_MAPPING`** (dict on the class): maps airport/base codes (`BR0G`, `BR8A`, etc.) to warehouse codes (`AE01`, `AE13`). This is business logic — changes here affect what gets written to SAP's custom field (`ZZDEP_FORNEC`).

**Timezone correction in CH flow:** the JS injected into the browser adds 3 hours to boarding/landing timestamps (`fixTimezone`) to compensate for the Cargo Heroes server automatically discounting Brasília time (UTC-3).

## Configuration

`config.ini` sections: `[SAP]`, `[GOOGLE]`, `[CARGO_HEROES]`. All paths and credentials are set through the GUI's "Configurações" tab and persisted to this file (passwords via keyring when possible).

Logs rotate at 5 MB (3 backups) to `app_log.txt` in the app's data directory.

## Constraints

- **Windows-only**: depends on `win32com`, `pywin32`, `pywintypes`, and SAP GUI scripting.
- SAP GUI scripting must be enabled in the SAP Logon settings on the target machine.
- Chrome + ChromeDriver must be installed and on PATH for the Cargo Heroes flow.
- The app must run as the same Windows user that has the SAP session open (COM object is process-local).
