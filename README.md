# Stock Screener — Google Sheets + Apps Script
## Overview A modular Google Sheets-based stock screener powered by the **EODHD API**. It
extracts fundamentals, prices, and per-share metrics, then computes derived CAGRs and valuation
ratios directly inside Sheets.
### Key Features - Automated EODHD data extraction (annual + quarterly) - Per-share calculations
(RevPS, GPPS, OpPS, EPS, FCFPS, EquityPS) - CAGR engine with ≤0 → 0.01 adjustment rule -
Upsert logic to prevent duplicate ticker-year rows - Menu-driven Sheets interface (via onOpen
menu) - Modular scripts (Extractor, Per_Share, Calculated_Metrics, Control_Center) - CursorAI +
ChatGPT compatible workflow
### Folder Structure src/ ■■ code.ts — menu + entrypoints ■■ globals.ts — constants, config, API
key loader ■■ extractor/ — fundamentals, prices, splits, dividends ■■ per_share/ — per-share
logic + builder ■■ calc_metrics/ — CAGR + flags + upsert to Calculated_Metrics ■■
control_center/ — data validation + dropdowns ■■ utils/ — shared helpers (upsert, sheets, logging)
### Sheet Contracts See SHEET_CONTRACTS.md for column specs and keys.
### Project Rules See PROJECT_RULES.md for coding style, upsert policy, and per-share logic.
### CursorAI & ChatGPT Includes CURSOR_RULES.md and .cursorrules for AI pair-programmer
config.
### Version Control - Managed via Git + GitHub - Pushed to Google Apps Script via clasp - Each
update logged in CHANGELOG.md
### Secrets EODHD API key stored in Script Properties: const EODHD_KEY =
PropertiesService.getScriptProperties().getProperty('EODHD_KEY');
### Workflow Summary 1. Edit files locally in src/ 2. clasp push → test in Sheet 3. git add . && git
commit -m "feat(...): description" 4. git push → sync with GitHub 5. (Optional) tag version in Apps
Script
### Example Menu Stock Screener → Extract Fundamentals → Annual / Quarterly Per Share →
Build Per Share, Compute CAGRs Control Center → Refresh Dropdowns Settings → Toggle Mock
Mode, Set API Key
Author: Denver So License: MIT © 2025 Denver So