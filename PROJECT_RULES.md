# Project Rules — Stock Screener (Google Sheets + Apps Script)

## General Principles
- Keep the code **simple, explicit, and modular** — one main function per file.
- All ChatGPT rewrites must be **full-file replacements**, never snippets.
- Use **strict mode** and clear variable names (`camelCase`).
- Centralize configuration and constants in `Globals.gs`.
- Prefer **pure functions** that accept and return values; avoid hidden state.

## Sheet Interaction
- Always read and write by **header name**, not hard-coded column index.
- Batch reads/writes: never exceed ~2,000 cells per operation.
- Use `SpreadsheetApp.getActiveSpreadsheet()` only inside menu entrypoints; elsewhere, pass sheet references explicitly.

## Data Integrity
- **Upsert semantics:** key = (Sheet, Ticker, FiscalYear).
- No duplicate keys allowed; existing rows updated in place.
- Keep “tall” source tables normalized; generate “wide” aggregates separately.

## Metrics Rules
### Per-Share Logic
- Use **Weighted Average Shares Diluted**; skip the row if missing or zero.
- Derived fields:
  - `RevenuePerShare` = Revenue / SharesDiluted  
  - `OperatingIncomePerShare` = OperatingIncome / SharesDiluted  
  - `EquityPerShare` = TotalStockholdersEquity / SharesDiluted  
  - `DebtToEquity` = TotalDebt / TotalStockholdersEquity

### CAGR Calculation
- If either start ≤ 0 or end ≤ 0 → replace that endpoint with 0.01.
- Set `AdjustedFlag = TRUE` if any endpoint was adjusted.
- Use standard formula:  
CAGR = (End / Start)^(1/Years) - 1
- Round CAGRs to 4 decimals.

## API Rules
- Source: **EODHD** Fundamentals + Prices endpoints.
- Cache raw API responses in `Raw_Fundamentals_*` sheets.
- Respect API call limits; throttle via `Utilities.sleep()` if needed.

## Logging & UX
- Wrap long operations with `withUiSpinner(label, fn)` to show start/finish time.
- Log counts (rows read/written) and duration in ms.
- Never leave console logs in production; use `Logger.log()` sparingly.

## Versioning & Safety
- Tag working versions in **Apps Script → File > Manage Versions…**
- Each commit to GitHub must include a line in `CHANGELOG.md`.
- Do not push untested code to `main`.

## Security
- Store API keys in **Script Properties**, never in code.
- Retrieve via:
```js
const EODHD_KEY = PropertiesService.getScriptProperties().getProperty('EODHD_KEY');
