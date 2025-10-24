# CursorAI Coding Rules — Stock Screener (Google Sheets + Apps Script)
## 1. General Behavior - Treat this repo as a **TypeScript / Apps Script hybrid** project. - Always
rewrite **complete files**, never partial diffs or inline patches. - Follow the structure enforced by
`PROJECT_RULES.md` and `SHEET_CONTRACTS.md`. - Assume execution environment:
**Google Apps Script (V8 runtime)**. - Target ES2018 or later. - Always preserve top-of-file
comments, headers, and version notes.
## 2. File Organization - Keep each module focused on one function: - `code.ts` → menu wiring /
entrypoints only - `extractor/` → EODHD fundamentals, prices, splits, dividends - `per_share/` →
per-share metrics logic - `calc_metrics/` → CAGR + summary metrics - `control_center/` → data
validation UI + dropdowns - `utils/` → I/O, upsert, logging - Never move files silently; explicit rename
+ comment required.
## 3. Coding Style - Use `const` and `let` (no `var`). - Always specify types when using `.ts`. - Use
camelCase for variables, PascalCase for exported functions. - Indentation: 2 spaces. - Place
imports at top; no unused imports. - No magic numbers — constants go to `Globals.ts`.
## 4. Google Sheets I/O - Read/write by **header name**, not numeric index. - When updating
rows, always match headers by name. - Batch writes with `sheet.getRange().setValues()`. - Use
`upsertRange()` helper for insert/update logic.
## 5. CAGR Logic - Apply ≤0 → 0.01 rule before calculation. - Track flags per metric (e.g.,
`OpPS_AdjustedFlag`, `EPS_AdjustedFlag`). - Formula: `CAGR = (End / Start)^(1/Years) - 1`. -
Round to 4 decimals.
## 6. Per-Share Rules - Skip if `SharesDiluted <= 0` or missing. - Use Weighted Avg Shares
Diluted. - Derived columns: RevPS, GPPS, OpPS, EPS, FCFPS, EquityPS, DebtToEquity.
## 7. Upsert Semantics - Unique key = (SheetName, Ticker, FiscalYear). - Check existing rows
before appending. - Preserve header order.
## 8. Logging & UI - Wrap long ops with `withUiSpinner(label, fn)`. - Log start, rows processed,
elapsed ms. - Use `Logger.log()` only when debugging.
## 9. API Handling - Get key from Script Properties. - Throttle requests (200–500 ms). - Catch and
rethrow parsing errors safely.
## 10. Project Safety - No hardcoded keys or sheet IDs. - No global mutable state. - Wrap network
code in try/catch.
## 11. Documentation - File headers: name, purpose, dependencies, called by. - Each exported fn:
include JSDoc docstring.
## 12. Collaboration - Cursor must rewrite **full files** only. - Must summarize plan before large
refactors. - Must follow PROJECT_RULES and SHEET_CONTRACTS.
## 13. Commit & Versioning - Use clear commit prefixes: feat:, fix:, refactor:, chore:. - Update
CHANGELOG.md each push.
## 14. Testing - Pure helpers testable with Vitest locally.
## 15. Output Formatting - Round numbers to 4 decimals. - Write blank for invalid cells.
**Summary:** CursorAI must treat this repo as production-grade code. Follow sheet contracts,
upsert logic, and ≤0→0.01 rule for CAGRs. Produce full-file rewrites only.