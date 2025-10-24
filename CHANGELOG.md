# Changelog — Stock Screener (Google Sheets + Apps Script)

### 2025-10-25 — Reverted TypeScript Migration

**Decision: Keep Original JavaScript Files**
- Attempted TypeScript migration with modular structure
- Google Apps Script doesn't support ES6 imports/exports natively
- Reverted to original working .js files
- All functionality preserved and working correctly
- TypeScript migration requires build tooling (webpack/rollup) - deferred to future phase

**Current State:**
- ✅ All original .js files working in Google Apps Script
- ✅ Menu functions correctly
- ✅ Ready for production use
- 📋 TypeScript refactoring documented in REFACTOR_SUMMARY.md for future reference

---

### 2025-10-24 — Phase 1: Code Cleanup & TypeScript Migration (REVERTED)

**Major Refactoring:**
- ✓ Migrated entire codebase from .js to TypeScript (.ts)
- ✓ Implemented modular folder structure:
  - `globals.ts` — centralized constants, API key management, configuration
  - `utils/` — format.ts, sheets.ts, spinner.ts, upsert.ts (reusable helpers)
  - `extractor/` — fundamentals.ts, splits_dividends.ts (EODHD data extraction)
  - `per_share/` — builder.ts (per-share metrics calculation)
  - `calc_metrics/` — cagr.ts (CAGR computation with ≤0→0.01 rule)
  - `control_center/` — data_validation.ts (dropdown management)
  - `code.ts` — menu registration and entrypoints

**New Features:**
- Type-safe number formatting utilities (format.ts)
- Comprehensive sheet I/O operations with header-name-based reading (sheets.ts)
- UI spinner with timing and error handling (spinner.ts)
- Generic upsert operations for both single-key and composite-key tables (upsert.ts)
- Modular CAGR calculation with turnaround flag tracking
- Control Center schema discovery from JSON (API or Drive mock)

**Architecture Improvements:**
- One main export per file (eliminates code duplication)
- Pure functions for testable logic (no hidden globals)
- Consistent error handling and logging
- Batch operations for performance (<2000 cells per write)
- Header-based column resolution (no hard-coded indices)

**Code Quality:**
- Full TypeScript type annotations
- JSDoc documentation for all exported functions
- Consistent naming conventions (camelCase vars, PascalCase exports)
- Explicit imports/exports (no global namespace pollution)

**Backward Compatibility:**
- All menu items preserved with same functionality
- Legacy function aliases maintained (e.g., computeAnnualOpsMetrics_AndUpsert)
- Sheet contracts unchanged (Per_Share, Calculated_Metrics, etc.)

**Next Steps:**
- Delete old .js files after verification
- Add Vitest unit tests for pure functions
- Implement prices extraction module
- Add screener filter logic

---

### 2025-10-24 — Initial Setup

- Integrated clasp sync with Google Apps Script
- Initialized local Git repo and connected to GitHub
- Added PROJECT_RULES.md, SHEET_CONTRACTS.md, CURSOR_RULES.md
- Generated .cursorrules for CursorAI enforcement
- Verified clasp push and pull connectivity