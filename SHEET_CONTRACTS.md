
---

### **`SHEET_CONTRACTS.md`**
```md
# Sheet Contracts — Stock Screener

## 1. Raw_Fundamentals_Annual
| Column | Type | Notes |
|---------|------|-------|
| Ticker | text | PSE/SGX/US symbol |
| FiscalYear | number (YYYY) | Unique per ticker |
| Revenue | number | annual total |
| GrossProfit | number | annual total |
| OperatingIncome | number | annual total |
| NetIncome | number | annual total |
| TotalStockholdersEquity | number | balance-sheet value |
| TotalDebt | number | balance-sheet value |
| SharesDiluted | number | weighted average shares |
**Primary key:** Ticker + FiscalYear  
**Purpose:** raw annual fundamentals from EODHD.

---

## 2. Per_Share
| Column | Type | Notes |
|---------|------|-------|
| Ticker | text | from source |
| FiscalYear | number | same as above |
| RevenuePerShare | number | Revenue / SharesDiluted |
| GrossProfitPerShare | number | GrossProfit / SharesDiluted |
| OperatingIncomePerShare | number | OperatingIncome / SharesDiluted |
| NetIncomePerShare | number | NetIncome / SharesDiluted |
| EquityPerShare | number | Equity / SharesDiluted |
| DebtToEquity | number | Debt / Equity |
**Primary key:** Ticker + FiscalYear  
**Expected rows:** latest 10 FY per ticker.  
**Upstream:** `Raw_Fundamentals_Annual`.

---

## 3. Calculated_Metrics
| Column | Type | Notes |
|---------|------|-------|
| Ticker | text | unique |
| OpPS_Latest | number | latest Operating Income per share |
| OpPS_5Y_CAGR | number | 5-year CAGR |
| OpPS_10Y_CAGR | number | 10-year CAGR |
| OpPS_AdjustedFlag | boolean | TRUE if ≤0→0.01 rule applied |
| EPS_Latest | number | latest EPS |
| EPS_5Y_CAGR | number | 5-year CAGR |
| EPS_10Y_CAGR | number | 10-year CAGR |
| EPS_AdjustedFlag | boolean | same logic |
| FCFPS_Latest | number | latest FCF per share |
| FCFPS_5Y_CAGR | number | 5-year CAGR |
| FCFPS_10Y_CAGR | number | 10-year CAGR |
| FCFPS_AdjustedFlag | boolean | same logic |
**Primary key:** Ticker.  
**Upstream:** `Per_Share`.  
**Purpose:** derived metrics used in screener filters.

---

## 4. Prices
| Column | Type | Notes |
|---------|------|-------|
| Ticker | text | symbol |
| Date | date | trading date |
| Close | number | EOD close |
| Volume | number | daily volume |
**Primary key:** Ticker + Date.  
**Purpose:** price history for momentum/volatility metrics.

---

## 5. Splits_Dividends
| Column | Type | Notes |
|---------|------|-------|
| Ticker | text | symbol |
| Date | date | event date |
| Type | text | 'Split' or 'Dividend' |
| Ratio_or_Amount | number | value |
**Primary key:** Ticker + Date + Type.  
**Purpose:** corporate actions history.

---

## 6. Control_Center
| Column | Type | Notes |
|---------|------|-------|
| Section | text | e.g., IncomeStatement, BalanceSheet |
| FieldName | text | user-selected metric |
| JSONPath | text | key path from EODHD API |
| PeriodType | text | annual/quarterly |
**Purpose:** user-maintained mapping of API fields to sheet columns.
