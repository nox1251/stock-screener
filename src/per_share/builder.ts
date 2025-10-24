/**
 * per_share/builder.ts — Per-share metrics calculation
 * ======================================================
 * Purpose:
 *   - Read Raw_Fundamentals_Annual (long table format)
 *   - Calculate per-share metrics using SharesDiluted
 *   - Handle Debt/Equity with fallback field names
 *   - Write to Per_Share sheet (latest 10 FY per ticker)
 * 
 * Dependencies:
 *   - globals.ts
 *   - utils/sheets.ts, utils/format.ts
 * 
 * Called by: Menu entrypoints in code.ts
 */

import {
  SHEET_NAMES,
  PER_SHARE_FIELDS,
  FINANCIAL_SECTIONS,
  NUMBER_FORMATS
} from '../globals';

import {
  getSheet,
  writeTable,
  autoResizeColumns,
  formatColumns
} from '../utils/sheets';

import {
  toNum,
  toInt,
  round3,
  safeDiv,
  firstNonNull
} from '../utils/format';

/**
 * Configuration for per-share builder
 */
const PER_SHARE_CONFIG = {
  SOURCE_SHEET: SHEET_NAMES.RAW_FUNDAMENTALS_ANNUAL,
  OUTPUT_SHEET: SHEET_NAMES.PER_SHARE,
  
  HEADERS: [
    'Ticker', 'FiscalYear',
    'RevenuePerShare', 'GrossProfitPerShare', 'OperatingIncomePerShare', 'NetIncomePerShare',
    'EquityPerShare', 'DebtToEquity'
  ],
  
  MAX_YEARS_PER_TICKER: 10,
  LOG_PREFIX: '[PerShare:Builder] '
} as const;

/**
 * Ticker-year data structure
 */
interface TickerYearData {
  [ticker: string]: {
    [year: number]: YearData;
  };
}

/**
 * Data for a single fiscal year
 */
interface YearData {
  is: {
    Revenue?: number;
    GrossProfit?: number;
    OperatingIncome?: number;
    NetIncome?: number;
    SharesDiluted?: number;
  };
  bs: {
    Equity?: number;
    Debt?: number;
    SharesDiluted?: number;
    EquityAlt1?: number;
    EquityAlt2?: number;
    EquityAlt3?: number;
    DebtAlt1?: number;
    DebtAlt2?: number;
  };
}

/**
 * Builds Per_Share sheet from Raw_Fundamentals_Annual.
 * Main entry point called from menu.
 */
export function buildPerShareFull(): void {
  logMessage('Starting Per_Share build...');
  
  const ss = SpreadsheetApp.getActive();
  
  // Read source data
  const sourceSheet = ss.getSheetByName(PER_SHARE_CONFIG.SOURCE_SHEET);
  if (!sourceSheet) {
    throw new Error(`Source sheet "${PER_SHARE_CONFIG.SOURCE_SHEET}" not found.`);
  }
  
  // Read and index data
  const longTable = readLongTable(sourceSheet);
  const byTickerYear = indexByTickerYear(longTable);
  
  // Calculate per-share metrics
  const outputRows = calculatePerShareMetrics(byTickerYear);
  
  // Write to output sheet
  const outputSheet = getSheet(PER_SHARE_CONFIG.OUTPUT_SHEET, true);
  writeTable(outputSheet, PER_SHARE_CONFIG.HEADERS, outputRows, true);
  
  // Apply formatting
  applyPerShareFormatting(outputSheet);
  
  const tickerCount = Object.keys(byTickerYear).length;
  logMessage(`Completed Per_Share build: ${tickerCount} tickers, ${outputRows.length} rows.`);
}

/**
 * Reads the long table format from Raw_Fundamentals_Annual.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Source sheet
 * @returns {any[]} Array of row objects
 */
function readLongTable(sheet: GoogleAppsScript.Spreadsheet.Sheet): any[] {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 6) {
    return [];
  }
  
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v || '').trim());
  
  // Find required columns
  const idx = {
    Ticker: header.indexOf('Ticker'),
    Date: header.indexOf('Date'),
    FiscalYear: header.indexOf('FiscalYear'),
    Section: header.indexOf('Section'),
    Field: header.indexOf('Field'),
    Value: header.indexOf('Value')
  };
  
  // Validate columns exist
  if (Object.values(idx).some(i => i < 0)) {
    throw new Error(
      'Raw_Fundamentals_Annual must have columns: Ticker, Date, FiscalYear, Section, Field, Value'
    );
  }
  
  const rows: any[] = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    rows.push({
      Ticker: String(row[idx.Ticker] || '').trim(),
      Date: row[idx.Date],
      FiscalYear: toInt(row[idx.FiscalYear]),
      Section: String(row[idx.Section] || '').trim(),
      Field: String(row[idx.Field] || '').trim(),
      Value: toNum(row[idx.Value])
    });
  }
  
  return rows;
}

/**
 * Indexes long table data by ticker and fiscal year.
 * @param {any[]} rows - Long table rows
 * @returns {TickerYearData} Nested structure by ticker and year
 */
function indexByTickerYear(rows: any[]): TickerYearData {
  const indexed: TickerYearData = {};
  
  const IS = PER_SHARE_FIELDS.IS;
  const BS = PER_SHARE_FIELDS.BS;
  
  for (const row of rows) {
    if (!row.Ticker || !row.FiscalYear) continue;
    
    // Normalize ticker (add .PSE if missing)
    const ticker = row.Ticker.includes('.PSE') 
      ? row.Ticker 
      : `${row.Ticker}.PSE`;
    
    const year = row.FiscalYear;
    
    // Initialize nested structure
    if (!indexed[ticker]) {
      indexed[ticker] = {};
    }
    if (!indexed[ticker][year]) {
      indexed[ticker][year] = { is: {}, bs: {} };
    }
    
    const yearData = indexed[ticker][year];
    const section = row.Section;
    const field = row.Field;
    const value = row.Value;
    
    // Bucket by section
    if (section === FINANCIAL_SECTIONS.INCOME_STATEMENT) {
      if (field === IS.Revenue) yearData.is.Revenue = value;
      else if (field === IS.GrossProfit) yearData.is.GrossProfit = value;
      else if (field === IS.OperatingIncome) yearData.is.OperatingIncome = value;
      else if (field === IS.NetIncome) yearData.is.NetIncome = value;
      else if (field === IS.SharesDiluted) yearData.is.SharesDiluted = value;
      
    } else if (section === FINANCIAL_SECTIONS.BALANCE_SHEET) {
      if (field === BS.Equity) yearData.bs.Equity = value;
      else if (field === BS.Debt) yearData.bs.Debt = value;
      else if (field === BS.SharesDiluted) yearData.bs.SharesDiluted = value;
      
      // Fallback fields
      else if (field === BS.EquityAlts[0]) yearData.bs.EquityAlt1 = value;
      else if (field === BS.EquityAlts[1]) yearData.bs.EquityAlt2 = value;
      else if (field === BS.EquityAlts[2]) yearData.bs.EquityAlt3 = value;
      else if (field === BS.DebtAlts[0]) yearData.bs.DebtAlt1 = value;
      else if (field === BS.DebtAlts[1]) yearData.bs.DebtAlt2 = value;
    }
  }
  
  return indexed;
}

/**
 * Calculates per-share metrics for all tickers.
 * @param {TickerYearData} byTickerYear - Indexed data
 * @returns {any[][]} Output rows
 */
function calculatePerShareMetrics(byTickerYear: TickerYearData): any[][] {
  const outputRows: any[][] = [];
  const tickers = Object.keys(byTickerYear).sort();
  
  for (const ticker of tickers) {
    const years = Object.keys(byTickerYear[ticker])
      .map(y => parseInt(y, 10))
      .sort((a, b) => b - a); // Descending order
    
    // Pick latest N years with valid shares
    const validYears: Array<{ year: number; shares: number; data: YearData }> = [];
    
    for (const year of years) {
      const data = byTickerYear[ticker][year];
      const shares = getValidShares(data);
      
      if (shares > 0) {
        validYears.push({ year, shares, data });
      }
      
      if (validYears.length >= PER_SHARE_CONFIG.MAX_YEARS_PER_TICKER) {
        break;
      }
    }
    
    // Sort by year ascending for output
    validYears.sort((a, b) => a.year - b.year);
    
    // Calculate metrics for each valid year
    for (const { year, shares, data } of validYears) {
      const row = calculateRowMetrics(ticker, year, shares, data);
      outputRows.push(row);
    }
  }
  
  return outputRows;
}

/**
 * Calculates metrics for a single row (ticker + fiscal year).
 * @param {string} ticker - Ticker symbol
 * @param {number} year - Fiscal year
 * @param {number} shares - Diluted shares
 * @param {YearData} data - Year data
 * @returns {any[]} Row array
 */
function calculateRowMetrics(
  ticker: string,
  year: number,
  shares: number,
  data: YearData
): any[] {
  // Extract values
  const revenue = toNum(data.is.Revenue);
  const grossProfit = toNum(data.is.GrossProfit);
  const operatingIncome = toNum(data.is.OperatingIncome);
  const netIncome = toNum(data.is.NetIncome);
  
  // Handle equity with fallbacks
  const equityRaw = firstNonNull([
    data.bs.Equity,
    data.bs.EquityAlt1,
    data.bs.EquityAlt2,
    data.bs.EquityAlt3
  ]);
  const equity = toNum(equityRaw);
  
  // Handle debt with fallbacks
  const debtRaw = firstNonNull([
    data.bs.Debt,
    data.bs.DebtAlt1,
    data.bs.DebtAlt2
  ]);
  const debt = toNum(debtRaw);
  
  // Calculate per-share values
  const revenuePS = safeDiv(revenue, shares);
  const gpPS = safeDiv(grossProfit, shares);
  const opPS = safeDiv(operatingIncome, shares);
  const niPS = safeDiv(netIncome, shares);
  const equityPS = safeDiv(equity, shares);
  
  // Calculate debt-to-equity ratio
  const debtToEquity = safeDiv(debt, equity);
  
  return [
    ticker,
    year,
    round3(revenuePS),
    round3(gpPS),
    round3(opPS),
    round3(niPS),
    round3(equityPS),
    round3(debtToEquity)
  ];
}

/**
 * Gets valid shares diluted value, preferring IS over BS.
 * @param {YearData} data - Year data
 * @returns {number} Valid shares or 0
 */
function getValidShares(data: YearData): number {
  const isShares = toNum(data.is.SharesDiluted);
  const bsShares = toNum(data.bs.SharesDiluted);
  
  const shares = isShares && isShares > 0 ? isShares : bsShares;
  
  return shares && shares > 0 ? shares : 0;
}

/**
 * Applies number formatting to Per_Share sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to format
 */
function applyPerShareFormatting(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;
  
  // Format per-share columns (columns 3-8: C-H)
  const perShareCols = [2, 3, 4, 5, 6, 7]; // 0-indexed
  formatColumns(sheet, 2, perShareCols, NUMBER_FORMATS.DECIMAL_3);
  
  // Auto-resize all columns
  autoResizeColumns(sheet);
}

/**
 * Logs a message with prefix.
 * @param {string} message - Message to log
 */
function logMessage(message: string): void {
  console.log(PER_SHARE_CONFIG.LOG_PREFIX + message);
}

