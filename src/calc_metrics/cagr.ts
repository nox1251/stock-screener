/**
 * calc_metrics/cagr.ts — CAGR calculations with turnaround handling
 * ==================================================================
 * Purpose:
 *   - Read Per_Share sheet (latest ~10 FY per ticker)
 *   - Compute Latest, 5Y, 9Y, and 10Y CAGRs for OpPS, EPS, FCFPS
 *   - Apply ≤0 → 0.01 floor rule (turnaround handling)
 *   - Set AdjustedFlag when floor rule is applied
 *   - Upsert to Calculated_Metrics sheet
 * 
 * Dependencies:
 *   - globals.ts
 *   - utils/sheets.ts, utils/upsert.ts, utils/format.ts
 * 
 * Called by: Menu entrypoints in code.ts
 */

import {
  SHEET_NAMES,
  CAGR_CONFIG
} from '../globals';

import {
  getSheet,
  readTable,
  resolveColumns
} from '../utils/sheets';

import {
  toNum,
  toInt,
  floor01,
  round4
} from '../utils/format';

import {
  upsertCalculatedMetrics
} from '../utils/upsert';

/**
 * Configuration for CAGR calculations
 */
const CALC_CONFIG = {
  SOURCE_SHEET: SHEET_NAMES.PER_SHARE,
  OUTPUT_SHEET: SHEET_NAMES.CALCULATED_METRICS,
  
  // Column name aliases for flexible header matching
  COLUMN_ALIASES: {
    Ticker: ['Ticker', 'Symbol'],
    FiscalYear: ['FiscalYear', 'FY', 'Year'],
    OpPS: [
      'OpPS', 'OpIncPS', 'OperatingIncomePS', 'OperatingIncomePerShare',
      'Operating Income/Share', 'OpInc/Share', 'OpIncomePS'
    ],
    EPS: ['EPS', 'EarningsPS', 'EarningsPerShare', 'NetIncomePerShare'],
    FCFPS: ['FCFPS', 'FCF Per Share', 'FreeCashFlowPS', 'FCF/Share', 'FCF_PS']
  },
  
  LOG_PREFIX: '[CalcMetrics:CAGR] '
} as const;

/**
 * Ticker time series data
 */
interface TickerSeries {
  [ticker: string]: YearRecord[];
}

/**
 * Single year record
 */
interface YearRecord {
  fy: number;
  OpPS: number | null;
  EPS: number | null;
  FCFPS: number | null;
}

/**
 * Calculated metrics for a ticker
 */
interface CalculatedMetrics {
  Ticker: string;
  
  OpPS_Latest: number | null;
  OpPS_5Y_CAGR: number | null;
  OpPS_9Y_CAGR: number | null;
  OpPS_10Y_CAGR: number | null;
  OpPS_AdjustedFlag: boolean;
  
  EPS_Latest: number | null;
  EPS_5Y_CAGR: number | null;
  EPS_9Y_CAGR: number | null;
  EPS_10Y_CAGR: number | null;
  EPS_AdjustedFlag: boolean;
  
  FCFPS_Latest: number | null;
  FCFPS_5Y_CAGR: number | null;
  FCFPS_9Y_CAGR: number | null;
  FCFPS_10Y_CAGR: number | null;
  FCFPS_AdjustedFlag: boolean;
  
  Calc_Timestamp: Date;
}

/**
 * Computes CAGRs for all tickers and upserts to Calculated_Metrics.
 * Main entry point called from menu.
 */
export function computeCAGRs(): void {
  logMessage('Starting CAGR computation...');
  
  const ss = SpreadsheetApp.getActive();
  
  // Read Per_Share data
  const perShareSheet = ss.getSheetByName(CALC_CONFIG.SOURCE_SHEET);
  if (!perShareSheet) {
    throw new Error(`Sheet "${CALC_CONFIG.SOURCE_SHEET}" not found.`);
  }
  
  const table = readTable(perShareSheet);
  
  // Resolve column indices with alias support
  const cols = resolveColumns(table.header, CALC_CONFIG.COLUMN_ALIASES);
  
  // Index data by ticker
  const byTicker = indexByTicker(table.rows, cols);
  
  // Calculate metrics for each ticker
  const metricsMap = new Map<string, CalculatedMetrics>();
  const timestamp = new Date();
  
  for (const [ticker, series] of Object.entries(byTicker)) {
    if (series.length === 0) continue;
    
    const metrics = calculateTickerMetrics(ticker, series, timestamp);
    metricsMap.set(ticker, metrics);
  }
  
  // Upsert to Calculated_Metrics
  const outputSheet = getSheet(CALC_CONFIG.OUTPUT_SHEET, true);
  upsertCalculatedMetrics(outputSheet, metricsMap);
  
  logMessage(`Completed CAGR computation for ${metricsMap.size} tickers.`);
}

/**
 * Indexes Per_Share data by ticker.
 * @param {any[][]} rows - Data rows
 * @param {Object} cols - Column indices
 * @returns {TickerSeries} Time series by ticker
 */
function indexByTicker(rows: any[][], cols: any): TickerSeries {
  const byTicker: TickerSeries = {};
  
  for (const row of rows) {
    const ticker = String(row[cols.Ticker] || '').trim();
    if (!ticker) continue;
    
    const fy = toInt(row[cols.FiscalYear]);
    if (!fy) continue;
    
    const record: YearRecord = {
      fy,
      OpPS: toNum(row[cols.OpPS]),
      EPS: toNum(row[cols.EPS]),
      FCFPS: toNum(row[cols.FCFPS])
    };
    
    if (!byTicker[ticker]) {
      byTicker[ticker] = [];
    }
    
    byTicker[ticker].push(record);
  }
  
  // Sort each ticker's series by fiscal year (ascending)
  for (const series of Object.values(byTicker)) {
    series.sort((a, b) => a.fy - b.fy);
  }
  
  return byTicker;
}

/**
 * Calculates all metrics for a single ticker.
 * @param {string} ticker - Ticker symbol
 * @param {YearRecord[]} series - Time series data (sorted by year)
 * @param {Date} timestamp - Calculation timestamp
 * @returns {CalculatedMetrics} Calculated metrics
 */
function calculateTickerMetrics(
  ticker: string,
  series: YearRecord[],
  timestamp: Date
): CalculatedMetrics {
  const metrics: CalculatedMetrics = {
    Ticker: ticker,
    
    OpPS_Latest: null,
    OpPS_5Y_CAGR: null,
    OpPS_9Y_CAGR: null,
    OpPS_10Y_CAGR: null,
    OpPS_AdjustedFlag: false,
    
    EPS_Latest: null,
    EPS_5Y_CAGR: null,
    EPS_9Y_CAGR: null,
    EPS_10Y_CAGR: null,
    EPS_AdjustedFlag: false,
    
    FCFPS_Latest: null,
    FCFPS_5Y_CAGR: null,
    FCFPS_9Y_CAGR: null,
    FCFPS_10Y_CAGR: null,
    FCFPS_AdjustedFlag: false,
    
    Calc_Timestamp: timestamp
  };
  
  if (series.length === 0) return metrics;
  
  const lastIdx = series.length - 1;
  
  // Calculate for each metric
  for (const metricName of CAGR_CONFIG.METRICS) {
    const result = calculateMetricCAGRs(series, metricName as any);
    
    // Assign to metrics object
    const prefix = metricName;
    (metrics as any)[`${prefix}_Latest`] = result.latest;
    (metrics as any)[`${prefix}_5Y_CAGR`] = result.cagr5Y;
    (metrics as any)[`${prefix}_9Y_CAGR`] = result.cagr9Y;
    (metrics as any)[`${prefix}_10Y_CAGR`] = result.cagr10Y;
    (metrics as any)[`${prefix}_AdjustedFlag`] = result.adjustedFlag;
  }
  
  return metrics;
}

/**
 * Calculates CAGRs for a specific metric.
 * @param {YearRecord[]} series - Time series
 * @param {string} metricName - Metric name (OpPS, EPS, FCFPS)
 * @returns {Object} CAGR results
 */
function calculateMetricCAGRs(
  series: YearRecord[],
  metricName: 'OpPS' | 'EPS' | 'FCFPS'
): {
  latest: number | null;
  cagr5Y: number | null;
  cagr9Y: number | null;
  cagr10Y: number | null;
  adjustedFlag: boolean;
} {
  const lastIdx = series.length - 1;
  
  // Latest value
  const latestRaw = series[lastIdx][metricName];
  const latestAdj = floor01(latestRaw);
  const latestWasAdjusted = latestRaw !== null && latestRaw <= 0;
  
  let cagr5Y: number | null = null;
  let adj5Y = false;
  
  let cagr9Y: number | null = null;
  let adj9Y = false;
  
  let cagr10Y: number | null = null;
  let adj10Y = false;
  
  // 5-year CAGR (need at least 6 rows)
  if (series.length >= CAGR_CONFIG.MIN_ROWS_5Y) {
    const endIdx = lastIdx;
    const startIdx = lastIdx - 5;
    
    const endRaw = series[endIdx][metricName];
    const startRaw = series[startIdx][metricName];
    
    const endAdj = floor01(endRaw);
    const startAdj = floor01(startRaw);
    
    if (endAdj !== null && startAdj !== null && startAdj > 0) {
      cagr5Y = round4(Math.pow(endAdj / startAdj, 1 / 5) - 1);
    }
    
    adj5Y = (endRaw !== null && endRaw <= 0) || (startRaw !== null && startRaw <= 0);
  }
  
  // 9-year CAGR (need at least 10 rows)
  if (series.length >= CAGR_CONFIG.MIN_ROWS_9Y) {
    const endIdx = lastIdx;
    const startIdx = lastIdx - 9;
    
    const endRaw = series[endIdx][metricName];
    const startRaw = series[startIdx][metricName];
    
    const endAdj = floor01(endRaw);
    const startAdj = floor01(startRaw);
    
    if (endAdj !== null && startAdj !== null && startAdj > 0) {
      cagr9Y = round4(Math.pow(endAdj / startAdj, 1 / 9) - 1);
    }
    
    adj9Y = (endRaw !== null && endRaw <= 0) || (startRaw !== null && startRaw <= 0);
  }
  
  // 10-year CAGR (need at least 11 rows)
  if (series.length >= CAGR_CONFIG.MIN_ROWS_10Y) {
    const endIdx = lastIdx;
    const startIdx = lastIdx - 10;
    
    const endRaw = series[endIdx][metricName];
    const startRaw = series[startIdx][metricName];
    
    const endAdj = floor01(endRaw);
    const startAdj = floor01(startRaw);
    
    if (endAdj !== null && startAdj !== null && startAdj > 0) {
      cagr10Y = round4(Math.pow(endAdj / startAdj, 1 / 10) - 1);
    }
    
    adj10Y = (endRaw !== null && endRaw <= 0) || (startRaw !== null && startRaw <= 0);
  }
  
  // Set adjusted flag if any endpoint was floored
  const adjustedFlag = latestWasAdjusted || adj5Y || adj9Y || adj10Y;
  
  return {
    latest: latestAdj,
    cagr5Y,
    cagr9Y,
    cagr10Y,
    adjustedFlag
  };
}

/**
 * Logs a message with prefix.
 * @param {string} message - Message to log
 */
function logMessage(message: string): void {
  console.log(CALC_CONFIG.LOG_PREFIX + message);
}

/**
 * Legacy alias for backward compatibility.
 */
export function computeAnnualOpsMetrics_AndUpsert(): void {
  computeCAGRs();
}

