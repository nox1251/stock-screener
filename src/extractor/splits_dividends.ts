/**
 * extractor/splits_dividends.ts — Stock splits and dividends extraction
 * =======================================================================
 * Purpose:
 *   - Extract splits and dividends data from EODHD API or Drive mock
 *   - Write to Raw_Splits_Dividends sheet
 *   - Support upsert by (Ticker, Kind, Date)
 * 
 * Dependencies:
 *   - globals.ts
 *   - utils/sheets.ts, utils/upsert.ts
 * 
 * Called by: Menu entrypoints in code.ts
 */

import {
  SHEET_NAMES,
  API_CONFIG,
  getEODHDApiKey,
  getMockFolderId,
  normalizeTickerSymbol
} from '../globals';

import {
  getSheet,
  ensureHeaders,
  getTickers
} from '../utils/sheets';

import {
  buildKeyMap,
  upsertLongTable
} from '../utils/upsert';

import { nullToEmpty } from '../utils/format';

/**
 * Configuration for splits/dividends extraction
 */
const SPLITS_DIVS_CONFIG = {
  HEADERS: [
    'Ticker', 'Kind', 'Date', 'DeclarationDate', 'RecordDate', 'PaymentDate',
    'Dividend', 'AdjDividend', 'ForFactor', 'ToFactor', 'Ratio', 'Notes'
  ],
  KEY_COLUMNS: ['Ticker', 'Kind', 'Date'],
  LOG_PREFIX: '[Extractor:SplitsDividends] '
} as const;

/**
 * Extracts splits and dividends for all tickers.
 */
export function extractSplitsDividends(): void {
  logMessage('Starting splits and dividends extraction...');
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAMES.RAW_SPLITS_DIVIDENDS);
  
  if (!sheet) {
    logMessage(`Sheet "${SHEET_NAMES.RAW_SPLITS_DIVIDENDS}" not found - skipping.`);
    return;
  }
  
  // Ensure headers
  ensureHeaders(sheet, SPLITS_DIVS_CONFIG.HEADERS);
  
  // Get tickers
  const tickers = getTickers(SHEET_NAMES.TICKERS);
  
  // Build existing key map
  const existingKeys = buildKeyMap(sheet, SPLITS_DIVS_CONFIG.KEY_COLUMNS);
  
  // Process in batches
  const batchSize = API_CONFIG.BATCH_SIZE;
  let totalRows = 0;
  
  for (let i = 0; i < tickers.length; i += batchSize) {
    const batch = tickers.slice(i, i + batchSize);
    const rows: any[][] = [];
    
    for (const ticker of batch) {
      try {
        const json = fetchFundamentalsJSON(ticker);
        
        if (!json || !json.SplitsDividends) {
          logMessage(`No splits/dividends data for ${ticker}`);
          continue;
        }
        
        const extracted = extractSplitsDividendsFromJSON(json, ticker);
        rows.push(...extracted);
        
      } catch (e: any) {
        logMessage(`Error extracting ${ticker}: ${e.message}`);
      }
      
      // Throttle requests
      Utilities.sleep(API_CONFIG.REQUEST_THROTTLE_MS);
    }
    
    // Upsert batch
    if (rows.length > 0) {
      upsertLongTable(
        sheet,
        existingKeys,
        SPLITS_DIVS_CONFIG.HEADERS,
        rows,
        SPLITS_DIVS_CONFIG.KEY_COLUMNS
      );
      totalRows += rows.length;
      logMessage(`Batch ${Math.floor(i / batchSize) + 1}: upserted ${rows.length} rows`);
    }
  }
  
  logMessage(`Total rows processed: ${totalRows} for ${tickers.length} tickers.`);
}

/**
 * Extracts splits and dividends from JSON.
 * @param {any} json - EODHD fundamentals JSON
 * @param {string} ticker - Ticker symbol
 * @returns {any[][]} Array of row arrays
 */
function extractSplitsDividendsFromJSON(json: any, ticker: string): any[][] {
  const rows: any[][] = [];
  const sd = json.SplitsDividends;
  
  // Extract dividends
  const dividends = Array.isArray(sd.Dividends) ? sd.Dividends : [];
  for (const d of dividends) {
    rows.push([
      ticker,
      'Dividend',
      nullToEmpty(d.date),
      nullToEmpty(d.declarationDate),
      nullToEmpty(d.recordDate),
      nullToEmpty(d.paymentDate),
      toNumOrEmpty(d.dividend),
      toNumOrEmpty(d.adjDividend),
      '', // ForFactor (not applicable for dividends)
      '', // ToFactor (not applicable for dividends)
      '', // Ratio (not applicable for dividends)
      `decl=${nullToEmpty(d.declarationDate)}`
    ]);
  }
  
  // Extract splits
  const splits = Array.isArray(sd.Splits) ? sd.Splits : [];
  for (const s of splits) {
    rows.push([
      ticker,
      'Split',
      nullToEmpty(s.date),
      '', // DeclarationDate (not in splits data)
      '', // RecordDate (not in splits data)
      '', // PaymentDate (not in splits data)
      '', // Dividend (not applicable for splits)
      '', // AdjDividend (not applicable for splits)
      toNumOrEmpty(s.forFactor),
      toNumOrEmpty(s.toFactor),
      nullToEmpty(s.ratio),
      ''
    ]);
  }
  
  // Extract last split metadata (optional)
  if (sd.LastSplitDate && sd.LastSplitFactor) {
    rows.push([
      ticker,
      'LastSplitMeta',
      nullToEmpty(sd.LastSplitDate),
      '', '', '',
      '', '',
      '', '',
      nullToEmpty(sd.LastSplitFactor),
      'from LastSplit*'
    ]);
  }
  
  return rows;
}

/**
 * Fetches fundamentals JSON (which includes SplitsDividends).
 * @param {string} ticker - Ticker symbol
 * @returns {any} Parsed JSON
 */
function fetchFundamentalsJSON(ticker: string): any {
  const mockFolderId = getMockFolderId();
  const useMock = mockFolderId.length > 0;
  
  return useMock 
    ? fetchFromDriveMock(ticker, mockFolderId) 
    : fetchFromEODHDApi(ticker);
}

/**
 * Fetches from EODHD API.
 * @param {string} ticker - Ticker symbol
 * @returns {any} Parsed JSON
 */
function fetchFromEODHDApi(ticker: string): any {
  const apiKey = getEODHDApiKey();
  const normalizedTicker = normalizeTickerSymbol(ticker);
  
  const url = `${API_CONFIG.EODHD_BASE_URL}/${API_CONFIG.EODHD_FUNDAMENTALS_ENDPOINT}/${encodeURIComponent(normalizedTicker)}?api_token=${encodeURIComponent(apiKey)}&fmt=json`;
  
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(
      `API error ${response.getResponseCode()} for ${normalizedTicker}: ` +
      response.getContentText().slice(0, 200)
    );
  }
  
  try {
    return JSON.parse(response.getContentText());
  } catch (e: any) {
    throw new Error(`JSON parse error for ${normalizedTicker}: ${e.message}`);
  }
}

/**
 * Fetches from Drive mock folder.
 * @param {string} ticker - Ticker symbol
 * @param {string} folderId - Drive folder ID
 * @returns {any} Parsed JSON
 */
function fetchFromDriveMock(ticker: string, folderId: string): any {
  const sanitizedId = sanitizeFolderId(folderId);
  
  let folder: GoogleAppsScript.Drive.Folder;
  try {
    folder = DriveApp.getFolderById(sanitizedId);
  } catch (e: any) {
    throw new Error(
      'Could not access Mock Folder. Ensure:\n' +
      '• appsscript.json includes Drive scope\n' +
      '• You have authorized Drive access\n' +
      `• Folder ID is correct: ${sanitizedId}\n\n` +
      `Error: ${e.message}`
    );
  }
  
  const normalizedTicker = normalizeTickerSymbol(ticker);
  const fileName = `${normalizedTicker}.json`;
  
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    throw new Error(`Mock JSON not found: ${fileName}`);
  }
  
  const content = files.next().getBlob().getDataAsString('utf-8');
  
  try {
    return JSON.parse(content);
  } catch (e: any) {
    throw new Error(`Mock JSON parse error for ${fileName}: ${e.message}`);
  }
}

/**
 * Sanitizes a Drive folder ID.
 * @param {string} input - Raw folder ID or URL
 * @returns {string} Clean folder ID
 */
function sanitizeFolderId(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/[-\w]{25,}/);
  return match ? match[0] : trimmed;
}

/**
 * Converts to number or empty string for sheet output.
 * @param {any} v - Value to convert
 * @returns {number | string} Number or empty string
 */
function toNumOrEmpty(v: any): number | string {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

/**
 * Logs a message with prefix.
 * @param {string} message - Message to log
 */
function logMessage(message: string): void {
  console.log(SPLITS_DIVS_CONFIG.LOG_PREFIX + message);
}

