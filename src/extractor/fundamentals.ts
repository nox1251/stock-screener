/**
 * extractor/fundamentals.ts — EODHD fundamentals extraction
 * ===========================================================
 * Purpose:
 *   - Extract annual and quarterly financial data from EODHD API or Drive mock
 *   - Honor Control Center field selections (Active? = Yes only)
 *   - Write to Raw_Fundamentals_Annual and Raw_Fundamentals_Quarterly
 *   - Support upsert by (Ticker, Date, Section, Field)
 * 
 * Dependencies: 
 *   - globals.ts
 *   - utils/sheets.ts, utils/upsert.ts, utils/format.ts
 *   - control_center/field_config.ts (for loading Control Center selections)
 * 
 * Called by: Menu entrypoints in code.ts
 */

import {
  SHEET_NAMES,
  API_CONFIG,
  FINANCIAL_SECTIONS,
  PERIOD_TYPES,
  getEODHDApiKey,
  getMockFolderId,
  normalizeTickerSymbol
} from '../globals';

import {
  ensureSheetWithHeaders,
  getTickers,
  type HeaderIndex
} from '../utils/sheets';

import {
  buildKeyMap,
  upsertLongTable
} from '../utils/upsert';

/**
 * Configuration for fundamentals extraction
 */
const FUNDAMENTALS_CONFIG = {
  ANNUAL_HEADERS: ['Ticker', 'Date', 'FiscalYear', 'Section', 'Field', 'Value', 'Source'],
  QUARTERLY_HEADERS: ['Ticker', 'Date', 'FiscalYear', 'FiscalQuarter', 'Section', 'Field', 'Value', 'Source'],
  SOURCE_TAG: 'EODHD',
  LOG_PREFIX: '[Extractor:Fundamentals] '
} as const;

/**
 * Field configuration from Control Center
 */
export interface FieldConfig {
  annual: {
    Income_Statement: Set<string>;
    Balance_Sheet: Set<string>;
    Cash_Flow: Set<string>;
  };
  quarterly: {
    Income_Statement: Set<string>;
    Balance_Sheet: Set<string>;
    Cash_Flow: Set<string>;
  };
}

/**
 * Extracts annual fundamentals for all tickers.
 * Reads field configuration from Control Center.
 */
export function extractAnnualFundamentals(): void {
  logMessage('Starting annual fundamentals extraction...');
  
  const tickers = getTickers(SHEET_NAMES.TICKERS);
  const fieldConfig = loadFieldConfigFromControlCenter();
  
  // Check if any annual fields are active
  const hasActiveFields = Object.keys(fieldConfig.annual).some(
    section => fieldConfig.annual[section as keyof typeof fieldConfig.annual].size > 0
  );
  
  if (!hasActiveFields) {
    logMessage('No annual fields marked Active in Control Center - skipping.');
    return;
  }
  
  extractFundamentalsForPeriod(tickers, fieldConfig.annual, 'annual');
  
  logMessage(`Completed annual extraction for ${tickers.length} tickers.`);
}

/**
 * Extracts quarterly fundamentals for all tickers.
 * Reads field configuration from Control Center.
 */
export function extractQuarterlyFundamentals(): void {
  logMessage('Starting quarterly fundamentals extraction...');
  
  const tickers = getTickers(SHEET_NAMES.TICKERS);
  const fieldConfig = loadFieldConfigFromControlCenter();
  
  // Check if any quarterly fields are active
  const hasActiveFields = Object.keys(fieldConfig.quarterly).some(
    section => fieldConfig.quarterly[section as keyof typeof fieldConfig.quarterly].size > 0
  );
  
  if (!hasActiveFields) {
    logMessage('No quarterly fields marked Active in Control Center - skipping.');
    return;
  }
  
  extractFundamentalsForPeriod(tickers, fieldConfig.quarterly, 'quarterly');
  
  logMessage(`Completed quarterly extraction for ${tickers.length} tickers.`);
}

/**
 * Extracts both annual and quarterly fundamentals.
 */
export function extractAllFundamentals(): void {
  extractAnnualFundamentals();
  extractQuarterlyFundamentals();
}

/**
 * Core extraction logic for a specific period (annual or quarterly).
 * @param {string[]} tickers - Ticker symbols to extract
 * @param {Object} wantedFields - Field configuration by section
 * @param {string} period - 'annual' or 'quarterly'
 */
function extractFundamentalsForPeriod(
  tickers: string[],
  wantedFields: any,
  period: 'annual' | 'quarterly'
): void {
  const isAnnual = period === 'annual';
  const sheetName = isAnnual 
    ? SHEET_NAMES.RAW_FUNDAMENTALS_ANNUAL 
    : SHEET_NAMES.RAW_FUNDAMENTALS_QUARTERLY;
  
  const headers = isAnnual 
    ? FUNDAMENTALS_CONFIG.ANNUAL_HEADERS 
    : FUNDAMENTALS_CONFIG.QUARTERLY_HEADERS;
  
  // Ensure target sheet exists
  const sheet = ensureSheetWithHeaders(sheetName, headers);
  
  // Build existing key map for upsert
  const keyColumns = ['Ticker', 'Date', 'Section', 'Field'];
  const existingKeys = buildKeyMap(sheet, keyColumns);
  
  // Process tickers in batches
  const batchSize = API_CONFIG.BATCH_SIZE;
  let totalRows = 0;
  
  for (let i = 0; i < tickers.length; i += batchSize) {
    const batch = tickers.slice(i, i + batchSize);
    const rows: any[][] = [];
    
    for (const ticker of batch) {
      try {
        const json = fetchFundamentalsJSON(ticker);
        
        if (!json || !json.Financials) {
          logMessage(`No financials data for ${ticker}`);
          continue;
        }
        
        const extracted = extractFieldsFromJSON(
          json,
          ticker,
          wantedFields,
          isAnnual
        );
        
        rows.push(...extracted);
        
      } catch (e: any) {
        logMessage(`Error extracting ${ticker}: ${e.message}`);
      }
      
      // Throttle requests
      Utilities.sleep(API_CONFIG.REQUEST_THROTTLE_MS);
    }
    
    // Upsert batch
    if (rows.length > 0) {
      upsertLongTable(sheet, existingKeys, headers, rows, keyColumns);
      totalRows += rows.length;
      logMessage(`Batch ${Math.floor(i / batchSize) + 1}: upserted ${rows.length} rows`);
    }
  }
  
  logMessage(`Total rows processed: ${totalRows}`);
}

/**
 * Extracts wanted fields from a fundamentals JSON response.
 * @param {any} json - EODHD fundamentals JSON
 * @param {string} ticker - Ticker symbol
 * @param {Object} wantedFields - Fields to extract by section
 * @param {boolean} isAnnual - True for annual, false for quarterly
 * @returns {any[][]} Array of row arrays
 */
function extractFieldsFromJSON(
  json: any,
  ticker: string,
  wantedFields: any,
  isAnnual: boolean
): any[][] {
  const rows: any[][] = [];
  const financials = json.Financials;
  
  const sections = [
    FINANCIAL_SECTIONS.INCOME_STATEMENT,
    FINANCIAL_SECTIONS.BALANCE_SHEET,
    FINANCIAL_SECTIONS.CASH_FLOW
  ];
  
  for (const section of sections) {
    const needed = wantedFields[section];
    if (!needed || needed.size === 0) continue;
    
    const sectionData = financials[section];
    if (!sectionData) continue;
    
    // Get yearly or quarterly data
    const periodKey = isAnnual ? PERIOD_TYPES.YEARLY : PERIOD_TYPES.QUARTERLY;
    const periods = sectionData[periodKey];
    
    if (!Array.isArray(periods)) continue;
    
    for (const periodObj of periods) {
      const date = periodObj.date || '';
      const fiscalYear = periodObj.fiscalYear || '';
      const fiscalQuarter = periodObj.fiscalQuarter || '';
      
      // Extract each wanted field
      for (const field of needed) {
        if (Object.prototype.hasOwnProperty.call(periodObj, field)) {
          const value = periodObj[field];
          
          if (isAnnual) {
            rows.push([
              ticker,
              date,
              fiscalYear,
              section,
              field,
              value,
              FUNDAMENTALS_CONFIG.SOURCE_TAG
            ]);
          } else {
            rows.push([
              ticker,
              date,
              fiscalYear,
              fiscalQuarter,
              section,
              field,
              value,
              FUNDAMENTALS_CONFIG.SOURCE_TAG
            ]);
          }
        }
      }
    }
  }
  
  return rows;
}

/**
 * Fetches fundamentals JSON for a ticker from API or mock.
 * @param {string} ticker - Ticker symbol
 * @returns {any} Parsed JSON object
 */
function fetchFundamentalsJSON(ticker: string): any {
  const mockFolderId = getMockFolderId();
  const useMock = mockFolderId.length > 0;
  
  return useMock 
    ? fetchFromDriveMock(ticker, mockFolderId) 
    : fetchFromEODHDApi(ticker);
}

/**
 * Fetches fundamentals from EODHD API.
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
 * Fetches fundamentals from Drive mock folder.
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
  
  // Determine filename
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
 * Loads field configuration from Control Center sheet.
 * Returns which fields are marked Active for annual and quarterly.
 * @returns {FieldConfig} Field configuration
 */
function loadFieldConfigFromControlCenter(): FieldConfig {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONTROL_CENTER);
  
  if (!sheet) {
    throw new Error('Control Center sheet not found.');
  }
  
  const HEADER_ROW = 9;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  const config: FieldConfig = {
    annual: {
      Income_Statement: new Set(),
      Balance_Sheet: new Set(),
      Cash_Flow: new Set()
    },
    quarterly: {
      Income_Statement: new Set(),
      Balance_Sheet: new Set(),
      Cash_Flow: new Set()
    }
  };
  
  if (lastRow <= HEADER_ROW) {
    return config;
  }
  
  // Read headers
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(v => String(v || '').trim());
  
  const colIndices = {
    Section: headers.indexOf('Section'),
    Field: headers.indexOf('Field Name'),
    Period: headers.indexOf('Period Type'),
    Active: headers.indexOf('Active?')
  };
  
  // Validate headers
  if (colIndices.Section < 0 || colIndices.Field < 0 || 
      colIndices.Period < 0 || colIndices.Active < 0) {
    throw new Error(
      'Control Center must have headers: Section, Field Name, Period Type, Active?'
    );
  }
  
  // Read data rows
  const values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();
  
  for (const row of values) {
    const active = String(row[colIndices.Active] || '').toLowerCase() === 'yes';
    if (!active) continue;
    
    const section = String(row[colIndices.Section] || '').trim();
    const field = String(row[colIndices.Field] || '').trim();
    const period = String(row[colIndices.Period] || '').trim().toLowerCase();
    
    if (!section || !field || !period) continue;
    
    // Validate section
    if (section !== FINANCIAL_SECTIONS.INCOME_STATEMENT &&
        section !== FINANCIAL_SECTIONS.BALANCE_SHEET &&
        section !== FINANCIAL_SECTIONS.CASH_FLOW) {
      continue;
    }
    
    // Add to appropriate period config
    if (period === 'annual') {
      config.annual[section as keyof typeof config.annual].add(field);
    } else if (period === 'quarterly') {
      config.quarterly[section as keyof typeof config.quarterly].add(field);
    }
  }
  
  return config;
}

/**
 * Sanitizes a Drive folder ID from various input formats.
 * @param {string} input - Raw folder ID or URL
 * @returns {string} Clean folder ID
 */
function sanitizeFolderId(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/[-\w]{25,}/);
  return match ? match[0] : trimmed;
}

/**
 * Logs a message with prefix.
 * @param {string} message - Message to log
 */
function logMessage(message: string): void {
  console.log(FUNDAMENTALS_CONFIG.LOG_PREFIX + message);
}

