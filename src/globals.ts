/**
 * globals.ts — Centralized constants and configuration
 * =====================================================
 * Purpose:
 *   - Define all sheet names, API endpoints, and configuration constants
 *   - Provide API key loader from Script Properties
 *   - Single source of truth for project-wide constants
 * 
 * Dependencies: None
 * Called by: All modules
 */

// ==================== Sheet Names ====================

export const SHEET_NAMES = {
  // Source data
  TICKERS: 'Tickers',
  CONTROL_CENTER: 'Control Center',
  
  // Raw data from API/Mock
  RAW_FUNDAMENTALS_ANNUAL: 'Raw_Fundamentals_Annual',
  RAW_FUNDAMENTALS_QUARTERLY: 'Raw_Fundamentals_Quarterly',
  RAW_SPLITS_DIVIDENDS: 'Raw_Splits_Dividends',
  RAW_PRICES: 'Raw_Prices',
  
  // Processed data
  PER_SHARE: 'Per_Share',
  CALCULATED_METRICS: 'Calculated_Metrics',
  PRICES_LATEST: 'Prices_Latest',
  PRICES_STATS: 'Prices_Stats',
  
  // Screener
  SCREENER: 'Screener',
  
  // Internal lookups
  CC_LOOKUPS: 'CC_Lookups'
} as const;

// ==================== API Configuration ====================

export const API_CONFIG = {
  EODHD_BASE_URL: 'https://eodhd.com/api',
  EODHD_FUNDAMENTALS_ENDPOINT: 'fundamentals',
  EODHD_PRICES_ENDPOINT: 'eod',
  SCRIPT_PROPERTY_KEY: 'EODHD_API_KEY',
  DEFAULT_EXCHANGE_SUFFIX: '.PSE',
  REQUEST_THROTTLE_MS: 250,
  BATCH_SIZE: 10
} as const;

// ==================== Financial Sections ====================

export const FINANCIAL_SECTIONS = {
  INCOME_STATEMENT: 'Income_Statement',
  BALANCE_SHEET: 'Balance_Sheet',
  CASH_FLOW: 'Cash_Flow'
} as const;

export const PERIOD_TYPES = {
  ANNUAL: 'annual',
  QUARTERLY: 'quarterly',
  YEARLY: 'yearly'  // EODHD uses 'yearly' in JSON
} as const;

// ==================== Per-Share Field Configuration ====================

export const PER_SHARE_FIELDS = {
  // Income Statement fields
  IS: {
    Revenue: 'Revenue',
    GrossProfit: 'GrossProfit',
    OperatingIncome: 'OperatingIncome',
    NetIncome: 'NetIncome',
    SharesDiluted: 'SharesDiluted'
  },
  
  // Balance Sheet fields
  BS: {
    Equity: 'TotalStockholdersEquity',
    Debt: 'TotalDebt',
    SharesDiluted: 'SharesDiluted',
    
    // Accepted fallback field names
    EquityAlts: ['TotalStockholderEquity', 'TotalEquity', 'StockholdersEquity'],
    DebtAlts: ['TotalLiab', 'TotalLiabilities']
  }
} as const;

// ==================== CAGR Configuration ====================

export const CAGR_CONFIG = {
  // Apply ≤0 → 0.01 rule for turnaround metrics
  FLOOR_VALUE: 0.01,
  
  // CAGR period requirements (minimum rows needed)
  MIN_ROWS_5Y: 6,   // Latest + 5 years back
  MIN_ROWS_9Y: 10,  // Latest + 9 years back
  MIN_ROWS_10Y: 11, // Latest + 10 years back
  
  // Rounding precision
  DECIMAL_PLACES: 4,
  
  // Metrics to compute
  METRICS: ['OpPS', 'EPS', 'FCFPS'] as const
} as const;

// ==================== Control Center Configuration ====================

export const CONTROL_CENTER_CONFIG = {
  HEADER_ROW: 9,
  MOCK_FOLDER_ID_CELL: 'B7',
  
  COLUMNS: {
    SECTION: 1,      // A
    FIELD_NAME: 2,   // B
    PERIOD_TYPE: 3,  // C
    ACTIVE: 4,       // D
    NOTES: 5         // E
  },
  
  MAX_DISCOVERY_FILES: 12
} as const;

// ==================== Data Formats ====================

export const NUMBER_FORMATS = {
  DECIMAL_3: '#,##0.000',
  DECIMAL_2: '#,##0.00',
  PERCENT_2: '0.00%',
  TIMESTAMP: 'yyyy-mm-dd hh:mm:ss',
  DATE: 'yyyy-mm-dd'
} as const;

// ==================== Batch Processing ====================

export const BATCH_CONFIG = {
  MAX_CELLS_PER_WRITE: 2000,
  SLEEP_BETWEEN_BATCHES_MS: 50
} as const;

// ==================== API Key Management ====================

/**
 * Retrieves the EODHD API key from Script Properties.
 * @returns {string} The API key
 * @throws {Error} If API key is not set
 */
export function getEODHDApiKey(): string {
  const key = PropertiesService.getScriptProperties()
    .getProperty(API_CONFIG.SCRIPT_PROPERTY_KEY);
  
  if (!key || key.trim() === '') {
    throw new Error(
      `EODHD API key not found. Please set it via:\n` +
      `Settings → Set API Key (EODHD)`
    );
  }
  
  return key.trim();
}

/**
 * Sets the EODHD API key in Script Properties.
 * @param {string} key - The API key to store
 */
export function setEODHDApiKey(key: string): void {
  if (!key || key.trim() === '') {
    throw new Error('API key cannot be empty');
  }
  
  PropertiesService.getScriptProperties()
    .setProperty(API_CONFIG.SCRIPT_PROPERTY_KEY, key.trim());
}

/**
 * Checks if the EODHD API key is configured.
 * @returns {boolean} True if key exists
 */
export function hasEODHDApiKey(): boolean {
  const key = PropertiesService.getScriptProperties()
    .getProperty(API_CONFIG.SCRIPT_PROPERTY_KEY);
  return !!(key && key.trim());
}

// ==================== Mock Mode Configuration ====================

/**
 * Gets the Mock Prices folder ID from Control Center sheet (B7).
 * @returns {string} The folder ID or empty string if not set
 */
export function getMockFolderId(): string {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAMES.CONTROL_CENTER);
    if (!sh) return '';
    
    const raw = sh.getRange(CONTROL_CENTER_CONFIG.MOCK_FOLDER_ID_CELL)
      .getValue();
    return String(raw || '').trim();
  } catch (e) {
    Logger.log('Error reading Mock Folder ID from Control Center: ' + e);
    return '';
  }
}

/**
 * Checks if mock mode is enabled by looking for a valid folder ID.
 * @returns {boolean} True if mock mode folder ID is set
 */
export function isMockModeEnabled(): boolean {
  const folderId = getMockFolderId();
  return folderId.length > 0;
}

// ==================== Ticker Utilities ====================

/**
 * Normalizes ticker symbol by adding exchange suffix if missing.
 * @param {string} ticker - Raw ticker symbol
 * @param {string} suffix - Exchange suffix (default: .PSE)
 * @returns {string} Normalized ticker with suffix
 */
export function normalizeTickerSymbol(
  ticker: string,
  suffix: string = API_CONFIG.DEFAULT_EXCHANGE_SUFFIX
): string {
  const clean = ticker.trim();
  return clean.includes('.') ? clean : `${clean}${suffix}`;
}

/**
 * Extracts the base ticker symbol without exchange suffix.
 * @param {string} ticker - Ticker with or without suffix
 * @returns {string} Base ticker symbol
 */
export function getBaseTicker(ticker: string): string {
  const parts = ticker.split('.');
  return parts[0];
}

