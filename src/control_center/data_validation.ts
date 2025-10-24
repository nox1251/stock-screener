/**
 * control_center/data_validation.ts — Control Center dropdown management
 * ========================================================================
 * Purpose:
 *   - Discover available sections and fields from EODHD JSON data
 *   - Build dropdown lists for Section, Field Name, Period Type, Active?
 *   - Apply data validation to Control Center sheet
 *   - Handle dynamic field dropdowns based on selected section
 * 
 * Dependencies:
 *   - globals.ts
 *   - utils/sheets.ts
 *   - extractor/fundamentals.ts (for JSON fetching)
 * 
 * Called by: Menu entrypoints and onEdit trigger
 */

import {
  SHEET_NAMES,
  CONTROL_CENTER_CONFIG,
  FINANCIAL_SECTIONS,
  getMockFolderId,
  API_CONFIG,
  normalizeTickerSymbol,
  getEODHDApiKey
} from '../globals';

import {
  getSheet,
  getTickers,
  hideSheet,
  setNamedRange
} from '../utils/sheets';

/**
 * Configuration for Control Center data validation
 */
const CC_CONFIG = {
  LOOKUP_SHEET: SHEET_NAMES.CC_LOOKUPS,
  CONTROL_CENTER_SHEET: SHEET_NAMES.CONTROL_CENTER,
  
  PERIOD_VALUES: ['Annual', 'Quarterly'],
  ACTIVE_VALUES: ['Yes', 'No'],
  
  MAX_DISCOVERY_FILES: CONTROL_CENTER_CONFIG.MAX_DISCOVERY_FILES,
  
  LOG_PREFIX: '[ControlCenter:DataValidation] '
} as const;

/**
 * Schema discovery result
 */
interface SchemaDiscovery {
  sections: string[];
  fieldsBySection: { [section: string]: Set<string> };
}

/**
 * Applies data validation to Control Center sheet.
 * Discovers schema from JSON files and builds dropdowns.
 * Main entry point called from menu.
 */
export function applyControlCenterValidations(): void {
  logMessage('Starting Control Center validation setup...');
  
  const ss = SpreadsheetApp.getActive();
  const ccSheet = ss.getSheetByName(CC_CONFIG.CONTROL_CENTER_SHEET);
  
  if (!ccSheet) {
    throw new Error(`Sheet "${CC_CONFIG.CONTROL_CENTER_SHEET}" not found.`);
  }
  
  // 1. Discover sections and fields from JSON data
  const schema = discoverSchemaFromJSON();
  logMessage(`Discovered ${schema.sections.length} sections with fields.`);
  
  // 2. Rebuild lookup sheet and named ranges
  const lookupSheet = ensureLookupSheet(ss);
  writeNamedLookups(ss, lookupSheet, schema);
  
  // 3. Apply data validation rules
  applyValidationRules(ss, ccSheet);
  
  logMessage('Control Center validations applied successfully.');
}

/**
 * Discovers available sections and fields from JSON data.
 * @returns {SchemaDiscovery} Schema with sections and fields
 */
function discoverSchemaFromJSON(): SchemaDiscovery {
  const sections = new Set<string>();
  const fieldsBySection: { [section: string]: Set<string> } = {};
  
  const mockFolderId = getMockFolderId();
  const useMock = mockFolderId.length > 0;
  
  if (useMock) {
    logMessage('Using mock mode - discovering from Drive folder...');
    discoverFromDriveFolder(mockFolderId, sections, fieldsBySection);
  } else {
    logMessage('Using API mode - discovering from EODHD API...');
    discoverFromAPI(sections, fieldsBySection);
  }
  
  return {
    sections: Array.from(sections).sort(),
    fieldsBySection
  };
}

/**
 * Discovers schema from Drive mock folder.
 * @param {string} folderId - Drive folder ID
 * @param {Set<string>} sections - Set to populate with sections
 * @param {Object} fieldsBySection - Object to populate with fields
 */
function discoverFromDriveFolder(
  folderId: string,
  sections: Set<string>,
  fieldsBySection: { [section: string]: Set<string> }
): void {
  const sanitizedId = sanitizeFolderId(folderId);
  
  let folder: GoogleAppsScript.Drive.Folder;
  try {
    folder = DriveApp.getFolderById(sanitizedId);
  } catch (e: any) {
    throw new Error(
      'Could not access Mock Folder for discovery. ' +
      'Ensure Drive scope and folder ID are correct.'
    );
  }
  
  const files = folder.getFiles();
  let count = 0;
  
  while (files.hasNext() && count < CC_CONFIG.MAX_DISCOVERY_FILES) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();
    
    if (!fileName.endsWith('.json')) continue;
    
    try {
      const content = file.getBlob().getDataAsString('utf-8');
      const json = JSON.parse(content);
      accumulateSchema(json, sections, fieldsBySection);
      count++;
    } catch (e) {
      logMessage(`Skipping ${file.getName()}: parse error`);
    }
  }
  
  logMessage(`Discovered from ${count} JSON files.`);
}

/**
 * Discovers schema from EODHD API.
 * @param {Set<string>} sections - Set to populate with sections
 * @param {Object} fieldsBySection - Object to populate with fields
 */
function discoverFromAPI(
  sections: Set<string>,
  fieldsBySection: { [section: string]: Set<string> }
): void {
  const tickers = getTickers(SHEET_NAMES.TICKERS);
  const sampleTickers = tickers.slice(0, CC_CONFIG.MAX_DISCOVERY_FILES);
  
  for (const ticker of sampleTickers) {
    try {
      const json = fetchFromAPI(ticker);
      accumulateSchema(json, sections, fieldsBySection);
    } catch (e) {
      logMessage(`Skipping ${ticker}: API error`);
    }
  }
  
  logMessage(`Discovered from ${sampleTickers.length} tickers.`);
}

/**
 * Accumulates schema from a single JSON object.
 * @param {any} json - EODHD fundamentals JSON
 * @param {Set<string>} sections - Set to populate
 * @param {Object} fieldsBySection - Object to populate
 */
function accumulateSchema(
  json: any,
  sections: Set<string>,
  fieldsBySection: { [section: string]: Set<string> }
): void {
  const financials = json?.Financials;
  if (!financials) return;
  
  const sectionNames = [
    FINANCIAL_SECTIONS.INCOME_STATEMENT,
    FINANCIAL_SECTIONS.BALANCE_SHEET,
    FINANCIAL_SECTIONS.CASH_FLOW
  ];
  
  for (const section of sectionNames) {
    const sectionData = financials[section];
    if (!sectionData) continue;
    
    sections.add(section);
    
    if (!fieldsBySection[section]) {
      fieldsBySection[section] = new Set();
    }
    
    // Check both yearly and quarterly arrays
    for (const period of ['yearly', 'quarterly']) {
      const arr = sectionData[period];
      
      if (Array.isArray(arr)) {
        for (const obj of arr) {
          if (obj && typeof obj === 'object') {
            for (const key of Object.keys(obj)) {
              // Skip metadata fields
              if (key === 'date' || key === 'fiscalYear' || key === 'fiscalQuarter') {
                continue;
              }
              fieldsBySection[section].add(key);
            }
          }
        }
      }
    }
  }
}

/**
 * Ensures lookup sheet exists and is clean.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Lookup sheet
 */
function ensureLookupSheet(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet {
  let sheet = ss.getSheetByName(CC_CONFIG.LOOKUP_SHEET);
  
  if (!sheet) {
    sheet = ss.insertSheet(CC_CONFIG.LOOKUP_SHEET);
  }
  
  sheet.clear();
  
  // Remove old named ranges
  const rangeNames = ['SectionsRange', 'PeriodTypeList', 'ActiveList'];
  const fieldRangePrefix = 'Fields_';
  
  for (const nr of ss.getNamedRanges()) {
    const name = nr.getName();
    if (rangeNames.includes(name) || name.startsWith(fieldRangePrefix)) {
      nr.remove();
    }
  }
  
  return sheet;
}

/**
 * Writes named lookups to lookup sheet and creates named ranges.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Lookup sheet
 * @param {SchemaDiscovery} schema - Discovered schema
 */
function writeNamedLookups(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  schema: SchemaDiscovery
): void {
  // Column 1: Sections
  sheet.getRange(1, 1).setValue('Sections');
  if (schema.sections.length > 0) {
    const sectionValues = schema.sections.map(s => [s]);
    sheet.getRange(2, 1, sectionValues.length, 1).setValues(sectionValues);
    setNamedRange('SectionsRange', sheet.getRange(2, 1, sectionValues.length, 1));
  } else {
    setNamedRange('SectionsRange', sheet.getRange(2, 1, 1, 1));
  }
  
  // Column 2: Period Type
  sheet.getRange(1, 2).setValue('PeriodType');
  const periodValues = CC_CONFIG.PERIOD_VALUES.map(v => [v]);
  sheet.getRange(2, 2, periodValues.length, 1).setValues(periodValues);
  setNamedRange('PeriodTypeList', sheet.getRange(2, 2, periodValues.length, 1));
  
  // Column 3: Active
  sheet.getRange(1, 3).setValue('Active');
  const activeValues = CC_CONFIG.ACTIVE_VALUES.map(v => [v]);
  sheet.getRange(2, 3, activeValues.length, 1).setValues(activeValues);
  setNamedRange('ActiveList', sheet.getRange(2, 3, activeValues.length, 1));
  
  // Per-section field lists (starting at column 5, spaced by 2)
  let col = 5;
  for (const section of schema.sections) {
    const fields = Array.from(schema.fieldsBySection[section] || new Set()).sort();
    
    sheet.getRange(1, col).setValue(section);
    
    if (fields.length > 0) {
      const fieldValues = fields.map(f => [f]);
      sheet.getRange(2, col, fieldValues.length, 1).setValues(fieldValues);
      setNamedRange(`Fields_${section}`, sheet.getRange(2, col, fieldValues.length, 1));
    } else {
      setNamedRange(`Fields_${section}`, sheet.getRange(2, col, 1, 1));
    }
    
    col += 2; // Space columns by 2
  }
  
  // Hide lookup sheet
  hideSheet(sheet);
}

/**
 * Applies validation rules to Control Center sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Control Center sheet
 */
function applyValidationRules(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): void {
  const HEADER_ROW = CONTROL_CENTER_CONFIG.HEADER_ROW;
  const START_ROW = HEADER_ROW + 1;
  const lastRow = Math.max(sheet.getLastRow(), START_ROW + 50);
  
  const COL_SECTION = CONTROL_CENTER_CONFIG.COLUMNS.SECTION;
  const COL_FIELD = CONTROL_CENTER_CONFIG.COLUMNS.FIELD_NAME;
  const COL_PERIOD = CONTROL_CENTER_CONFIG.COLUMNS.PERIOD_TYPE;
  const COL_ACTIVE = CONTROL_CENTER_CONFIG.COLUMNS.ACTIVE;
  
  // Clear existing validations
  sheet.getRange(START_ROW, COL_SECTION, lastRow - START_ROW + 1, 4)
    .clearDataValidations();
  
  // Section dropdown (Column A)
  const sectionsRange = ss.getRangeByName('SectionsRange');
  if (sectionsRange) {
    const sectionRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sectionsRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(START_ROW, COL_SECTION, lastRow - START_ROW + 1, 1)
      .setDataValidation(sectionRule);
  }
  
  // Period Type dropdown (Column C)
  const periodRange = ss.getRangeByName('PeriodTypeList');
  if (periodRange) {
    const periodRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(periodRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(START_ROW, COL_PERIOD, lastRow - START_ROW + 1, 1)
      .setDataValidation(periodRule);
  }
  
  // Active dropdown (Column D)
  const activeRange = ss.getRangeByName('ActiveList');
  if (activeRange) {
    const activeRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(activeRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(START_ROW, COL_ACTIVE, lastRow - START_ROW + 1, 1)
      .setDataValidation(activeRule);
  }
  
  // Field Name dropdown (Column B) - depends on Section per row
  for (let row = START_ROW; row <= lastRow; row++) {
    const section = String(sheet.getRange(row, COL_SECTION).getValue() || '').trim();
    applyFieldValidationForRow(ss, sheet, row, section, COL_FIELD);
  }
}

/**
 * Applies field validation for a single row based on selected section.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet
 * @param {number} row - Row number
 * @param {string} section - Selected section
 * @param {number} col - Field column number
 */
function applyFieldValidationForRow(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  section: string,
  col: number
): void {
  const fieldCell = sheet.getRange(row, col);
  fieldCell.clearDataValidations();
  
  if (!section) return; // No section selected
  
  const fieldsRange = ss.getRangeByName(`Fields_${section}`);
  
  if (fieldsRange) {
    const fieldRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(fieldsRange, true)
      .setAllowInvalid(true)
      .build();
    fieldCell.setDataValidation(fieldRule);
  }
}

/**
 * Fetches fundamentals from EODHD API.
 * @param {string} ticker - Ticker symbol
 * @returns {any} Parsed JSON
 */
function fetchFromAPI(ticker: string): any {
  const apiKey = getEODHDApiKey();
  const normalizedTicker = normalizeTickerSymbol(ticker);
  
  const url = `${API_CONFIG.EODHD_BASE_URL}/${API_CONFIG.EODHD_FUNDAMENTALS_ENDPOINT}/${encodeURIComponent(normalizedTicker)}?api_token=${encodeURIComponent(apiKey)}&fmt=json`;
  
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`API error ${response.getResponseCode()}`);
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * Sanitizes Drive folder ID.
 * @param {string} input - Raw input
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
  console.log(CC_CONFIG.LOG_PREFIX + message);
}

