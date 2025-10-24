/**
 * code.ts — Stock Screener main menu and entrypoints
 * ====================================================
 * Purpose:
 *   - Register onOpen menu
 *   - Wire menu items to module functions
 *   - Handle user prompts for settings
 *   - Provide configuration check utilities
 * 
 * Dependencies: All modules (extractor, per_share, calc_metrics, control_center, utils)
 * Called by: Google Apps Script runtime (onOpen trigger)
 */

import { withUiSpinner, promptUser, showInfo } from './utils/spinner';
import { setEODHDApiKey, hasEODHDApiKey, getMockFolderId } from './globals';
import { extractAnnualFundamentals, extractQuarterlyFundamentals, extractAllFundamentals } from './extractor/fundamentals';
import { extractSplitsDividends } from './extractor/splits_dividends';
import { buildPerShareFull } from './per_share/builder';
import { computeCAGRs } from './calc_metrics/cagr';
import { applyControlCenterValidations } from './control_center/data_validation';

/**
 * Creates custom menu on spreadsheet open.
 * This function is automatically called by Google Apps Script.
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Stock Screener')
    .addSubMenu(
      ui.createMenu('Extract Fundamentals')
        .addItem('Extract — Annual (All Tickers)', 'menu_ExtractAnnual')
        .addItem('Extract — Quarterly (All Tickers)', 'menu_ExtractQuarterly')
        .addItem('Extract — All (Annual + Quarterly)', 'menu_ExtractAll')
        .addItem('Extract — Splits & Dividends', 'menu_ExtractSplitsDividends')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Per Share')
        .addItem('Build Per Share (Full)', 'menu_BuildPerShareFull')
        .addItem('Compute CAGRs', 'menu_ComputeCAGRs')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Control Center')
        .addItem('Refresh Dropdowns', 'menu_RefreshControlCenter')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Settings')
        .addItem('Set API Key (EODHD)', 'menu_SetApiKey')
        .addItem('Check Config', 'menu_CheckConfig')
    )
    .addToUi();
}

/* ===================== Extract Fundamentals ===================== */

/**
 * Menu: Extract Annual Fundamentals
 */
function menu_ExtractAnnual(): void {
  withUiSpinner(
    () => extractAnnualFundamentals(),
    'Extracting Annual Fundamentals...'
  );
}

/**
 * Menu: Extract Quarterly Fundamentals
 */
function menu_ExtractQuarterly(): void {
  withUiSpinner(
    () => extractQuarterlyFundamentals(),
    'Extracting Quarterly Fundamentals...'
  );
}

/**
 * Menu: Extract All (Annual + Quarterly)
 */
function menu_ExtractAll(): void {
  withUiSpinner(
    () => extractAllFundamentals(),
    'Extracting Annual + Quarterly...'
  );
}

/**
 * Menu: Extract Splits & Dividends
 */
function menu_ExtractSplitsDividends(): void {
  withUiSpinner(
    () => extractSplitsDividends(),
    'Extracting Splits & Dividends...'
  );
}

/* ===================== Per Share ===================== */

/**
 * Menu: Build Per Share (Full)
 */
function menu_BuildPerShareFull(): void {
  withUiSpinner(
    () => buildPerShareFull(),
    'Building Per Share (Full)...'
  );
}

/**
 * Menu: Compute CAGRs
 */
function menu_ComputeCAGRs(): void {
  withUiSpinner(
    () => computeCAGRs(),
    'Computing 5Y/9Y/10Y CAGRs...'
  );
}

/* ===================== Control Center ===================== */

/**
 * Menu: Refresh Control Center Dropdowns
 */
function menu_RefreshControlCenter(): void {
  withUiSpinner(
    () => applyControlCenterValidations(),
    'Refreshing Control Center Dropdowns...'
  );
}

/* ===================== Settings ===================== */

/**
 * Menu: Set API Key
 */
function menu_SetApiKey(): void {
  const key = promptUser(
    'EODHD API Key',
    'Enter your EODHD api_token:'
  );
  
  if (!key) {
    showInfo('No key provided.');
    return;
  }
  
  try {
    setEODHDApiKey(key);
    showInfo('API key saved successfully.');
  } catch (e: any) {
    showInfo(`Error saving API key:\n${e.message}`);
  }
}

/**
 * Menu: Check Configuration
 */
function menu_CheckConfig(): void {
  const ss = SpreadsheetApp.getActive();
  const apiKeyPresent = hasEODHDApiKey() ? 'YES' : 'NO';
  const mockFolderId = getMockFolderId();
  const mockMode = mockFolderId.length > 0 ? 'ON' : 'OFF';
  
  // Check critical sheets
  const criticalSheets = [
    'Tickers',
    'Control Center',
    'Per_Share',
    'Calculated_Metrics',
    'Raw_Fundamentals_Annual',
    'Raw_Fundamentals_Quarterly',
    'Raw_Splits_Dividends'
  ];
  
  const lines = [
    '=== Configuration Status ===',
    '',
    `API Key present: ${apiKeyPresent}`,
    `Mock Mode: ${mockMode}`,
    `Mock Folder ID: ${mockFolderId || '(not set)'}`,
    '',
    '=== Sheet Status ===',
    ''
  ];
  
  for (const name of criticalSheets) {
    const exists = ss.getSheetByName(name) ? 'OK' : 'MISSING';
    lines.push(`• ${name}: ${exists}`);
  }
  
  showInfo(lines.join('\n'));
}

/* ===================== onEdit Trigger Handler ===================== */

/**
 * Handles edits to Control Center for dynamic validation.
 * This function is automatically called by Google Apps Script when a cell is edited.
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  try {
    if (!e || !e.range || !e.source) return;
    
    const sheet = e.range.getSheet();
    
    if (sheet.getName() !== 'Control Center') return;
    
    const HEADER_ROW = 9;
    const START_ROW = HEADER_ROW + 1;
    const COL_SECTION = 1; // Column A
    const COL_FIELD = 2;   // Column B
    
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    
    // Only handle edits to Section column (A) in data rows
    if (editedRow < START_ROW || editedCol !== COL_SECTION) {
      return;
    }
    
    // Clear Field Name and reapply validation
    const ss = e.source;
    const section = String(e.range.getValue() || '').trim();
    const fieldCell = sheet.getRange(editedRow, COL_FIELD);
    
    fieldCell.setValue(''); // Clear field
    fieldCell.clearDataValidations();
    
    // Reapply validation if section is selected
    if (section) {
      const fieldsRange = ss.getRangeByName(`Fields_${section}`);
      if (fieldsRange) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(fieldsRange, true)
          .setAllowInvalid(true)
          .build();
        fieldCell.setDataValidation(rule);
      }
    }
    
  } catch (err: any) {
    Logger.log('onEdit error: ' + err.message);
  }
}

/* ===================== Utility Functions ===================== */

/**
 * One-time authorization helper for Drive access.
 * Run this manually from script editor if Drive access is needed.
 */
function authorizeDrive(): void {
  try {
    const folderId = getMockFolderId();
    if (!folderId) {
      throw new Error('Set Mock Folder ID in Control Center!B7 first.');
    }
    
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('Drive access authorized. Folder: ' + folder.getName());
    showInfo('Drive access is authorized.');
    
  } catch (e: any) {
    showInfo('Authorization failed:\n' + e.message);
  }
}

