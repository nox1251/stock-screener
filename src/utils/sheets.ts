/**
 * utils/sheets.ts — Google Sheets reading and writing utilities
 * ===============================================================
 * Purpose:
 *   - Read data by header name (not column index)
 *   - Batch write operations
 *   - Sheet creation and header management
 *   - Column auto-sizing and formatting
 * 
 * Dependencies: utils/format.ts
 * Called by: All modules that interact with sheets
 */

import { nullToEmpty } from './format';

/**
 * Header-to-index mapping type
 */
export type HeaderIndex = { [header: string]: number };

/**
 * Table structure with headers and data rows
 */
export interface TableData {
  header: string[];
  rows: any[][];
}

/**
 * Table structure with rows as objects (keyed by header)
 */
export interface TableObjects<T = any> {
  header: string[];
  rows: T[];
}

/**
 * Gets or creates a sheet by name.
 * @param {string} name - Sheet name
 * @param {boolean} create - If true, creates sheet if missing
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 * @throws {Error} If sheet not found and create=false
 */
export function getSheet(name: string, create: boolean = false): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  
  if (!sh && create) {
    sh = ss.insertSheet(name);
  }
  
  if (!sh) {
    throw new Error(`Sheet not found: ${name}`);
  }
  
  return sh;
}

/**
 * Reads entire sheet as table with header row and data rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to read
 * @returns {TableData} Table with header and rows
 */
export function readTable(sheet: GoogleAppsScript.Spreadsheet.Sheet): TableData {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) {
    return { header: [], rows: [] };
  }
  
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  if (values.length === 0) {
    return { header: [], rows: [] };
  }
  
  const header = values[0].map(v => String(v || '').trim());
  const rows = values.slice(1);
  
  return { header, rows };
}

/**
 * Reads sheet data and converts rows to objects keyed by header names.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to read
 * @param {string | null} keyCol - Optional: filter out rows where this column is empty
 * @returns {TableObjects} Table with header and row objects
 */
export function readTableAsObjects<T = any>(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyCol: string | null = null
): TableObjects<T> {
  const table = readTable(sheet);
  const rows: T[] = [];
  
  for (const row of table.rows) {
    const obj: any = {};
    
    for (let c = 0; c < table.header.length; c++) {
      obj[table.header[c]] = row[c];
    }
    
    // Filter: skip if keyCol is specified and empty
    if (keyCol && !obj[keyCol]) {
      continue;
    }
    
    rows.push(obj as T);
  }
  
  return { header: table.header, rows };
}

/**
 * Creates a header-to-column-index mapping.
 * @param {string[]} headers - Array of header names
 * @returns {HeaderIndex} Mapping of header name to column index
 */
export function buildHeaderIndex(headers: string[]): HeaderIndex {
  const index: HeaderIndex = {};
  headers.forEach((h, i) => {
    index[String(h).trim()] = i;
  });
  return index;
}

/**
 * Resolves column indices for multiple column names with alias support.
 * Throws error if any required column is missing.
 * @param {string[]} headers - Sheet header row
 * @param {Object} aliasMap - Map of logical name to array of accepted aliases
 * @returns {HeaderIndex} Map of logical name to column index
 */
export function resolveColumns(
  headers: string[],
  aliasMap: { [key: string]: string[] }
): HeaderIndex {
  const lowerHeaders = headers.map(h => String(h).toLowerCase());
  const resolved: HeaderIndex = {};
  
  for (const [logicalName, aliases] of Object.entries(aliasMap)) {
    let foundIdx = -1;
    
    for (const alias of aliases) {
      const idx = lowerHeaders.indexOf(alias.toLowerCase());
      if (idx !== -1) {
        foundIdx = idx;
        break;
      }
    }
    
    if (foundIdx === -1) {
      throw new Error(
        `Required column "${logicalName}" not found. ` +
        `Looked for: ${aliases.join(', ')}`
      );
    }
    
    resolved[logicalName] = foundIdx;
  }
  
  return resolved;
}

/**
 * Ensures a sheet exists with specified headers.
 * Creates sheet if missing, updates headers if different.
 * @param {string} sheetName - Name of sheet
 * @param {string[]} headers - Expected header row
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 */
export function ensureSheetWithHeaders(
  sheetName: string,
  headers: string[]
): GoogleAppsScript.Spreadsheet.Sheet {
  const sh = getSheet(sheetName, true);
  ensureHeaders(sh, headers);
  return sh;
}

/**
 * Ensures a sheet has the correct headers.
 * Clears and rewrites if headers don't match.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to update
 * @param {string[]} headers - Expected headers
 */
export function ensureHeaders(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): void {
  const neededCols = headers.length;
  
  // Ensure enough columns exist
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      neededCols - sheet.getMaxColumns()
    );
  }
  
  // Check if headers match
  const lastCol = Math.max(1, sheet.getLastColumn());
  const existing = lastCol >= neededCols
    ? sheet.getRange(1, 1, 1, neededCols).getValues()[0]
    : [];
  
  const mismatch = existing.length !== headers.length ||
    headers.some((h, i) => existing[i] !== h);
  
  if (mismatch) {
    // Headers don't match — clear and rewrite
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  // Always ensure frozen header row
  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
}

/**
 * Writes a table (header + rows) to a sheet.
 * Clears the sheet first.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to write to
 * @param {string[]} headers - Header row
 * @param {any[][]} rows - Data rows
 * @param {boolean} freezeHeader - If true, freezes first row
 */
export function writeTable(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[],
  rows: any[][],
  freezeHeader: boolean = true
): void {
  sheet.clear();
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // Write data
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Freeze header
  if (freezeHeader) {
    sheet.setFrozenRows(1);
  }
}

/**
 * Appends rows to a sheet (assumes headers already exist).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to append to
 * @param {any[][]} rows - Rows to append
 */
export function appendRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: any[][]
): void {
  if (rows.length === 0) return;
  
  const lastRow = sheet.getLastRow();
  const numCols = sheet.getLastColumn();
  
  sheet.getRange(lastRow + 1, 1, rows.length, numCols).setValues(rows);
}

/**
 * Updates specific rows in place by row number.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to update
 * @param {Object} updates - Map of rowNumber -> rowValues
 * @param {number} numCols - Number of columns per row
 */
export function updateRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  updates: { [rowNum: number]: any[] },
  numCols: number
): void {
  for (const [rowNumStr, rowValues] of Object.entries(updates)) {
    const rowNum = parseInt(rowNumStr, 10);
    sheet.getRange(rowNum, 1, 1, numCols).setValues([rowValues]);
  }
}

/**
 * Auto-resizes columns to fit content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to resize
 * @param {number} maxColumns - Maximum number of columns to resize (default: 20)
 */
export function autoResizeColumns(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  maxColumns: number = 20
): void {
  try {
    const lastCol = Math.min(maxColumns, sheet.getLastColumn());
    if (lastCol > 0) {
      sheet.autoResizeColumns(1, lastCol);
    }
  } catch (e) {
    Logger.log('Auto-resize failed: ' + e);
  }
}

/**
 * Applies number format to a range of columns.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to format
 * @param {number} startRow - Starting row (1-indexed)
 * @param {number[]} columns - Column indices (0-indexed) to format
 * @param {string} format - Number format string
 */
export function formatColumns(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number,
  columns: number[],
  format: string
): void {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  const numRows = lastRow - startRow + 1;
  
  for (const col of columns) {
    try {
      sheet.getRange(startRow, col + 1, numRows, 1).setNumberFormat(format);
    } catch (e) {
      Logger.log(`Format column ${col} failed: ` + e);
    }
  }
}

/**
 * Reads a specific named range value.
 * @param {string} rangeName - Named range name
 * @returns {any} Cell value or null if range not found
 */
export function readNamedRange(rangeName: string): any {
  try {
    const ss = SpreadsheetApp.getActive();
    const range = ss.getRangeByName(rangeName);
    if (!range) return null;
    return range.getValue();
  } catch (e) {
    Logger.log(`Read named range ${rangeName} failed: ` + e);
    return null;
  }
}

/**
 * Batch-converts row objects to row arrays based on header order.
 * @param {any[]} objects - Array of row objects
 * @param {string[]} headers - Ordered header names
 * @returns {any[][]} Array of row arrays
 */
export function objectsToRows(objects: any[], headers: string[]): any[][] {
  return objects.map(obj => {
    return headers.map(h => nullToEmpty(obj[h]));
  });
}

/**
 * Gets all ticker symbols from a configured tickers source.
 * Tries multiple strategies: configured sheet, auto-detect, or selection.
 * @param {string} preferredSheet - Preferred sheet name (e.g., 'Tickers')
 * @param {string} preferredRange - Preferred range (e.g., 'A2:A')
 * @returns {string[]} Array of ticker symbols
 * @throws {Error} If no tickers found
 */
export function getTickers(
  preferredSheet: string = 'Tickers',
  preferredRange: string = 'A2:A'
): string[] {
  const ss = SpreadsheetApp.getActive();
  
  // Strategy 1: Configured sheet
  const configuredSheet = ss.getSheetByName(preferredSheet);
  if (configuredSheet) {
    const vals = configuredSheet.getRange(preferredRange)
      .getValues()
      .map(r => String(r[0] || '').trim())
      .filter(Boolean);
    
    if (vals.length > 0) return vals;
  }
  
  // Strategy 2: Auto-detect sheet with "Ticker" column
  for (const sh of ss.getSheets()) {
    const lastCol = Math.max(1, sh.getLastColumn());
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const lowerHeader = header.map(v => String(v || '').toLowerCase());
    const colIdx = lowerHeader.indexOf('ticker');
    
    if (colIdx >= 0) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const vals = sh.getRange(2, colIdx + 1, lastRow - 1, 1)
          .getValues()
          .map(r => String(r[0] || '').trim())
          .filter(Boolean);
        
        if (vals.length > 0) return vals;
      }
    }
  }
  
  // Strategy 3: Use current selection (first column)
  const selection = ss.getActiveRange();
  if (selection) {
    const vals = selection.getValues()
      .map(r => String(r[0] || '').trim())
      .filter(Boolean);
    
    if (vals.length > 0) return vals;
  }
  
  throw new Error(
    'Tickers not found.\n' +
    `Create a sheet named "${preferredSheet}" with A1=Ticker and ` +
    'ticker symbols in column A starting at A2,\n' +
    'or select a single-column range before running.'
  );
}

/**
 * Hides a sheet from view.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to hide
 */
export function hideSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  try {
    sheet.hideSheet();
  } catch (e) {
    Logger.log('Hide sheet failed: ' + e);
  }
}

/**
 * Sets or updates a named range.
 * @param {string} name - Named range name
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Range to name
 */
export function setNamedRange(
  name: string,
  range: GoogleAppsScript.Spreadsheet.Range
): void {
  const ss = SpreadsheetApp.getActive();
  
  // Remove existing named range if present
  const existing = ss.getRangeByName(name);
  if (existing) {
    ss.getNamedRanges()
      .filter(nr => nr.getName() === name)
      .forEach(nr => nr.remove());
  }
  
  // Set new named range
  ss.setNamedRange(name, range);
}

