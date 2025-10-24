/**
 * utils/upsert.ts — Generic upsert utilities for sheet data
 * ===========================================================
 * Purpose:
 *   - Upsert (update or insert) rows by unique key
 *   - Preserve existing columns not in update set
 *   - Support for "long" table format with composite keys
 *   - Batch operations for performance
 * 
 * Dependencies: utils/sheets.ts, utils/format.ts
 * Called by: Extractor, per_share, calc_metrics modules
 */

import { buildHeaderIndex, type HeaderIndex } from './sheets';
import { nullToEmpty } from './format';

/**
 * Upserts rows into a sheet by a single unique key column.
 * Updates existing rows or appends new ones.
 * Preserves columns not in the columnsToWrite list.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} keyColumn - Unique key column name
 * @param {string[]} columnsToWrite - Columns to update/write
 * @param {any[]} incomingObjects - Row objects to upsert
 * @param {boolean} appendMissingColumns - If true, adds new columns to header
 */
export function upsertByKey(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyColumn: string,
  columnsToWrite: string[],
  incomingObjects: any[],
  appendMissingColumns: boolean = true
): void {
  if (!incomingObjects || incomingObjects.length === 0) {
    return;
  }
  
  // Ensure key column is in write list
  if (!columnsToWrite.includes(keyColumn)) {
    columnsToWrite = [keyColumn, ...columnsToWrite];
  }
  
  // Read existing data
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(1, sheet.getLastColumn());
  
  let header: string[] = [];
  let data: any[][] = [];
  
  if (lastRow > 0) {
    const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    header = allValues[0].map(v => String(v || '').trim());
    data = allValues;
  }
  
  // Build or update header
  const headerIndex = buildHeaderIndex(header);
  
  // Ensure key column exists
  if (!(keyColumn in headerIndex)) {
    header.push(keyColumn);
    headerIndex[keyColumn] = header.length - 1;
  }
  
  // Add missing columns if requested
  if (appendMissingColumns) {
    for (const col of columnsToWrite) {
      if (!(col in headerIndex)) {
        header.push(col);
        headerIndex[col] = header.length - 1;
      }
    }
  }
  
  // Extend existing rows to match new header length
  for (let i = 0; i < data.length; i++) {
    while (data[i].length < header.length) {
      data[i].push('');
    }
  }
  
  // Build key-to-row map
  const keyColIdx = headerIndex[keyColumn];
  const rowByKey = new Map<string, number>();
  
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][keyColIdx] || '').trim();
    if (key) {
      rowByKey.set(key, i);
    }
  }
  
  // Process incoming objects
  for (const obj of incomingObjects) {
    const key = String(obj[keyColumn] || '').trim();
    if (!key) continue;
    
    let rowIdx = rowByKey.get(key);
    
    if (rowIdx === undefined) {
      // New row - append
      rowIdx = data.length;
      const newRow = new Array(header.length).fill('');
      newRow[keyColIdx] = key;
      data.push(newRow);
      rowByKey.set(key, rowIdx);
    }
    
    // Update columns
    for (const col of columnsToWrite) {
      const colIdx = headerIndex[col];
      if (colIdx !== undefined && Object.prototype.hasOwnProperty.call(obj, col)) {
        data[rowIdx][colIdx] = nullToEmpty(obj[col]);
      }
    }
  }
  
  // Write back to sheet
  sheet.clearContents();
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, header.length).setValues(data);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  
  // Freeze header row
  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
}

/**
 * Upserts rows into a "long" table format with composite keys.
 * Used for Raw_Fundamentals tables where key = (Ticker, Date, Section, Field).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {Map<string, number>} existingKeyMap - Map of composite key to row number
 * @param {string[]} headers - Header row
 * @param {any[][]} newRows - New rows to upsert
 * @param {string[]} keyColumns - Columns that form the composite key
 */
export function upsertLongTable(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  existingKeyMap: Map<string, number>,
  headers: string[],
  newRows: any[][],
  keyColumns: string[]
): void {
  if (newRows.length === 0) return;
  
  const headerIndex = buildHeaderIndex(headers);
  const keyIndices = keyColumns.map(col => headerIndex[col]);
  
  // Separate updates vs. appends
  const updates: { [rowNum: number]: any[] } = {};
  const appends: any[][] = [];
  
  for (const row of newRows) {
    // Build composite key
    const keyParts = keyIndices.map(idx => String(row[idx] || ''));
    const compositeKey = keyParts.join('¦');
    
    const existingRow = existingKeyMap.get(compositeKey);
    
    if (existingRow !== undefined) {
      // Update existing row
      updates[existingRow] = row;
    } else {
      // Append new row
      appends.push(row);
      
      // Update map for future upserts in this batch
      const nextRow = sheet.getLastRow() + appends.length;
      existingKeyMap.set(compositeKey, nextRow);
    }
  }
  
  // Execute updates
  for (const [rowNumStr, rowValues] of Object.entries(updates)) {
    const rowNum = parseInt(rowNumStr, 10);
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
  }
  
  // Execute appends
  if (appends.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, appends.length, headers.length)
      .setValues(appends);
  }
}

/**
 * Reads a sheet and builds a map of composite key to row number.
 * Used for preparing long table upserts.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to read
 * @param {string[]} keyColumns - Columns that form the composite key
 * @returns {Map<string, number>} Map of composite key to row number
 */
export function buildKeyMap(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyColumns: string[]
): Map<string, number> {
  const keyMap = new Map<string, number>();
  
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(1, sheet.getLastColumn());
  
  if (lastRow < 2) {
    return keyMap;
  }
  
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v || '').trim());
  const headerIndex = buildHeaderIndex(header);
  
  // Check if all key columns exist
  const keyIndices: number[] = [];
  for (const col of keyColumns) {
    const idx = headerIndex[col];
    if (idx === undefined) {
      // Key column missing - return empty map (sheet is new)
      return keyMap;
    }
    keyIndices.push(idx);
  }
  
  // Build map
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const keyParts = keyIndices.map(idx => String(row[idx] || ''));
    const compositeKey = keyParts.join('¦');
    
    if (keyParts.some(p => p.trim())) {
      // At least one key part is non-empty
      keyMap.set(compositeKey, i + 1); // Row number (1-indexed)
    }
  }
  
  return keyMap;
}

/**
 * Upserts data into Calculated_Metrics sheet with specific formatting.
 * Convenience wrapper for the most common upsert operation.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Calculated_Metrics sheet
 * @param {Map<string, any>} dataByTicker - Map of ticker to metrics object
 */
export function upsertCalculatedMetrics(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dataByTicker: Map<string, any>
): void {
  const ownedColumns = [
    'Ticker',
    'OpPS_Latest', 'OpPS_5Y_CAGR', 'OpPS_9Y_CAGR', 'OpPS_10Y_CAGR', 'OpPS_AdjustedFlag',
    'EPS_Latest', 'EPS_5Y_CAGR', 'EPS_9Y_CAGR', 'EPS_10Y_CAGR', 'EPS_AdjustedFlag',
    'FCFPS_Latest', 'FCFPS_5Y_CAGR', 'FCFPS_9Y_CAGR', 'FCFPS_10Y_CAGR', 'FCFPS_AdjustedFlag',
    'Calc_Timestamp'
  ];
  
  const incomingObjects = Array.from(dataByTicker.values());
  
  upsertByKey(sheet, 'Ticker', ownedColumns, incomingObjects, true);
  
  // Apply formatting
  applyCalculatedMetricsFormatting(sheet);
}

/**
 * Applies number formatting to Calculated_Metrics sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to format
 */
function applyCalculatedMetricsFormatting(
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): void {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 1) return;
  
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerIndex = buildHeaderIndex(header);
  
  // Format decimal columns (Latest values)
  const decimalCols = ['OpPS_Latest', 'EPS_Latest', 'FCFPS_Latest'];
  for (const col of decimalCols) {
    const idx = headerIndex[col];
    if (idx !== undefined) {
      sheet.getRange(2, idx + 1, lastRow - 1, 1).setNumberFormat('0.00');
    }
  }
  
  // Format percentage columns (CAGRs)
  const percentCols = [
    'OpPS_5Y_CAGR', 'OpPS_9Y_CAGR', 'OpPS_10Y_CAGR',
    'EPS_5Y_CAGR', 'EPS_9Y_CAGR', 'EPS_10Y_CAGR',
    'FCFPS_5Y_CAGR', 'FCFPS_9Y_CAGR', 'FCFPS_10Y_CAGR'
  ];
  for (const col of percentCols) {
    const idx = headerIndex[col];
    if (idx !== undefined) {
      sheet.getRange(2, idx + 1, lastRow - 1, 1).setNumberFormat('0.00%');
    }
  }
  
  // Format timestamp
  const tsIdx = headerIndex['Calc_Timestamp'];
  if (tsIdx !== undefined) {
    sheet.getRange(2, tsIdx + 1, lastRow - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

/**
 * Ensures required columns exist in a sheet and returns header info.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to check
 * @param {string[]} requiredColumns - Required column names
 * @returns {Object} Header array and index map
 */
export function ensureColumns(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  requiredColumns: string[]
): { header: string[]; index: HeaderIndex } {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  let header: string[] = [];
  
  if (lastRow > 0 && lastCol > 0) {
    header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(v => String(v || '').trim());
  }
  
  const index = buildHeaderIndex(header);
  let mutated = false;
  
  // Add missing columns
  for (const col of requiredColumns) {
    if (!(col in index)) {
      header.push(col);
      index[col] = header.length - 1;
      mutated = true;
    }
  }
  
  // Write header if mutated or empty
  if (mutated || header.length === 0) {
    const finalHeader = header.length > 0 ? header : requiredColumns.slice();
    sheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);
    sheet.getRange(1, 1, 1, finalHeader.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return { header: finalHeader, index: buildHeaderIndex(finalHeader) };
  }
  
  return { header, index };
}

