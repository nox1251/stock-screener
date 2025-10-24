/***************************************************************
 * UpsertLogic.gs — generic, batched upsert utilities
 * - NO global sheet-name constants here (avoids collisions)
 * - Works with row objects; preserves other columns
 ***************************************************************/

function UL_getSheet_(name, create) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function UL_indexHeader_(row1) {
  const idx = {};
  row1.forEach((h, i) => { idx[String(h).trim()] = i; });
  return idx;
}

function UL_ensureColumns_(sh, requiredCols) {
  const range = sh.getLastRow() ? sh.getRange(1,1,1,sh.getLastColumn()) : null;
  const header = range ? range.getValues()[0] : [];
  const idx = UL_indexHeader_(header);

  let mutated = false;
  for (const col of requiredCols) {
    if (idx[col] == null) {
      header.push(col);
      idx[col] = header.length - 1;
      mutated = true;
    }
  }
  if (mutated || header.length === 0) {
    sh.getRange(1,1,1,Math.max(1, header.length)).setValues([header.length ? header : requiredCols.slice()]);
  }
  return { header: header.length ? header : requiredCols.slice(), idx: UL_indexHeader_(header.length ? header : requiredCols) };
}

function UL_readAsObjects_(sh, keyCol) {
  const values = sh.getDataRange().getValues();
  if (values.length === 0) return { header: [], rows: [] };
  const header = values[0].map(h => String(h).trim());
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    for (let c = 0; c < header.length; c++) obj[header[c]] = values[r][c];
    if (keyCol && !obj[keyCol]) continue;
    rows.push(obj);
  }
  return { header, rows };
}

/**
 * Batched upsert by key.
 * @param {Sheet} sh
 * @param {string} keyCol                unique key column
 * @param {string[]} columnsToWrite      columns to write (must include keyCol)
 * @param {Object[]} incoming            row objects including keyCol
 * @param {boolean} appendMissingColumns append unknown columns to header
 */
function UL_upsertByKey_(sh, keyCol, columnsToWrite, incoming, appendMissingColumns) {
  if (!incoming || !incoming.length) return;

  const ensureCols = appendMissingColumns ? Array.from(new Set(columnsToWrite)) : [keyCol];
  const { header, idx } = UL_ensureColumns_(sh, ensureCols);

  const data = sh.getDataRange().getValues();
  const keyIndex = idx[keyCol];
  const rowIndexByKey = new Map();
  for (let r = 1; r < data.length; r++) {
    const key = data[r][keyIndex];
    if (key !== '' && key != null) rowIndexByKey.set(String(key), r);
  }

  let headerMutated = false;
  for (const col of columnsToWrite) {
    if (idx[col] == null) {
      header.push(col);
      idx[col] = header.length - 1;
      headerMutated = true;
    }
  }
  if (headerMutated) sh.getRange(1,1,1,header.length).setValues([header]);

  let grid = data;
  for (const obj of incoming) {
    const key = obj[keyCol];
    if (key == null || key === '') continue;

    let r = rowIndexByKey.get(String(key));
    if (r == null) {
      r = grid.length;
      const empty = new Array(header.length).fill('');
      empty[idx[keyCol]] = key;
      grid.push(empty);
      rowIndexByKey.set(String(key), r);
    }

    for (const col of columnsToWrite) {
      const c = idx[col];
      if (c == null) continue;
      if (Object.prototype.hasOwnProperty.call(obj, col)) {
        grid[r][c] = obj[col];
      }
    }
  }

  sh.clearContents();
  sh.getRange(1,1,grid.length,header.length).setValues(grid);
}

/* ---------- Project-specific wrapper (no globals declared here) ---------- */

/** Upsert into Calculated_Metrics (expects columns below). */
function UL_upsertCalculatedMetrics_(rows) {
  const sh = UL_getSheet_('Calculated_Metrics'); // literal here to avoid SH_* collisions
  const cols = ['Ticker','OpPS_Latest','OpPS_5Y_CAGR','OpPS_9Y_CAGR','Debt_to_Equity','Calc_Timestamp'];
  UL_upsertByKey_(sh, 'Ticker', cols, rows, true);

  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx = UL_indexHeader_(hdr);
  sh.setFrozenRows(1);
  if (idx['OpPS_Latest'] != null)    sh.getRange(2, idx['OpPS_Latest']+1, sh.getLastRow()-1, 1).setNumberFormat('0.00');
  if (idx['OpPS_5Y_CAGR'] != null)   sh.getRange(2, idx['OpPS_5Y_CAGR']+1, sh.getLastRow()-1, 1).setNumberFormat('0.00%');
  if (idx['OpPS_9Y_CAGR'] != null)   sh.getRange(2, idx['OpPS_9Y_CAGR']+1, sh.getLastRow()-1, 1).setNumberFormat('0.00%');
  if (idx['Debt_to_Equity'] != null) sh.getRange(2, idx['Debt_to_Equity']+1, sh.getLastRow()-1, 1).setNumberFormat('0.00');
  if (idx['Calc_Timestamp'] != null) sh.getRange(2, idx['Calc_Timestamp']+1, sh.getLastRow()-1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}
