/***************************************************************
 * Stock Screener — Extractor (Control-Center–driven)
 * -------------------------------------------------------------
 * - Honors Control Center (row 9 headers):
 *     A: Section | B: Field Name | C: Period Type | D: Active? | E: Notes
 * - Extracts ONLY fields marked Active?=Yes for the period
 * - Mock Mode (Drive JSONs; Control Center!B7) or API Mode (EODHD)
 * - Outputs long tables:
 *     Raw_Fundamentals_Annual / Raw_Fundamentals_Quarterly
 *     Raw_Splits_Dividends (if sheet exists)
 ***************************************************************/

/** ===================== CONFIG ===================== **/
const EXTRACTOR_CFG = {
  // ---- Source toggle ----
  MOCK_MODE: true, // true=Drive JSON; false=EODHD API

  // ---- Drive (folder ID comes from Control Center!B7) ----
  get MOCK_FOLDER_ID() { return _getMockFolderIdFromControlCenter_(); },

  // ---- API (used if MOCK_MODE=false) ----
  EODHD_API_KEY_PROP: 'EODHD_API_KEY', // set in Script Properties
  EODHD_FUNDAMENTALS_URL: 'https://eodhd.com/api/fundamentals',

  // ---- Tickers source ----
  TICKERS_SHEET: 'Tickers',
  TICKERS_RANGE: 'A2:A',

  // ---- Targets ----
  SHEET_ANNUAL: 'Raw_Fundamentals_Annual',
  SHEET_QUARTERLY: 'Raw_Fundamentals_Quarterly',
  SHEET_SPLITS_DIVS: 'Raw_Splits_Dividends', // optional

  // ---- Table headers ----
  ANNUAL_HEADERS:    ['Ticker','Date','FiscalYear','Section','Field','Value','Source'],
  QUARTERLY_HEADERS: ['Ticker','Date','FiscalYear','FiscalQuarter','Section','Field','Value','Source'],
  SPLITS_DIVS_HEADERS: [
    'Ticker','Kind','Date','DeclarationDate','RecordDate','PaymentDate',
    'Dividend','AdjDividend','ForFactor','ToFactor','Ratio','Notes'
  ],

  SOURCE_TAG: 'EODHD',
  BATCH_SIZE: 10,
  LOG_PREFIX: '[Extractor] '
};

/** ===================== PUBLIC ENTRY POINTS ===================== **/

function run_ExtractAll() {
  const tickers = _getTickers_();
  _log(`ExtractAll: ${tickers.length} tickers`);
  _extractAnnualForTickers_(tickers);
  _extractQuarterlyForTickers_(tickers);
  _extractSplitsDividendsForTickers_(tickers);
}

function run_ExtractAnnual() {
  const tickers = _getTickers_();
  _log(`ExtractAnnual: ${tickers.length} tickers`);
  _extractAnnualForTickers_(tickers);
}

function run_ExtractQuarterly() {
  const tickers = _getTickers_();
  _log(`ExtractQuarterly: ${tickers.length} tickers`);
  _extractQuarterlyForTickers_(tickers);
}

function run_ExtractSplitsDividends() {
  const tickers = _getTickers_();
  _log(`ExtractSplitsDividends: ${tickers.length} tickers`);
  _extractSplitsDividendsForTickers_(tickers);
}

/** ===================== CONTROL CENTER LOADER ===================== **/

// Reads Control Center selections and returns { annual: {Section:Set(fields)}, quarterly:{...} }
function _loadFieldConfigFromControlCenter_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Control Center');
  if (!sh) throw new Error('Control Center sheet not found.');

  const HEADER_ROW = 9;
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) {
    return {
      annual:   { Income_Statement:new Set(), Balance_Sheet:new Set(), Cash_Flow:new Set() },
      quarterly:{ Income_Statement:new Set(), Balance_Sheet:new Set(), Cash_Flow:new Set() }
    };
  }

  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const idx = {
    Section: headers.indexOf('Section') + 1,
    Field:   headers.indexOf('Field Name') + 1,
    Period:  headers.indexOf('Period Type') + 1,
    Active:  headers.indexOf('Active?') + 1
  };
  if (!idx.Section || !idx.Field || !idx.Period || !idx.Active) {
    throw new Error('Control Center headers must be: Section, Field Name, Period Type, Active?');
  }

  const cfg = {
    annual:   { Income_Statement:new Set(), Balance_Sheet:new Set(), Cash_Flow:new Set() },
    quarterly:{ Income_Statement:new Set(), Balance_Sheet:new Set(), Cash_Flow:new Set() }
  };

  const values = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  values.forEach(r => {
    const active = String(r[idx.Active - 1] || '').toLowerCase() === 'yes';
    if (!active) return;

    const section = String(r[idx.Section - 1] || '').trim();
    const field   = String(r[idx.Field - 1] || '').trim();
    const period  = String(r[idx.Period - 1] || '').trim();

    if (!section || !field || !period) return;
    if (!/^(Income_Statement|Balance_Sheet|Cash_Flow)$/.test(section)) return;

    if (/^annual$/i.test(period)) {
      cfg.annual[section].add(field);
    } else if (/^quarterly$/i.test(period)) {
      cfg.quarterly[section].add(field);
    }
  });

  return cfg;
}

/** ===================== CORE EXTRACTION ===================== **/

function _extractAnnualForTickers_(tickers) {
  const wanted = _loadFieldConfigFromControlCenter_().annual;
  const anyWanted = Object.keys(wanted).some(sec => wanted[sec] && wanted[sec].size);
  if (!anyWanted) {
    _log('No Annual fields are active in Control Center — skipping Annual extraction.');
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = _ensureSheetWithHeaders_(ss, EXTRACTOR_CFG.SHEET_ANNUAL, EXTRACTOR_CFG.ANNUAL_HEADERS);
  const existing = _readSheetMapByKey_(sh, ['Ticker','Date','Section','Field']); // upsert key

  tickers.forEachBatch(EXTRACTOR_CFG.BATCH_SIZE, (batch) => {
    const rows = [];
    batch.forEach((ticker) => {
      const json = _fetchFundamentals_(ticker);
      if (!json || !json.Financials) return;

      const fin = json.Financials;
      ['Income_Statement','Balance_Sheet','Cash_Flow'].forEach((section) => {
        const needed = wanted[section];
        if (!needed || needed.size === 0) return;

        const yearly = fin[section] && fin[section].yearly ? fin[section].yearly : [];
        yearly.forEach((obj) => {
          const base = { Ticker: ticker, Date: obj.date || '', FiscalYear: obj.fiscalYear || '', Section: section };
          needed.forEach((fld) => {
            if (Object.prototype.hasOwnProperty.call(obj, fld)) {
              rows.push([
                base.Ticker, base.Date, base.FiscalYear, base.Section, fld, obj[fld], EXTRACTOR_CFG.SOURCE_TAG
              ]);
            }
          });
        });
      });
    });
    _upsertLong_(sh, existing, EXTRACTOR_CFG.ANNUAL_HEADERS, rows, ['Ticker','Date','Section','Field']);
  });
}

function _extractQuarterlyForTickers_(tickers) {
  const wanted = _loadFieldConfigFromControlCenter_().quarterly;
  const anyWanted = Object.keys(wanted).some(sec => wanted[sec] && wanted[sec].size);
  if (!anyWanted) {
    _log('No Quarterly fields are active in Control Center — skipping Quarterly extraction.');
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = _ensureSheetWithHeaders_(ss, EXTRACTOR_CFG.SHEET_QUARTERLY, EXTRACTOR_CFG.QUARTERLY_HEADERS);
  const existing = _readSheetMapByKey_(sh, ['Ticker','Date','Section','Field']); // upsert key

  tickers.forEachBatch(EXTRACTOR_CFG.BATCH_SIZE, (batch) => {
    const rows = [];
    batch.forEach((ticker) => {
      const json = _fetchFundamentals_(ticker);
      if (!json || !json.Financials) return;

      const fin = json.Financials;
      ['Income_Statement','Balance_Sheet','Cash_Flow'].forEach((section) => {
        const needed = wanted[section];
        if (!needed || needed.size === 0) return;

        const quarterly = fin[section] && fin[section].quarterly ? fin[section].quarterly : [];
        quarterly.forEach((obj) => {
          const base = {
            Ticker: ticker,
            Date: obj.date || '',
            FiscalYear: obj.fiscalYear || '',
            FiscalQuarter: obj.fiscalQuarter || '',
            Section: section
          };
          needed.forEach((fld) => {
            if (Object.prototype.hasOwnProperty.call(obj, fld)) {
              rows.push([
                base.Ticker, base.Date, base.FiscalYear, base.FiscalQuarter,
                base.Section, fld, obj[fld], EXTRACTOR_CFG.SOURCE_TAG
              ]);
            }
          });
        });
      });
    });
    _upsertLong_(sh, existing, EXTRACTOR_CFG.QUARTERLY_HEADERS, rows, ['Ticker','Date','Section','Field']);
  });
}

function _extractSplitsDividendsForTickers_(tickers) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(EXTRACTOR_CFG.SHEET_SPLITS_DIVS);
  if (!sh) { _log(`'${EXTRACTOR_CFG.SHEET_SPLITS_DIVS}' not found — skipping.`); return; }
  _ensureHeaders_(sh, EXTRACTOR_CFG.SPLITS_DIVS_HEADERS);

  const existing = _readSheetMapByKey_(sh, ['Ticker','Kind','Date']); // upsert by ticker+kind+date

  tickers.forEachBatch(EXTRACTOR_CFG.BATCH_SIZE, (batch) => {
    const rows = [];
    batch.forEach((ticker) => {
      const json = _fetchFundamentals_(ticker);
      if (!json || !json.SplitsDividends) return;

      const sd = json.SplitsDividends;

      // Dividends
      const divs = Array.isArray(sd.Dividends) ? sd.Dividends : [];
      divs.forEach(d => {
        rows.push([
          ticker, 'Dividend',
          _nz(d.date), _nz(d.declarationDate), _nz(d.recordDate), _nz(d.paymentDate),
          _num(d.dividend), _num(d.adjDividend),
          '', '', '', `decl=${_nz(d.declarationDate)}`
        ]);
      });

      // Splits
      const splits = Array.isArray(sd.Splits) ? sd.Splits : [];
      splits.forEach(s => {
        rows.push([
          ticker, 'Split',
          _nz(s.date), '', '', '',
          '', '', _num(s.forFactor), _num(s.toFactor), _nz(s.ratio), ''
        ]);
      });

      // Last split meta (optional)
      if (sd.LastSplitDate && sd.LastSplitFactor) {
        rows.push([
          ticker, 'LastSplitMeta',
          _nz(sd.LastSplitDate), '', '', '',
          '', '', '', '', _nz(sd.LastSplitFactor), 'from LastSplit*'
        ]);
      }
    });

    _upsertLong_(sh, existing, EXTRACTOR_CFG.SPLITS_DIVS_HEADERS, rows, ['Ticker','Kind','Date']);
  });
}

/** ===================== FETCHERS ===================== **/

function _fetchFundamentals_(ticker) {
  return EXTRACTOR_CFG.MOCK_MODE ? _readMockFromDrive_(ticker) : _fetchFromApi_(ticker);
}

function _fetchFromApi_(ticker) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty(EXTRACTOR_CFG.EODHD_API_KEY_PROP);
  if (!apiKey) throw new Error('Missing API key in Script Properties (set EODHD_API_KEY).');

  const t = ticker.includes('.') ? ticker : `${ticker}.PSE`;
  const url = `${EXTRACTOR_CFG.EODHD_FUNDAMENTALS_URL}/${encodeURIComponent(t)}?api_token=${encodeURIComponent(apiKey)}&fmt=json`;

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`API error ${resp.getResponseCode()} for ${t}: ${resp.getContentText().slice(0,200)}`);
  }
  try {
    return JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error(`JSON parse error for ${t}: ${e}`);
  }
}

function _readMockFromDrive_(ticker) {
  const rawId = EXTRACTOR_CFG.MOCK_FOLDER_ID;
  const fid = _sanitizeFolderIdInput_(rawId);
  if (!fid) {
    throw new Error('Mock Mode ON but Control Center!B7 is empty. Put the Drive folder ID there.');
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(fid); // requires Drive scope
  } catch (e) {
    throw new Error(
      'Could not access the Mock Folder. Ensure:\n' +
      '• appsscript.json scope includes "https://www.googleapis.com/auth/drive"\n' +
      '• You ran _authorizeDriveOnce() and accepted permissions\n' +
      '• The ID is correct (Drive URL /folders/<ID>)\n\nRaw: ' + e
    );
  }

  // If ticker already has suffix (e.g., AC.PSE) → AC.PSE.json ; else → AC.PSE.json
  const fileName = /\.[A-Z]+$/i.test(ticker) ? `${ticker}.json` : `${ticker}.PSE.json`;

  // Iterate files; Drive MIME for JSON can be inconsistent, so match by name
  const it = folder.getFilesByName(fileName);
  if (!it.hasNext()) {
    throw new Error(`Mock JSON not found in folder: ${fileName}`);
  }

  const content = it.next().getBlob().getDataAsString('utf-8');
  try {
    return JSON.parse(content);
  } catch (e) {
    throw new Error(`Mock JSON parse error for ${fileName}: ${e}`);
  }
}

/** ===================== UPSERT HELPERS ===================== **/

function _upsertLong_(sheet, existingMap, headers, rows, keyCols) {
  if (!rows.length) return;

  const headerIndex = _headerIndex_(headers);
  const keyIdx = keyCols.map(c => headerIndex[c]);

  const appendRows = [];
  const updatesByRow = {}; // sheetRow -> rowValues

  rows.forEach(r => {
    const key = keyIdx.map(i => String(r[i])).join('¦');
    const foundRow = existingMap.get(key);
    if (foundRow) {
      updatesByRow[foundRow] = r;
    } else {
      appendRows.push(r);
    }
  });

  // In-place updates
  Object.keys(updatesByRow).forEach(rowNumStr => {
    const rowNum = parseInt(rowNumStr, 10);
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatesByRow[rowNum]]);
  });

  // Appends
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
  }
}

function _readSheetMapByKey_(sheet, keyCols) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const headerIndex = {};
  headers.forEach((h, i) => headerIndex[h] = i);

  const keyIdx = [];
  for (let k = 0; k < keyCols.length; k++) {
    if (!(keyCols[k] in headerIndex)) return new Map(); // empty/new sheet
    keyIdx.push(headerIndex[keyCols[k]]);
  }

  const lastRow = sheet.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const rng = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const values = rng.getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const key = keyIdx.map(ix => String(row[ix])).join('¦');
    map.set(key, i + 2); // absolute row
  }
  return map;
}

/** ===================== SHEET/HEADER HELPERS ===================== **/

function _ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  _ensureHeaders_(sh, headers);
  return sh;
}

function _ensureHeaders_(sh, headers) {
  const neededCols = headers.length;
  if (sh.getMaxColumns() < neededCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), neededCols - sh.getMaxColumns());
  }
  const existing = sh.getRange(1, 1, 1, neededCols).getValues()[0];
  const mismatch = existing.length !== headers.length ||
                   headers.some((h, i) => existing[i] !== h);

  if (mismatch) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getFrozenRows() < 1) {
    sh.setFrozenRows(1);
  }
}

/** ===================== TICKERS & CONTROL CENTER ===================== **/

function _getTickers_() {
  const ss = SpreadsheetApp.getActive();

  // 1) Preferred: configured sheet
  const listFromCfg = (() => {
    const sh = ss.getSheetByName(EXTRACTOR_CFG.TICKERS_SHEET);
    if (!sh) return null;
    const vals = sh.getRange(EXTRACTOR_CFG.TICKERS_RANGE).getValues()
      .map(r => (r[0] || '').toString().trim())
      .filter(Boolean);
    return vals.length ? vals : null;
  })();
  if (listFromCfg) return _normalizeTickersSuffix_(listFromCfg);

  // 2) Fallback: auto-detect a sheet with header "Ticker"
  for (const sh of ss.getSheets()) {
    const header = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]
      .map(v => (v||'').toString().trim().toLowerCase());
    const colIdx = header.indexOf('ticker');
    if (colIdx >= 0) {
      const vals = sh.getRange(2, colIdx+1, Math.max(0, sh.getLastRow()-1), 1)
        .getValues().map(r => (r[0]||'').toString().trim()).filter(Boolean);
      if (vals.length) return _normalizeTickersSuffix_(vals);
    }
  }

  // 3) Fallback: current selection (first column)
  const sel = ss.getActiveRange();
  if (sel) {
    const vals = sel.getValues().map(r => (r[0] || '').toString().trim()).filter(Boolean);
    if (vals.length) return _normalizeTickersSuffix_(vals);
  }

  throw new Error(
    'Tickers not found.\n' +
    'Create a sheet named "Tickers" with A1=Ticker and A2:A list (e.g., AC, ACEN, URC),\n' +
    'or select a single-column range before running.'
  );
}

function _normalizeTickersSuffix_(list) {
  return list.map(v => v.includes('.') ? v : `${v}.PSE`);
}

function _getMockFolderIdFromControlCenter_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Control Center');
    if (!sh) return '';
    const raw = (sh.getRange('B7').getValue() || '').toString().trim();
    return raw;
  } catch (e) {
    Logger.log('Error reading Control Center!B7: ' + e);
    return '';
  }
}

/** ===================== UTILITIES ===================== **/

function _headerIndex_(headers) { const m={}; headers.forEach((h,i)=>m[h]=i); return m; }
function _nz(v)  { return v == null ? '' : v; }
function _num(v) { return (v == null || v === '') ? '' : Number(v); }
function _log(msg) { console.log(EXTRACTOR_CFG.LOG_PREFIX + msg); }

function _sanitizeFolderIdInput_(input) {
  const s = (input || '').trim();
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}

// One-time helper to force Drive OAuth (run from editor)
function _authorizeDriveOnce() {
  const fid = _sanitizeFolderIdInput_(EXTRACTOR_CFG.MOCK_FOLDER_ID);
  if (!fid) throw new Error('Put your Drive folder ID in Control Center!B7 first.');
  const name = DriveApp.getFolderById(fid).getName();
  Logger.log('Drive OK. Folder: ' + name);
}

/** Batch helper */
Array.prototype.forEachBatch = function (batchSize, fn) {
  const arr = this;
  for (let i = 0; i < arr.length; i += batchSize) {
    fn(arr.slice(i, i + batchSize));
    Utilities.sleep(50);
  }
};
