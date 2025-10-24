/***************************************************************
 * Prices.gs — price & volume utilities (optimized)
 * -------------------------------------------------------------
 * Sheets (created if missing):
 *   - Prices_Latest: Ticker | Close | Volume | Price_Date | Source | Updated_At
 *   - Prices_Stats : Ticker | LastClose | Avg30Value_30d | BarsThrough | Updated_At
 *
 * Public entrypoints (used by menu):
 *   - run_ExtractPricesLatest()
 *   - run_ExtractPricesHistory(days)       // optional, keeps a long raw table
 *   - refreshPriceStats_ForAllTickers()
 *   - setMockPricesFolderId_NOW(), setMockMode_ON(), setMockMode_OFF()
 *
 * Fetch:
 *   - fetchDailyBars_(ticker, lookbackDays) -> [{date, close, volume}]
 *     * Mock Mode: reads Drive JSONs from a folder you set once
 *     * Live: EODHD /eod endpoint (needs EODHD_API_KEY in Script Properties)
 ***************************************************************/

/* ====================== Sheet name (no collisions) ====================== */
const PRICE_STATS_SHEET = (typeof SH_PRICE_STATS === 'string' ? SH_PRICE_STATS : 'Prices_Stats');

/* ======================== Mock Mode controls ======================== */

function setMockPricesFolderId_NOW() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Mock Prices', 'Enter Drive Folder ID (where the JSON files live):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const id = (resp.getResponseText() || '').trim();
  try {
    const f = DriveApp.getFolderById(id); // throws if invalid
    PropertiesService.getUserProperties().setProperty('MOCK_PRICES_FOLDER_ID', id);
    ui.alert('✅ Saved Mock Prices Folder: ' + f.getName());
  } catch (e) {
    ui.alert('❌ Invalid folder ID:\n' + e);
  }
}

function getMockPricesFolderId_() {
  return PropertiesService.getUserProperties().getProperty('MOCK_PRICES_FOLDER_ID') || '';
}

function isMockModeOn_() {
  const v = PropertiesService.getUserProperties().getProperty('MOCK_MODE');
  return String(v).toLowerCase() === 'true';
}
function setMockMode_ON()  { PropertiesService.getUserProperties().setProperty('MOCK_MODE', 'true');  SpreadsheetApp.getUi().alert('Mock Mode = ON'); }
function setMockMode_OFF() { PropertiesService.getUserProperties().setProperty('MOCK_MODE', 'false'); SpreadsheetApp.getUi().alert('Mock Mode = OFF'); }

/* ======================== Ticker Universe ======================== */

function _getTickerUniverse_() {
  const ss = SpreadsheetApp.getActive();
  let tickers = [];

  // Prefer Calculated_Metrics
  const shCM = ss.getSheetByName(typeof SH_CALC === 'string' ? SH_CALC : 'Calculated_Metrics');
  if (shCM && shCM.getLastRow() > 1) {
    const vals = shCM.getDataRange().getValues();
    const H = Object.fromEntries(vals[0].map((h,i)=>[String(h).trim(), i]));
    const iT = H['Ticker'];
    if (iT != null) tickers = vals.slice(1).map(r => r[iT]).filter(Boolean);
  }

  // Fallback: a sheet named "Tickers" with one column
  if (!tickers.length) {
    const shT = ss.getSheetByName('Tickers');
    if (shT) tickers = shT.getRange(1,1,shT.getLastRow(),1).getValues().flat().filter(Boolean);
  }

  tickers = Array.from(new Set(tickers)).sort();
  if (!tickers.length) throw new Error('No tickers found in Calculated_Metrics or Tickers sheet.');
  return { tickers };
}

/* ===================== Fast Mock folder index ===================== */

function _buildMockFileIndex_() {
  const folderId = getMockPricesFolderId_();
  if (!folderId) return { folderId: '', byName: {} };

  const cache = CacheService.getScriptCache();
  const key = 'MOCK_FILE_INDEX_' + folderId;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  let byName = {};
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    byName[f.getName().toUpperCase()] = f.getId();
  }
  const idx = { folderId, byName };
  cache.put(key, JSON.stringify(idx), 600); // 10 minutes
  return idx;
}

function _resolveMockFileId_(idx, ticker) {
  const cands = [
    `${ticker}.json`,
    `${ticker}`, // already includes .json
    `${String(ticker).replace(/\.PSE$/i,'')}.json`
  ];
  for (const n of cands) {
    const id = idx.byName[n.toUpperCase()];
    if (id) return id;
  }
  return null;
}

/* ==================== Mock JSON quick reader ==================== */

function _readMockBarsLastN_(fileId, n) {
  let arr;
  try { arr = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString()); }
  catch (_) { return []; }
  if (!Array.isArray(arr) || arr.length === 0) return [];
  const start = Math.max(0, arr.length - n);
  const slice = arr.slice(start); // JSON already oldest→newest
  const out = [];
  for (const r of slice) {
    const close = (typeof r.close === 'number') ? r.close : (typeof r.Close === 'number' ? r.Close : null);
    if (!isNum(close)) continue;
    const vol   = (typeof r.volume === 'number') ? r.volume : (typeof r.Volume === 'number' ? r.Volume : null);
    out.push({ date: r.date || r.Date, close, volume: vol });
  }
  return out;
}

/* ==================== Fetch (Mock or Live) ==================== */

function fetchDailyBars_(ticker, lookbackDays) {
  // ---- Mock path
  if (isMockModeOn_()) {
    const idx = _buildMockFileIndex_();
    const fileId = _resolveMockFileId_(idx, ticker);
    if (!fileId) { Logger.log(`fetchDailyBars_: mock file not found for ${ticker}`); return []; }
    return _readMockBarsLastN_(fileId, Math.max(5, lookbackDays || 60));
  }

  // ---- Live EODHD path
  const apiKey = PropertiesService.getScriptProperties().getProperty('EODHD_API_KEY');
  if (!apiKey) { Logger.log('fetchDailyBars_: missing EODHD_API_KEY'); return []; }

  const to = new Date();
  const from = new Date(to.getTime() - (Math.max(5, lookbackDays || 60) * 24 * 3600 * 1000));
  const fmt = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const url = `https://eodhd.com/api/eod/${encodeURIComponent(ticker)}?from=${fmt(from)}&to=${fmt(to)}&period=d&fmt=json&api_token=${encodeURIComponent(apiKey)}`;

  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) { Logger.log(`fetchDailyBars_: ${ticker} HTTP ${res.getResponseCode()}`); return []; }
    const json = JSON.parse(res.getContentText());
    if (!Array.isArray(json)) return [];
    return json.map(r => ({
      date: r.date || r.Date,
      close: (typeof r.close === 'number') ? r.close : (typeof r.Close === 'number' ? r.Close : null),
      volume: (typeof r.volume === 'number') ? r.volume : (typeof r.Volume === 'number' ? r.Volume : null)
    })).filter(b => isNum(b.close));
  } catch (e) {
    Logger.log(`fetchDailyBars_: request/parse failed for ${ticker}: ${e}`);
    return [];
  }
}

/* ================== Batch cache: avg30 + last close ================== */

function getAvg30ValueCached_(ticker) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRICE_STATS_SHEET) || ss.insertSheet(PRICE_STATS_SHEET);
  if (sh.getLastRow() === 0)
    sh.getRange(1,1,1,5).setValues([['Ticker','LastClose','Avg30Value_30d','BarsThrough','Updated_At']]);

  const vals = sh.getDataRange().getValues();
  const H = Object.fromEntries(vals[0].map((h,i)=>[String(h).trim(), i]));
  const iT = H['Ticker'], iC = H['LastClose'], iA = H['Avg30Value_30d'], iB = H['BarsThrough'], iU = H['Updated_At'];

  // find row
  let row = -1;
  for (let r=1; r<vals.length; r++) if (vals[r][iT] === ticker) { row = r; break; }

  // if today's cache exists, return it
  const today = new Date().toDateString();
  if (row !== -1 && vals[row][iU] instanceof Date && vals[row][iU].toDateString() === today) {
    return { lastClose: vals[row][iC], avg30Value: vals[row][iA] };
  }

  // refresh from fetcher (mock or live)
  const bars = fetchDailyBars_(ticker, 60);
  if (!bars.length) return { lastClose: '', avg30Value: '' };

  const last = bars[bars.length - 1];
  const recent = bars.slice(-30).filter(b => isNum(b.volume) && b.volume > 0);
  const avg30 = recent.length ? recent.reduce((s,b)=> s + (b.close * b.volume), 0) / recent.length : '';

  const out = [ticker, last.close, avg30, recent.length, new Date()];
  if (row === -1) sh.appendRow(out);
  else sh.getRange(row+1, 1, 1, out.length).setValues([out]);

  // light formatting
  const n = sh.getLastRow();
  sh.setFrozenRows(1);
  if (n > 1) {
    sh.getRange(2,2,n-1,1).setNumberFormat('#,##0.00'); // LastClose
    sh.getRange(2,3,n-1,1).setNumberFormat('#,##0.00'); // Avg30Value
    sh.getRange(2,4,n-1,1).setNumberFormat('0');        // BarsThrough
    sh.getRange(2,5,n-1,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  return { lastClose: last.close, avg30Value: avg30 };
}

function getLatestPriceCached_(ticker) {
  const sh = SpreadsheetApp.getActive().getSheetByName(PRICE_STATS_SHEET);
  if (!sh) return '';
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return getAvg30ValueCached_(ticker).lastClose || '';
  const H = Object.fromEntries(vals[0].map((h,i)=>[String(h).trim(), i]));
  const iT = H['Ticker'], iC = H['LastClose'], iU = H['Updated_At'];
  const today = new Date().toDateString();
  for (let r=1; r<vals.length; r++) {
    if (vals[r][iT] === ticker && vals[r][iU] instanceof Date && vals[r][iU].toDateString() === today) {
      return vals[r][iC];
    }
  }
  return getAvg30ValueCached_(ticker).lastClose || '';
}

/* ====================== Extractors (batch, fast) ====================== */

function run_ExtractPricesLatest() {
  const { tickers } = _getTickerUniverse_();
  const ss = SpreadsheetApp.getActive();

  const shLatest = ss.getSheetByName('Prices_Latest') || ss.insertSheet('Prices_Latest');
  const shStats  = ss.getSheetByName(PRICE_STATS_SHEET) || ss.insertSheet(PRICE_STATS_SHEET);

  const headerLatest = ['Ticker','Close','Volume','Price_Date','Source','Updated_At'];
  const headerStats  = ['Ticker','LastClose','Avg30Value_30d','BarsThrough','Updated_At'];
  if (shLatest.getLastRow() === 0) shLatest.getRange(1,1,1,headerLatest.length).setValues([headerLatest]);
  if (shStats.getLastRow()  === 0) shStats.getRange(1,1,1,headerStats.length).setValues([headerStats]);

  const now = new Date();
  const isMock = isMockModeOn_();
  const outLatest = [];
  const outStats  = [];

  if (isMock) {
    const idx = _buildMockFileIndex_();
    for (const t of tickers) {
      const fid = _resolveMockFileId_(idx, t);
      if (!fid) continue;
      const bars = _readMockBarsLastN_(fid, 60);
      if (!bars.length) continue;

      const last = bars[bars.length - 1];
      const recent = bars.slice(-30).filter(b => isNum(b.volume) && b.volume > 0);
      const avg30 = recent.length ? recent.reduce((s,b)=> s + (b.close*b.volume), 0)/recent.length : '';

      outLatest.push([t, last.close, (isNum(last.volume) ? last.volume : ''), last.date ? new Date(last.date) : '', 'mock', now]);
      outStats .push([t, last.close, avg30, recent.length, now]);
    }
  } else {
    for (const t of tickers) {
      const bars = fetchDailyBars_(t, 60);
      if (!bars.length) continue;
      const last = bars[bars.length - 1];
      const recent = bars.slice(-30).filter(b => isNum(b.volume) && b.volume > 0);
      const avg30 = recent.length ? recent.reduce((s,b)=> s + (b.close*b.volume), 0)/recent.length : '';

      outLatest.push([t, last.close, (isNum(last.volume) ? last.volume : ''), last.date ? new Date(last.date) : '', 'api', now]);
      outStats .push([t, last.close, avg30, recent.length, now]);
    }
  }

  if (outLatest.length) {
    shLatest.clearContents();
    shLatest.getRange(1,1,1,headerLatest.length).setValues([headerLatest]);
    shLatest.getRange(2,1,outLatest.length,headerLatest.length).setValues(outLatest);
    const n = outLatest.length;
    shLatest.setFrozenRows(1);
    shLatest.getRange(2,2,n,1).setNumberFormat('#,##0.00');      // Close
    shLatest.getRange(2,3,n,1).setNumberFormat('#,##0');         // Volume
    shLatest.getRange(2,4,n,1).setNumberFormat('yyyy-mm-dd');    // Price_Date
    shLatest.getRange(2,6,n,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  if (outStats.length) {
    shStats.clearContents();
    shStats.getRange(1,1,1,headerStats.length).setValues([headerStats]);
    shStats.getRange(2,1,outStats.length,headerStats.length).setValues(outStats);
    const n = outStats.length;
    shStats.setFrozenRows(1);
    shStats.getRange(2,2,n,1).setNumberFormat('#,##0.00');       // LastClose
    shStats.getRange(2,3,n,1).setNumberFormat('#,##0.00');       // Avg30Value
    shStats.getRange(2,4,n,1).setNumberFormat('0');              // BarsThrough
    shStats.getRange(2,5,n,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  SpreadsheetApp.getActive().toast(`Prices updated — Latest:${outLatest.length} / Stats:${outStats.length}`);
}

/* ========== Optional: History extractor (dedup append, still fast) ========== */

function run_ExtractPricesHistory(days) {
  const lookback = Math.max(5, days || 60);
  const { tickers } = _getTickerUniverse_();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Raw_Prices') || ss.insertSheet('Raw_Prices');
  const header = ['Ticker','Date','Close','Volume'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,header.length).setValues([header]);

  // build existing key set to dedupe (Ticker|yyyy-mm-dd)
  const vals = sh.getDataRange().getValues();
  const keys = new Set();
  for (let r=1; r<vals.length; r++) {
    const t = String(vals[r][0] || '');
    const d = vals[r][1] instanceof Date ? Utilities.formatDate(vals[r][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(vals[r][1] || '');
    if (t && d) keys.add(`${t}|${d}`);
  }

  const isMock = isMockModeOn_();
  const toAppend = [];

  if (isMock) {
    const idx = _buildMockFileIndex_();
    for (const t of tickers) {
      const fid = _resolveMockFileId_(idx, t);
      if (!fid) continue;
      const bars = _readMockBarsLastN_(fid, lookback);
      for (const b of bars) {
        const d = String(b.date || '').slice(0,10);
        const k = `${t}|${d}`;
        if (!isNum(b.close) || keys.has(k)) continue;
        keys.add(k);
        toAppend.push([t, new Date(d), b.close, isNum(b.volume) ? b.volume : '']);
      }
    }
  } else {
    for (const t of tickers) {
      const bars = fetchDailyBars_(t, lookback);
      for (const b of bars) {
        const d = String(b.date || '').slice(0,10);
        const k = `${t}|${d}`;
        if (!isNum(b.close) || keys.has(k)) continue;
        keys.add(k);
        toAppend.push([t, new Date(d), b.close, isNum(b.volume) ? b.volume : '']);
      }
    }
  }

  if (toAppend.length) {
    sh.getRange(sh.getLastRow()+1, 1, toAppend.length, header.length).setValues(toAppend);
    const n = toAppend.length;
    const start = sh.getLastRow() - n + 1;
    sh.setFrozenRows(1);
    sh.getRange(start, 2, n, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(start, 3, n, 1).setNumberFormat('#,##0.00');
    sh.getRange(start, 4, n, 1).setNumberFormat('#,##0');
  }

  SpreadsheetApp.getActive().toast(`Raw_Prices upserted: +${toAppend.length} rows (lookback ${lookback}d).`);
}

/* ============================ Admin ============================= */

function refreshPriceStats_ForAllTickers() {
  const { tickers } = _getTickerUniverse_();
  const now = new Date();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PRICE_STATS_SHEET) || ss.insertSheet(PRICE_STATS_SHEET);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,5).setValues([['Ticker','LastClose','Avg30Value_30d','BarsThrough','Updated_At']]);

  const byTicker = new Map();  // ticker → [LastClose, Avg30, Bars, Updated]
  if (isMockModeOn_()) {
    const idx = _buildMockFileIndex_();
    for (const t of tickers) {
      const fid = _resolveMockFileId_(idx, t);
      if (!fid) continue;
      const bars = _readMockBarsLastN_(fid, 60);
      if (!bars.length) continue;
      const last = bars[bars.length - 1];
      const recent = bars.slice(-30).filter(b => isNum(b.volume) && b.volume > 0);
      const avg30 = recent.length ? recent.reduce((s,b)=> s + (b.close*b.volume), 0)/recent.length : '';
      byTicker.set(t, [last.close, avg30, recent.length, now]);
    }
  } else {
    for (const t of tickers) {
      const bars = fetchDailyBars_(t, 60);
      if (!bars.length) continue;
      const last = bars[bars.length - 1];
      const recent = bars.slice(-30).filter(b => isNum(b.volume) && b.volume > 0);
      const avg30 = recent.length ? recent.reduce((s,b)=> s + (b.close*b.volume), 0)/recent.length : '';
      byTicker.set(t, [last.close, avg30, recent.length, now]);
    }
  }

  // Rebuild the table in one write
  const rows = Array.from(byTicker.entries()).map(([t,v]) => [t, ...v]);
  sh.clearContents();
  sh.getRange(1,1,1,5).setValues([['Ticker','LastClose','Avg30Value_30d','BarsThrough','Updated_At']]);
  if (rows.length) sh.getRange(2,1,rows.length,5).setValues(rows);

  const n = rows.length;
  if (n) {
    sh.setFrozenRows(1);
    sh.getRange(2,2,n,1).setNumberFormat('#,##0.00');
    sh.getRange(2,3,n,1).setNumberFormat('#,##0.00');
    sh.getRange(2,4,n,1).setNumberFormat('0');
    sh.getRange(2,5,n,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  SpreadsheetApp.getActive().toast(`Prices_Stats refreshed: ${rows.length} tickers.`);
}
