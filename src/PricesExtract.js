/***************************************************************
 * PricesExtract.gs — price + volume extraction (mock or live)
 * Outputs:
 *  - Prices_Latest(Ticker, Close, Volume, Price_Date, Source, Updated_At)
 *  - Raw_Prices(Ticker, Date, Close, Volume)   [de-duplicated upsert]
 *
 * Depends on:
 *  - fetchDailyBars_(ticker, lookbackDays)  // in Prices.gs (mock/live)
 *  - isNum(), toNum()                        // in Utils.gs
 *  - SH_CALC (Calculated_Metrics)            // in Globals.gs
 ***************************************************************/

const SH_PRICES_LATEST = 'Prices_Latest';
const SH_RAW_PRICES    = 'Raw_Prices';

/** Extract latest bar for all tickers in universe → Prices_Latest + refresh Avg30Value cache */
function run_ExtractPricesLatest() {
  const { tickers } = _getTickerUniverse_();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH_PRICES_LATEST) || ss.insertSheet(SH_PRICES_LATEST);

  // Ensure header
  const header = ['Ticker','Close','Volume','Price_Date','Source','Updated_At'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,header.length).setValues([header]);

  const out = [];
  const source = isMockModeOn_ && isMockModeOn_() ? 'mock' : 'api';

  for (const tkr of tickers) {
    const bars = fetchDailyBars_(tkr, 60);       // 60 trading sessions is plenty to find a last bar
    if (!bars.length) continue;
    bars.sort((a,b)=> new Date(a.date) - new Date(b.date));
    const last = bars[bars.length-1];
    if (!isNum(last.close)) continue;

    // also keep your 30d average cache warm
    if (typeof getAvg30ValueCached_ === 'function') {
      try { getAvg30ValueCached_(tkr); } catch (_) {}
    }

    out.push([tkr, last.close, toNum(last.volume), new Date(last.date), source, new Date()]);
  }

  // Upsert by Ticker (simple: rebuild table with latest snapshot)
  sh.clearContents();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if (out.length) sh.getRange(2,1,out.length,header.length).setValues(out);

  // Formats
  const n = Math.max(out.length, 1);
  sh.setFrozenRows(1);
  sh.getRange(2,2,n,1).setNumberFormat('#,##0.00');       // Close
  sh.getRange(2,3,n,1).setNumberFormat('#,##0');          // Volume
  sh.getRange(2,4,n,1).setNumberFormat('yyyy-mm-dd');     // Price_Date
  sh.getRange(2,6,n,1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // Updated_At

  SpreadsheetApp.getActive().toast(`Prices_Latest updated: ${out.length} tickers.`);
}

/** Extract & upsert last N trading days for all tickers → Raw_Prices (de-duplicated) */
function run_ExtractPricesHistory(days) {
  const lookback = Math.max(5, days || 60);
  const { tickers } = _getTickerUniverse_();

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH_RAW_PRICES) || ss.insertSheet(SH_RAW_PRICES);

  // Header
  const header = ['Ticker','Date','Close','Volume'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,header.length).setValues([header]);

  // Build existing key set (Ticker|Date) to avoid duplicates
  const vals = sh.getDataRange().getValues();
  const keys = new Set();
  for (let r=1; r<vals.length; r++) {
    const t = String(vals[r][0] || '');
    const d = vals[r][1] instanceof Date ? vals[r][1].toISOString().slice(0,10) : String(vals[r][1] || '');
    if (t && d) keys.add(`${t}|${d}`);
  }

  const toAppend = [];

  for (const tkr of tickers) {
    const bars = fetchDailyBars_(tkr, lookback);
    if (!bars.length) continue;
    for (const b of bars) {
      const dStr = (b.date instanceof Date) ? b.date.toISOString().slice(0,10) : String(b.date);
      const k = `${tkr}|${dStr}`;
      if (!isNum(b.close) || keys.has(k)) continue;
      keys.add(k);
      toAppend.push([tkr, new Date(dStr), b.close, toNum(b.volume)]);
    }
  }

  if (toAppend.length) sh.getRange(sh.getLastRow()+1, 1, toAppend.length, header.length).setValues(toAppend);

  // Formats
  const nAdded = toAppend.length;
  const lastRow = sh.getLastRow();
  if (nAdded) {
    sh.getRange(lastRow-nAdded+1, 2, nAdded, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(lastRow-nAdded+1, 3, nAdded, 1).setNumberFormat('#,##0.00');
    sh.getRange(lastRow-nAdded+1, 4, nAdded, 1).setNumberFormat('#,##0');
  }
  sh.setFrozenRows(1);

  SpreadsheetApp.getActive().toast(`Raw_Prices upserted: +${nAdded} rows (lookback ${lookback}d).`);
}

/* --------------------- helpers --------------------- */

/** Ticker universe: prefer Calculated_Metrics, fallback to EXTRACTOR_CFG.TICKERS_SHEET */
function _getTickerUniverse_() {
  const ss = SpreadsheetApp.getActive();
  const shCM = ss.getSheetByName(SH_CALC);
  let tickers = [];

  if (shCM && shCM.getLastRow() > 1) {
    const vals = shCM.getDataRange().getValues();
    const H = Object.fromEntries(vals[0].map((h,i)=>[String(h).trim(), i]));
    const iT = H['Ticker'];
    tickers = vals.slice(1).map(r => r[iT]).filter(Boolean);
  }

  if (!tickers.length && typeof EXTRACTOR_CFG === 'object') {
    const shT = ss.getSheetByName(EXTRACTOR_CFG.TICKERS_SHEET);
    if (shT) {
      const vals = shT.getDataRange().getValues();
      tickers = vals.flat().map(v => String(v).trim()).filter(Boolean);
    }
  }

  tickers = Array.from(new Set(tickers)).sort();
  if (!tickers.length) throw new Error('No tickers found in Calculated_Metrics or Tickers sheet.');
  return { tickers };
}
