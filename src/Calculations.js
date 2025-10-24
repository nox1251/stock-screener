/****************************************************************************************
 * Calculations.gs — Per‑Share CAGRs with ≤0→0.01 rule + turnaround markers
 * (robust to header name variants / aliases)
 * -------------------------------------------------------------------------------------
 * Purpose
 *   - Reads Per_Share (latest ~10 FY per ticker expected)
 *   - Computes Latest, 5Y, 9Y, and (optionally) 10Y CAGRs for OpPS, EPS, FCFPS
 *   - Applies ≤0→0.01 floor on endpoints ("turnaround" handling)
 *   - Emits *_AdjustedFlag booleans when any endpoint was floored for that metric
 *   - Upserts into Calculated_Metrics while preserving unknown columns
 *   - **Header tolerant**: accepts common aliases like "Operating Income/Share", "OpIncPS", etc.
 *
 * Public Entry Points (menus/buttons can call any of these):
 *   - menu_ComputeCAGRs()                  → wraps computeCAGRs_() with a UI spinner
 *   - menu_ComputeCAGRs_Adjusted()         → same as above (explicit naming)
 *   - computeCAGRs() / computeCAGRs_()     → core computation
 *   - computeAnnualOpsMetrics_AndUpsert()  → legacy alias (compat)
 *
 * Sheet Contracts
 *   Per_Share (tall, annual):
 *     Ticker | FiscalYear | ... per-share fields ...
 *     Accepted aliases:
 *       - Ticker:            [Ticker, Symbol]
 *       - FiscalYear:        [FiscalYear, FY, Year]
 *       - OpPS (OpIncPS):    [OpPS, OpIncPS, OperatingIncomePS, OperatingIncomePerShare, 'Operating Income/Share', 'OpInc/Share', OpIncomePS]
 *       - EPS:               [EPS, EarningsPS, EarningsPerShare]
 *       - FCFPS:             [FCFPS, 'FCF Per Share', FreeCashFlowPS, 'FCF/Share', FCF_PS]
 *   Calculated_Metrics (wide, by ticker):
 *     Ticker |
 *       OpPS_Latest | OpPS_5Y_CAGR | OpPS_9Y_CAGR | OpPS_10Y_CAGR | OpPS_AdjustedFlag |
 *       EPS_Latest  | EPS_5Y_CAGR  | EPS_9Y_CAGR  | EPS_10Y_CAGR  | EPS_AdjustedFlag  |
 *       FCFPS_Latest| FCFPS_5Y_CAGR| FCFPS_9Y_CAGR| FCFPS_10Y_CAGR| FCFPS_AdjustedFlag|
 *       Debt_to_Equity | Calc_Timestamp
 ****************************************************************************************/

/*********************************** Public Entrypoints ***********************************/
function menu_ComputeCAGRs(){
  return withUiSpinner_(() => computeCAGRs_(), 'Compute CAGRs (≤0→0.01 + markers)');
}
function menu_ComputeCAGRs_Adjusted(){
  return withUiSpinner_(() => computeCAGRs_(), 'Compute CAGRs (≤0→0.01 + markers)');
}
function computeCAGRs(){ return computeCAGRs_(); }
function computeAnnualOpsMetrics_AndUpsert(){ return computeCAGRs_(); } // legacy alias

/**************************************** Core Logic ****************************************/
function computeCAGRs_(){
  const ss = SpreadsheetApp.getActive();
  const perShareSh = ss.getSheetByName('Per_Share');
  if (!perShareSh) throw new Error('Per_Share sheet not found');

  const cmSh = ss.getSheetByName('Calculated_Metrics') || ss.insertSheet('Calculated_Metrics');

  // Read Per_Share into memory
  const ps = getTable_(perShareSh);
  const alias = {
    Ticker: ['Ticker','Symbol'],
    FiscalYear: ['FiscalYear','FY','Year'],
    OpPS: ['OpPS','OpIncPS','OperatingIncomePS','OperatingIncomePerShare','Operating Income/Share','OpInc/Share','OpIncomePS'],
    EPS: ['EPS','EarningsPS','EarningsPerShare'],
    FCFPS: ['FCFPS','FCF Per Share','FreeCashFlowPS','FCF/Share','FCF_PS']
  };
  const cols = resolveCols_(ps.header, alias);

  const byTicker = new Map();
  for (const row of ps.rows) {
    const t = String(row[cols.Ticker] || '').trim(); if (!t) continue;
    const fy = toInt_(row[cols.FiscalYear]);
    const rec = {
      fy,
      OpPS: toNum_(row[cols.OpPS]),
      EPS: toNum_(row[cols.EPS]),
      FCFPS: toNum_(row[cols.FCFPS])
    };
    if (!byTicker.has(t)) byTicker.set(t, []);
    byTicker.get(t).push(rec);
  }

  // Sort ascending by FY
  for (const arr of byTicker.values()) arr.sort((a,b)=>a.fy-b.fy);

  const now = new Date();
  const outMap = new Map();

  for (const [ticker, arr] of byTicker.entries()) {
    if (!arr.length) continue;
    const last = arr[arr.length-1];

    const metrics = ['OpPS','EPS','FCFPS'];
    const result = {Ticker: ticker, Calc_Timestamp: now};

    for (const m of metrics) {
      const latestRaw = toNum_(last[m]);
      const latestAdj = floor01_(latestRaw);
      const latestWasAdj = (latestRaw != null && latestRaw <= 0);
      result[m + '_Latest'] = latestAdj;

      // 5Y CAGR (need ≥6 rows)
      let cagr5 = null, adj5 = false;
      if (arr.length >= 6) {
        const end = arr[arr.length-1][m];
        const start = arr[arr.length-6][m];
        const e = floor01_(end), s = floor01_(start);
        if (isFiniteNum_(e) && isFiniteNum_(s) && s > 0) cagr5 = Math.pow(e/s, 1/5) - 1;
        adj5 = (toNum_(end) != null && end <= 0) || (toNum_(start) != null && start <= 0);
      }
      result[m + '_5Y_CAGR'] = cagr5;

      // 9Y CAGR (need ≥10 rows)
      let cagr9 = null, adj9 = false;
      if (arr.length >= 10) {
        const end = arr[arr.length-1][m];
        const start = arr[arr.length-10][m];
        const e = floor01_(end), s = floor01_(start);
        if (isFiniteNum_(e) && isFiniteNum_(s) && s > 0) cagr9 = Math.pow(e/s, 1/9) - 1;
        adj9 = (toNum_(end) != null && end <= 0) || (toNum_(start) != null && start <= 0);
      }
      result[m + '_9Y_CAGR'] = cagr9;

      // 10Y CAGR (need ≥11 rows) — optional/compat
      let cagr10 = null, adj10 = false;
      if (arr.length >= 11) {
        const end = arr[arr.length-1][m];
        const start = arr[arr.length-11][m];
        const e = floor01_(end), s = floor01_(start);
        if (isFiniteNum_(e) && isFiniteNum_(s) && s > 0) cagr10 = Math.pow(e/s, 1/10) - 1;
        adj10 = (toNum_(end) != null && end <= 0) || (toNum_(start) != null && start <= 0);
      }
      result[m + '_10Y_CAGR'] = cagr10;

      // Turnaround marker for this metric
      result[m + '_AdjustedFlag'] = !!(latestWasAdj || adj5 || adj9 || adj10);
    }

    outMap.set(ticker, result);
  }

  upsertCalculatedMetrics_(cmSh, outMap);
}

/**************************************** Utilities ****************************************/
function getTable_(sh){
  const rng = sh.getDataRange();
  const values = rng.getValues();
  if (values.length < 2) return {header:[], rows:[]};
  const header = values[0].map(v => String(v).trim());
  const rows = values.slice(1);
  return {header, rows};
}

function resolveCols_(header, aliasMap){
  const idx = {};
  const lower = header.map(h => h.toLowerCase());
  for (const [key, aliases] of Object.entries(aliasMap)) {
    let found = -1;
    for (const a of aliases) {
      const j = lower.indexOf(String(a).toLowerCase());
      if (j !== -1) { found = j; break; }
    }
    if (found === -1) {
      throw new Error('Missing required column for ' + key + '. Looked for any of: ' + aliases.join(', '));
    }
    idx[key] = found;
  }
  return idx;
}

function toInt_(v){ const n = Number(v); return Number.isFinite(n) ? Math.trunc(n) : null; }
function toNum_(v){ const n = Number(v); return Number.isFinite(n) ? n : null; }
function isFiniteNum_(v){ return Number.isFinite(Number(v)); }
function floor01_(v){ const n = Number(v); if (!Number.isFinite(n)) return null; return (n <= 0 ? 0.01 : n); }

function upsertCalculatedMetrics_(sh, mapByTicker){
  const owned = [
    'Ticker',
    'OpPS_Latest','OpPS_5Y_CAGR','OpPS_9Y_CAGR','OpPS_10Y_CAGR','OpPS_AdjustedFlag',
    'EPS_Latest','EPS_5Y_CAGR','EPS_9Y_CAGR','EPS_10Y_CAGR','EPS_AdjustedFlag',
    'FCFPS_Latest','FCFPS_5Y_CAGR','FCFPS_9Y_CAGR','FCFPS_10Y_CAGR','FCFPS_AdjustedFlag',
    'Calc_Timestamp'
  ];

  const tbl = getTable_(sh);
  let header = tbl.header.length ? tbl.header : ['Ticker'];
  for (const h of owned) if (!header.includes(h)) header.push(h);

  const tIdx = header.indexOf('Ticker');
  if (tIdx === -1) throw new Error('Calculated_Metrics requires a Ticker column');

  const existing = new Map();
  for (const row of tbl.rows) {
    const t = String(row[tIdx] || '').trim();
    if (t) existing.set(t, row.slice());
  }

  for (const [ticker, rec] of mapByTicker.entries()) {
    let row = existing.get(ticker);
    if (!row) { row = new Array(header.length).fill(''); row[tIdx] = ticker; existing.set(ticker, row); }
    for (const [k, v] of Object.entries(rec)) {
      const col = header.indexOf(k);
      if (col === -1) continue; // ignore keys we don't own
      row[col] = (v instanceof Date) ? v : v;
    }
  }

  const out = [header];
  for (const [t, row] of Array.from(existing.entries()).sort((a,b)=>a[0].localeCompare(b[0]))) {
    if (row.length < header.length) row.length = header.length;
    out.push(row);
  }

  sh.clear();
  sh.getRange(1,1,out.length,header.length).setValues(out);
  sh.getRange(1,1,1,header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  autoSize_(sh);
}

function autoSize_(sh){ try{ sh.autoResizeColumns(1, Math.min(20, sh.getLastColumn())); }catch(e){} }

/*************************************** UI helper ****************************************/
function withUiSpinner_(fn,label){
  try{ const ui=SpreadsheetApp.getUi(); const start=new Date(); const r=fn(); const s=((new Date()-start)/1000).toFixed(1); ui.alert(label+' — done in '+s+'s'); return r; }
  catch(err){ SpreadsheetApp.getUi().alert('Error: '+err.message); throw err; }
}
