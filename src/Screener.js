/****************************************************************************************
 * Screener.gs — Classic layout with persistent filter area (A1:D8 preserved)
 * ---------------------------------------------------------------------------------
 * Sheet layout:
 *  A1:D8  →  Filter area (never cleared)
 *  Row 10 →  Header row for data table
 *  Row 11 →  Data start
 *
 * Filter block values:
 *  C4: minimum 5Y CAGR (optional)
 *  C5: maximum Debt/Equity (default 1.5)
 *  C6: minimum 9Y/10Y CAGR (optional)
 *
 * Logic additions:
 *  - Applies ≤0 → 0.01 floor for OpPS when computing Payback.
 *  - Appends “*” to Operating Income/Share if any ≤0→0.01 adjustment applied.
 *  - Enforces acceleration rule (5Y > 9Y or 10Y) + optional filter thresholds.
 *  - Never clears A1:D8 area when rebuilding output.
 ****************************************************************************************/

const SCREENER_LAYOUT = {
  sheetName: 'Screener',
  headerRow: 10,
  startRow: 11,
  headerColumns: ['Ticker','Operating Income/Share','OpInc 5yr CAGR','OpInc 9yr CAGR','YoY Growth','Debt/Equity','Last Closed Price','Payback Period','Avg30Value']
};

function menu_RunScreener(){
  withUiSpinner_(() => buildClassicScreener_(), 'Run Screener');
}

function buildClassicScreener_(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SCREENER_LAYOUT.sheetName) || ss.insertSheet(SCREENER_LAYOUT.sheetName);
  const calc = ss.getSheetByName('Calculated_Metrics');
  if (!calc) throw new Error('Calculated_Metrics not found');
  const prStats = ss.getSheetByName('Prices_Stats');
  const prLatest = ss.getSheetByName('Prices_Latest');

  // Read filter thresholds
  const min5 = readOptionalNumber_(sh, 4, 3);
  const maxDE = readOptionalNumber_(sh, 5, 3, 1.5);
  const minLong = readOptionalNumber_(sh, 6, 3);

  const tbl = getTable_(calc);
  const c = indexCols_(tbl.header,[
    'Ticker','OpPS_Latest','OpPS_5Y_CAGR','OpPS_9Y_CAGR','OpPS_10Y_CAGR','OpPS_AdjustedFlag','Debt_to_Equity'
  ],true);

  const prices = readPricesMap_(prStats, prLatest);
  const has9 = c.OpPS_9Y_CAGR != null && c.OpPS_9Y_CAGR !== -1;
  const longKey = has9 ? 'OpPS_9Y_CAGR' : 'OpPS_10Y_CAGR';

  const rows = [];
  for (const r of tbl.rows) {
    const t = String(r[c.Ticker]||'').trim(); if (!t) continue;
    const op = toNum_(r[c.OpPS_Latest]);
    const g5 = toNum_(r[c.OpPS_5Y_CAGR]);
    const gl = toNum_(r[c[longKey]]);
    const adj = toBool_(r[c.OpPS_AdjustedFlag]);
    const de = toNum_(r[c.Debt_to_Equity]);

    const px = prices.get(t) || {};
    const last = toNum_(px.LastClose);
    const avg30 = toNum_(px.Avg30Value_30d);

    // Filters
    if (!(toNumOrZero_(g5) > toNumOrZero_(gl))) continue;
    if (isFiniteNum_(min5) && !(toNumOrZero_(g5) >= min5)) continue;
    if (isFiniteNum_(maxDE) && isFiniteNum_(de) && !(de <= maxDE)) continue;
    if (isFiniteNum_(minLong) && !(toNumOrZero_(gl) >= minLong)) continue;

    // Payback calculation (≤0→0.01 floor)
    let pay = null;
    if (isFiniteNum_(last) && isFiniteNum_(g5)) {
      const opAdj = adjustFloor_(op);
      if (isFiniteNum_(opAdj) && g5 !== -1) {
        const num = Math.log((last * g5 / opAdj) + (1+g5));
        const den = Math.log(1+g5);
        if (isFiniteNum_(num) && isFiniteNum_(den) && den !== 0) pay = (num/den)-1;
      }
    }

    const opDisp = op != null ? (adj ? op+'*' : op) : '';
    rows.push([t, opDisp, g5, gl, null, de, last, pay, avg30]);
  }

  // Preserve A1:D8 — only clear below row 9
  const lastRow = sh.getMaxRows();
  if (lastRow > 9) sh.getRange(9+1,1,lastRow-9,sh.getMaxColumns()).clear();

  // Write header + data
  sh.getRange(SCREENER_LAYOUT.headerRow,1,1,SCREENER_LAYOUT.headerColumns.length)
    .setValues([SCREENER_LAYOUT.headerColumns])
    .setFontWeight('bold');

  if (rows.length){
    sh.getRange(SCREENER_LAYOUT.startRow,1,rows.length,SCREENER_LAYOUT.headerColumns.length).setValues(rows);
  }

  // Formatting
  const n=Math.max(1,rows.length);
  sh.getRange(SCREENER_LAYOUT.startRow,3,n,2).setNumberFormat('0.00%');
  sh.getRange(SCREENER_LAYOUT.startRow,6,n,1).setNumberFormat('0.00');
  sh.getRange(SCREENER_LAYOUT.startRow,7,n,1).setNumberFormat('#,##0.00');
  sh.getRange(SCREENER_LAYOUT.startRow,8,n,1).setNumberFormat('0.0');
  sh.getRange(SCREENER_LAYOUT.startRow,9,n,1).setNumberFormat('#,##0');
  sh.setFrozenRows(SCREENER_LAYOUT.headerRow);
  autoSize_(sh);
}

/************* helpers *************/
function getTable_(sh){const v=sh.getDataRange().getValues();if(v.length<2)return{header:[],rows:[]};return{header:v[0].map(String),rows:v.slice(1)};}
function indexCols_(h,n,s){const i={};for(const x of n){const j=h.indexOf(x);if(j==-1){if(s){i[x]=-1;continue;}throw new Error('Missing '+x);}i[x]=j;}return i;}
function toNum_(v){const n=Number(v);return Number.isFinite(n)?n:null;}
function toNumOrZero_(v){const n=Number(v);return Number.isFinite(n)?n:0;}
function isFiniteNum_(v){return Number.isFinite(Number(v));}
function toBool_(v){return String(v).toLowerCase()==='true'||v===true;}
function adjustFloor_(v){const n=Number(v);if(!Number.isFinite(n))return null;return n<=0?0.01:n;}
function readPricesMap_(ps,pl){const m=new Map();if(ps){const t=getTable_(ps);const c=indexCols_(t.header,['Ticker','LastClose','Avg30Value_30d'],true);for(const r of t.rows){const tk=String(r[c.Ticker]||'').trim();if(!tk)continue;m.set(tk,{LastClose:toNum_(r[c.LastClose]),Avg30Value_30d:toNum_(r[c.Avg30Value_30d])});}}if(pl){const t=getTable_(pl);const c=indexCols_(t.header,['Ticker','Close'],true);for(const r of t.rows){const tk=String(r[c.Ticker]||'').trim();if(!tk)continue;if(!m.has(tk)||!isFiniteNum_(m.get(tk).LastClose))m.set(tk,{LastClose:toNum_(r[c.Close]),Avg30Value_30d:(m.get(tk)?.Avg30Value_30d)||null});}}return m;}
function readOptionalNumber_(sh,row,col,def){try{const v=sh.getRange(row,col).getValue();const n=Number(v);return Number.isFinite(n)?n:def;}catch(e){return def;}}
function withUiSpinner_(fn,label){try{const ui=SpreadsheetApp.getUi();const s=new Date();const r=fn();const t=((new Date()-s)/1000).toFixed(1);ui.alert(label+' — done in '+t+'s');return r;}catch(e){SpreadsheetApp.getUi().alert('Error: '+e.message);throw e;}}
