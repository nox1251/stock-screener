/***************************************************************
 * Per Share Builder (robust Debt/Equity mapping)
 * Inputs: Raw_Fundamentals_Annual (long table from extractor)
 * Output: Per_Share sheet with per-share metrics (latest 10 FY)
 ***************************************************************/

const PER_SHARE_CFG = {
  SOURCE_SHEET: 'Raw_Fundamentals_Annual',
  OUTPUT_SHEET: 'Per_Share',
  HEADERS: [
    'Ticker','FiscalYear',
    'RevenuePerShare','GrossProfitPerShare','OperatingIncomePerShare','NetIncomePerShare',
    'EquityPerShare','DebtToEquity'
  ],
  // Field keys we expect in the long table (TitleCase)
  FIELDS: {
    IS: {
      Revenue: 'Revenue',
      GrossProfit: 'GrossProfit',
      OperatingIncome: 'OperatingIncome',
      NetIncome: 'NetIncome',
      SharesDiluted: 'SharesDiluted'
    },
    BS: {
      // primary
      Equity: 'TotalStockholdersEquity',
      Debt: 'TotalDebt',
      SharesDiluted: 'SharesDiluted',
      // accepted fallbacks seen in some payloads / configs
      EquityAlts: ['TotalStockholderEquity','TotalEquity','StockholdersEquity'],
      DebtAlts: ['TotalLiab','TotalLiabilities']
    }
  },
  DEC_FMT: '#,##0.000',
  LOG_PREFIX: '[PerShare] '
};

/** Menu target */
function buildPerShare_Full() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(PER_SHARE_CFG.SOURCE_SHEET);
  if (!src) throw new Error(`Source sheet "${PER_SHARE_CFG.SOURCE_SHEET}" not found.`);

  const data = _ps_readLongTable_(src);
  const byTKY = _ps_indexByTickerYear_(data);
  const tickers = Object.keys(byTKY).sort();

  const outRows = [];
  for (const tkr of tickers) {
    // latest 10 FY that have valid diluted shares
    const years = Object.keys(byTKY[tkr]).map(y => +y).sort((a,b)=>b-a);
    const picked = [];
    for (const y of years) {
      const row = byTKY[tkr][y];
      const shares = _ps_validShares_(row);
      if (shares > 0) picked.push({ y, row, shares });
      if (picked.length >= 10) break;
    }

    for (const {y, row, shares} of picked.sort((a,b)=>a.y-b.y)) {
      const rev  = _num(row.is.Revenue);
      const gp   = _num(row.is.GrossProfit);
      const opi  = _num(row.is.OperatingIncome);
      const ni   = _num(row.is.NetIncome);

      const eqRaw = _firstNonNull_([
        row.bs.Equity,
        row.bs.EquityAlt1, row.bs.EquityAlt2, row.bs.EquityAlt3
      ]);
      const debtRaw = _firstNonNull_([
        row.bs.Debt,
        row.bs.DebtAlt1, row.bs.DebtAlt2
      ]);

      const eq   = _num(eqRaw);
      const debt = _num(debtRaw);

      const revenuePS = _ps_safeDiv_(rev, shares);
      const gpPS      = _ps_safeDiv_(gp, shares);
      const opiPS     = _ps_safeDiv_(opi, shares);
      const niPS      = _ps_safeDiv_(ni, shares);
      const eqPS      = _ps_safeDiv_(eq, shares);
      const dte       = _ps_safeDiv_(debt, eq); // ratio

      outRows.push([
        tkr, y,
        _ps_r3_(revenuePS),
        _ps_r3_(gpPS),
        _ps_r3_(opiPS),
        _ps_r3_(niPS),
        _ps_r3_(eqPS),
        _ps_r3_(dte)
      ]);
    }
  }

  _ps_writeOutput_(ss, outRows);
  _ps_log(`Built Per_Share for ${tickers.length} tickers, ${outRows.length} rows.`);
}

/** ================ Internals ================= */

function _ps_readLongTable_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 6) return [];

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const idx = _ps_idx_(headers, [
    'Ticker','Date','FiscalYear','Section','Field','Value'
  ]);

  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const rows = [];
  for (let i=0;i<values.length;i++) {
    const r = values[i];
    rows.push({
      Ticker:   (r[idx.Ticker]||'').toString().trim(),
      Date:     r[idx.Date],
      FiscalYear: parseInt(r[idx.FiscalYear],10),
      Section:  (r[idx.Section]||'').toString().trim(),
      Field:    (r[idx.Field]||'').toString().trim(),
      Value:    _num(r[idx.Value])
    });
  }
  return rows;
}

function _ps_indexByTickerYear_(rows) {
  const out = {};
  const IS = PER_SHARE_CFG.FIELDS.IS;
  const BS = PER_SHARE_CFG.FIELDS.BS;

  rows.forEach(r => {
    if (!r.Ticker || !r.FiscalYear) return;
    const t = r.Ticker.includes('.PSE') ? r.Ticker : (r.Ticker + '.PSE');

    out[t] = out[t] || {};
    out[t][r.FiscalYear] = out[t][r.FiscalYear] || { is:{}, bs:{} };

    // Bucket by section
    if (r.Section === 'Income_Statement') {
      if (r.Field === IS.Revenue)              out[t][r.FiscalYear].is.Revenue = r.Value;
      else if (r.Field === IS.GrossProfit)     out[t][r.FiscalYear].is.GrossProfit = r.Value;
      else if (r.Field === IS.OperatingIncome) out[t][r.FiscalYear].is.OperatingIncome = r.Value;
      else if (r.Field === IS.NetIncome)       out[t][r.FiscalYear].is.NetIncome = r.Value;
      else if (r.Field === IS.SharesDiluted)   out[t][r.FiscalYear].is.SharesDiluted = r.Value;

    } else if (r.Section === 'Balance_Sheet') {
      if (r.Field === BS.Equity)               out[t][r.FiscalYear].bs.Equity = r.Value;
      else if (r.Field === BS.Debt)            out[t][r.FiscalYear].bs.Debt = r.Value;
      else if (r.Field === BS.SharesDiluted)   out[t][r.FiscalYear].bs.SharesDiluted = r.Value;

      // Fallbacks we accept (so Debt/Equity won’t be zero if TotalDebt isn’t present)
      else if (r.Field === (BS.DebtAlts[0]))   out[t][r.FiscalYear].bs.DebtAlt1 = r.Value;       // TotalLiab
      else if (r.Field === (BS.DebtAlts[1]))   out[t][r.FiscalYear].bs.DebtAlt2 = r.Value;       // TotalLiabilities

      else if (r.Field === (BS.EquityAlts[0])) out[t][r.FiscalYear].bs.EquityAlt1 = r.Value;     // TotalStockholderEquity
      else if (r.Field === (BS.EquityAlts[1])) out[t][r.FiscalYear].bs.EquityAlt2 = r.Value;     // TotalEquity
      else if (r.Field === (BS.EquityAlts[2])) out[t][r.FiscalYear].bs.EquityAlt3 = r.Value;     // StockholdersEquity
    }
  });

  return out;
}

function _ps_validShares_(row) {
  // Prefer IS.SharesDiluted, fall back to BS.SharesDiluted
  const a = _num(row.is.SharesDiluted);
  const b = _num(row.bs.SharesDiluted);
  const s = a > 0 ? a : b;
  return (s && s > 0) ? s : 0;
}

function _ps_writeOutput_(ss, rows) {
  const sh = ss.getSheetByName(PER_SHARE_CFG.OUTPUT_SHEET) || ss.insertSheet(PER_SHARE_CFG.OUTPUT_SHEET);
  sh.clear();
  sh.getRange(1,1,1,PER_SHARE_CFG.HEADERS.length).setValues([PER_SHARE_CFG.HEADERS]);
  sh.setFrozenRows(1);

  if (rows.length) sh.getRange(2,1,rows.length,PER_SHARE_CFG.HEADERS.length).setValues(rows);

  const lastRow = Math.max(2, sh.getLastRow());
  if (lastRow >= 2) {
    sh.getRange(2,3,lastRow-1,6).setNumberFormat(PER_SHARE_CFG.DEC_FMT);
  }
  for (let c=1;c<=PER_SHARE_CFG.HEADERS.length;c++) sh.autoResizeColumn(c);
}

/** ================ Small helpers ================= */

function _ps_idx_(headers, names) {
  const map = {}; names.forEach(n => map[n] = headers.indexOf(n)); return map;
}
function _ps_safeDiv_(num, den) { return (num != null && den != null && den !== 0) ? (num/den) : 0; }
function _ps_r3_(x) { return Math.round((x || 0) * 1000) / 1000; }
function _num(v) {
  const n = Number(v);
  return (isNaN(n) || v === '' || v == null) ? 0 : n;
}
function _firstNonNull_(arr) {
  for (let i=0;i<arr.length;i++) if (arr[i] != null && arr[i] !== '') return arr[i];
  return 0;
}
function _ps_log(s){ console.log(PER_SHARE_CFG.LOG_PREFIX + s); }
