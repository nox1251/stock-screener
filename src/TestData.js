/** Mock generator: realistic structure (annual & quarterly) for 5 PH tickers */

function getCompanyName(t) {
  const map = { ALI:'Ayala Land Inc', BDO:'BDO Unibank Inc', SM:'SM Investments Corp', JFC:'Jollibee Foods Corp', MEG:'Megaworld Corp' };
  return map[t] || `${t} Corporation`;
}

function getSectorIndustry(t) {
  const map = {
    ALI:{ sector:'Real Estate', industry:'Real Estate Development' },
    BDO:{ sector:'Financial', industry:'Banking' },
    SM:{  sector:'Consumer Cyclical', industry:'Retail - Defensive' },
    JFC:{ sector:'Consumer Cyclical', industry:'Restaurants' },
    MEG:{ sector:'Real Estate', industry:'Real Estate Development' }
  };
  return map[t] || { sector:'Diversified', industry:'Conglomerate' };
}

function generateYearlyIS(startYear, endYear, scenario) {
  const out = {};
  const base = 5e10;     // ₱50B
  const g = 1.08;        // 8% CAGR
  for (let y=startYear; y<=endYear; y++){
    const idx = y - startYear;
    const rev = base * Math.pow(g, idx);
    const opi = rev * 0.20;
    const k = `${y}-12-31`;
    out[k] = {
      date: k,
      totalRevenue: Math.round(rev),
      grossProfit: Math.round(rev*0.40),
      operatingIncome: Math.round(opi),
      netIncome: Math.round(opi*0.75)
    };
  }
  if (scenario === 'new_year') {
    const y = endYear + 1;
    const idx = y - startYear;
    const rev = base * Math.pow(g, idx);
    const k = `${y}-12-31`;
    out[k] = {
      date: k,
      totalRevenue: Math.round(rev),
      grossProfit: Math.round(rev*0.40),
      operatingIncome: Math.round(rev*0.20),
      netIncome: Math.round(rev*0.15)
    };
  }
  return out;
}

function generateQuarterlyIS(startYear, endYear, scenario) {
  const out = {};
  const baseQ = 1.2e10;   // ₱12B per quarter
  const g = 1.02;         // 2% QoQ
  let n = 0;
  for (let y=startYear; y<=endYear; y++){
    ['03-31','06-30','09-30','12-31'].forEach(sfx => {
      const rev = baseQ * Math.pow(g, n++);
      const k = `${y}-${sfx}`;
      out[k] = {
        date: k,
        totalRevenue: Math.round(rev),
        grossProfit: Math.round(rev*0.40),
        operatingIncome: Math.round(rev*0.20),
        netIncome: Math.round(rev*0.15)
      };
    });
  }
  if (scenario === 'new_quarter') {
    const rev = baseQ * Math.pow(g, n);
    const k = `${endYear+1}-03-31`;
    out[k] = {
      date: k,
      totalRevenue: Math.round(rev),
      grossProfit: Math.round(rev*0.40),
      operatingIncome: Math.round(rev*0.20),
      netIncome: Math.round(rev*0.15)
    };
  }
  if (scenario === 'revision') {
    const k = `${endYear}-09-30`;
    if (out[k]) {
      out[k].totalRevenue = Math.round(out[k].totalRevenue * 1.05);
      out[k].operatingIncome = Math.round(out[k].totalRevenue * 0.22);
    }
  }
  return out;
}

function generateBalanceSheet(startYear, endYear) {
  const yearly = {};
  const quarterly = {};
  for (let y=startYear; y<=endYear; y++){
    const assets = 2e11 * Math.pow(1.06, y-startYear);
    const k = `${y}-12-31`;
    yearly[k] = {
      date: k,
      totalAssets: Math.round(assets),
      totalLiab: Math.round(assets*0.40),
      totalStockholderEquity: Math.round(assets*0.35),
      cash: Math.round(assets*0.10)
    };
  }
  for (let y=Math.max(startYear, endYear-2); y<=endYear; y++){
    ['03-31','06-30','09-30','12-31'].forEach(sfx => {
      const assets = 2e11 * Math.pow(1.06, y-startYear);
      const k = `${y}-${sfx}`;
      quarterly[k] = {
        date: k,
        totalAssets: Math.round(assets),
        totalLiab: Math.round(assets*0.40),
        totalStockholderEquity: Math.round(assets*0.35)
      };
    });
  }
  return { yearly, quarterly };
}

function generateOutstandingShares(startYear, endYear) {
  const arr = [];
  const base = 1e10; // 10B
  for (let y=startYear; y<=endYear; y++){
    arr.push({ date: `${y}-12-31`, shares: base, sharesDiluted: Math.round(base*1.02) });
  }
  return arr.slice(-40);
}

function generateMockJSON(ticker, scenario='initial') {
  const si = getSectorIndustry(ticker);
  return {
    General: {
      Code: ticker, Type: 'Common Stock', Name: getCompanyName(ticker),
      Exchange: EXCHANGE, CurrencyCode: 'PHP', CountryISO: 'PH',
      Sector: si.sector, Industry: si.industry, FiscalYearEnd: 'December'
    },
    Financials: {
      Income_Statement: {
        yearly: generateYearlyIS(2015, 2024, scenario),
        quarterly: generateQuarterlyIS(2022, 2024, scenario)
      },
      Balance_Sheet: generateBalanceSheet(2015, 2024),
      Cash_Flow: { yearly:{}, quarterly:{} }
    },
    SharesStats: { SharesOutstanding: 1e10, SharesOutstandingDiluted: 1.02e10 },
    OutstandingShares: generateOutstandingShares(2015, 2024),
    Highlights: { MarketCapitalization: 5e11 }
  };
}

/** Full simulation: create & update 5 tickers */
function runUpsertSimulation() {
  const tickers = ['ALI','BDO','SM','JFC','MEG'];

  logInfo('Scenario 1 — Initial create (5 tickers)');
  tickers.forEach(t => upsertJSONToDrive(t, EXCHANGE, generateMockJSON(t,'initial')));

  logInfo('Scenario 2 — New quarter added');
  tickers.forEach(t => upsertJSONToDrive(t, EXCHANGE, generateMockJSON(t,'new_quarter')));

  logInfo('Scenario 3 — New annual year added');
  tickers.forEach(t => upsertJSONToDrive(t, EXCHANGE, generateMockJSON(t,'new_year')));

  logInfo('Scenario 4 — Revision on ALI');
  upsertJSONToDrive('ALI', EXCHANGE, generateMockJSON('ALI','revision'));

  const files = listCachedFiles();
  logInfo(JSON.stringify(files, null, 2));
}

/** Convenience wrapper to demo Drive ops quickly */
function testDriveOperations() {
  const mock = generateMockJSON('ALI', 'initial');
  writeJSONToDrive('ALI', EXCHANGE, mock);
  const ok = fileExists('ALI', EXCHANGE);
  logInfo(`ALI exists? ${ok}`);
  const read = readJSONFromDrive('ALI', EXCHANGE);
  logInfo(`ALI annual years: ${Object.keys(read.Financials.Income_Statement.yearly).length}`);
  listCachedFiles();
}
