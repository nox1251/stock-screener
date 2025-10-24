function refreshDataDictionary() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const dict = ss.getSheetByName('Dictionary') || ss.insertSheet('Dictionary');
  dict.clear();
  dict.getRange(1,1,1,4).setValues([['Section','PeriodType','FieldName','JsonKeyPath']]);

  const tuples = new Map(); // key: section|period|field -> {section,period,field,path}
  const folder = getOrCreateCacheFolder();
  const files = folder.getFiles();

  let scannedFiles = 0;
  while (files.hasNext()) {
    const file = files.next();
    scannedFiles++;
    let json;
    try {
      json = JSON.parse(file.getBlob().getDataAsString());
    } catch (e) { continue; }

    const fin = json.Financials || {};
    [
      {sec:'Income_Statement', obj: fin.Income_Statement},
      {sec:'Balance_Sheet',    obj: fin.Balance_Sheet},
      {sec:'Cash_Flow',        obj: fin.Cash_Flow}
    ].forEach(({sec, obj}) => {
      if (!obj) return;
      if (obj.yearly && typeof obj.yearly === 'object') {
        collectKeysFromPeriodObj_(obj.yearly).forEach(field => {
          tuples.set(`${sec}|Annual|${field}`, {
            section: sec, period: 'Annual', field,
            path: `Financials.${sec}.yearly.*.${field}`
          });
        });
      }
      if (obj.quarterly && typeof obj.quarterly === 'object') {
        collectKeysFromPeriodObj_(obj.quarterly).forEach(field => {
          tuples.set(`${sec}|Quarterly|${field}`, {
            section: sec, period: 'Quarterly', field,
            path: `Financials.${sec}.quarterly.*.${field}`
          });
        });
      }
    });

    if (Array.isArray(json.OutstandingShares)) {
      collectKeysFromArray_(json.OutstandingShares).forEach(field => {
        tuples.set(`OutstandingShares|Annual|${field}`, {
          section: 'OutstandingShares', period: 'Annual', field,
          path: `OutstandingShares.*.${field}`
        });
      });
    }
    if (json.Highlights && typeof json.Highlights === 'object') {
      Object.keys(json.Highlights).forEach(field => {
        tuples.set(`Highlights|Both|${field}`, {
          section: 'Highlights', period: 'Both', field,
          path: `Highlights.${field}`
        });
      });
    }
    if (json.SharesStats && typeof json.SharesStats === 'object') {
      Object.keys(json.SharesStats).forEach(field => {
        tuples.set(`SharesStats|Both|${field}`, {
          section: 'SharesStats', period: 'Both', field,
          path: `SharesStats.${field}`
        });
      });
    }
  }

  const rows = Array.from(tuples.values())
    .sort((a,b) => (a.section+a.period+a.field).localeCompare(b.section+b.period+b.field))
    .map(x => [x.section, x.period, x.field, x.path]);

  if (rows.length) {
    dict.getRange(2,1,rows.length,4).setValues(rows);
  }

// Named ranges for dropdowns (no SheetTargetList anymore)
makeListNamedRange_(dict, 'SectionList',     1, uniqueInColumn_(rows, 0));
makeListNamedRange_(dict, 'PeriodTypeList',  1, ['Annual','Quarterly']);  // no "Both"
makeListNamedRange_(dict, 'ActiveList',      1, ['Yes','No']);

// Hide Dictionary (optional)
dict.hideSheet();

  // ------- NEW: summary popup / toast -------
  if (rows.length === 0) {
    ui.alert(
      'Refresh Data Dictionary',
      scannedFiles === 0
        ? 'No cached JSON files were found in Drive.\n\nTip: Run “Simulate: Upsert Suite” first (or fetch real data), then run “Refresh Data Dictionary” again.'
        : 'No fields were discovered in cached JSON files.\n\nCheck that JSON structure contains Financials/Highlights/Shares sections.',
      ui.ButtonSet.OK
    );
  } else {
    // summarize
    const sections = uniqueInColumn_(rows, 0);
    const totalFields = rows.length;
    const msg = `✅ Dictionary updated\n• Files scanned: ${scannedFiles}\n• Sections found: ${sections.length}\n• Fields discovered: ${totalFields}\n\nSections: ${sections.join(', ')}`;
    ss.toast(msg, 'Stock Screener', 8); // non-blocking toast (8 seconds)

    // Optional: also show a one-click alert (blocking) on first runs:
    // ui.alert('Refresh Data Dictionary', msg, ui.ButtonSet.OK);
  }
}

/** ---------- Helpers required by refreshDataDictionary() ---------- **/

// periodObj looks like: { "2024-12-31": { totalRevenue: ..., operatingIncome: ... }, ... }
function collectKeysFromPeriodObj_(periodObj) {
  const keySet = new Set();
  Object.keys(periodObj || {}).forEach(dateKey => {
    const row = periodObj[dateKey];
    if (row && typeof row === 'object') {
      Object.keys(row).forEach(k => {
        if (k !== 'date' && k !== 'dateFormatted') keySet.add(k);
      });
    }
  });
  return keySet;
}

// arr looks like: [ { date: "2024-12-31", shares: 123, sharesDiluted: 125 }, ... ]
function collectKeysFromArray_(arr) {
  const keySet = new Set();
  (arr || []).forEach(o => {
    if (o && typeof o === 'object') {
      Object.keys(o).forEach(k => {
        if (k !== 'date' && k !== 'dateFormatted') keySet.add(k);
      });
    }
  });
  return keySet;
}

// Deduplicates a specific column from the 2D rows array (0-based index)
function uniqueInColumn_(rows, idx) {
  return Array.from(new Set((rows || []).map(r => r[idx]))).filter(Boolean).sort();
}

// Writes a one-row horizontal list to the Dictionary sheet and assigns a named range to it
function makeListNamedRange_(sheet, name, padRows, list) {
  const lastCol = Math.max(4, sheet.getLastColumn()) + 2; // place lists to the right of main table
  const r = sheet.getRange(1 + padRows, lastCol, 1, Math.max(1, list.length));
  r.clearContent();
  if (list && list.length) r.setValues([list]);
  SpreadsheetApp.getActive().setNamedRange(name, r);
}

