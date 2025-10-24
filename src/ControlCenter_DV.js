/*********************************************************************
 * ControlCenter_DV.gs — JSON-Pathless Control Center
 * --------------------------------------------------------------
 * Control Center columns (header on row 9):
 *   A: Section            (dropdown from SectionsRange)
 *   B: Field Name         (depends on A → Fields_<Section>)
 *   C: Period Type        (dropdown from PeriodTypeList)
 *   D: Active?            (dropdown from ActiveList)
 *   E: Notes              (free text)
 *
 * Discovers sections & fields from mock JSONs (or API if MOCK_MODE=false),
 * builds hidden lookup sheet "CC_Lookups", and applies data validation.
 *********************************************************************/

/** Public: Refresh dropdowns from real data (wired to menu) */
function applyControlCenterValidations() {
  const ss = SpreadsheetApp.getActive();
  const cc = ss.getSheetByName('Control Center');
  if (!cc) throw new Error('Control Center sheet not found.');

  // 1) Discover sections/fields
  const lookups = buildControlCenterLookupsFromJSON_();

  // 2) Rebuild lookup sheet + named ranges
  const shLook = ensureLookupSheet_();
  writeNamedLookups_(ss, shLook, lookups);

  // 3) Apply data validation on Control Center
  const HEADER_ROW = 9;
  const START_ROW  = HEADER_ROW + 1;
  const lastRow    = Math.max(cc.getLastRow(), START_ROW + 20);

  const COL_SECTION = 1; // A
  const COL_FIELD   = 2; // B
  const COL_PERIOD  = 3; // C
  const COL_ACTIVE  = 4; // D

  // Clear all existing validations in A..D
  cc.getRange(START_ROW, COL_SECTION, lastRow - START_ROW + 1, 4).clearDataValidations();

  // Section dropdown (A)
  const sectionsRange = ss.getRangeByName('SectionsRange');
  if (!sectionsRange) throw new Error('Named range "SectionsRange" not found.');
  const secRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sectionsRange, true)
    .setAllowInvalid(true) // allow blanks (user can "clear" a row)
    .build();
  cc.getRange(START_ROW, COL_SECTION, lastRow - START_ROW + 1, 1).setDataValidation(secRule);

  // Period Type dropdown (C) — Annual / Quarterly
  const periodRange = ss.getRangeByName('PeriodTypeList');
  const perRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(periodRange, true)
    .setAllowInvalid(true)
    .build();
  cc.getRange(START_ROW, COL_PERIOD, lastRow - START_ROW + 1, 1).setDataValidation(perRule);

  // Active? dropdown (D) — Yes / No
  const activeRange = ss.getRangeByName('ActiveList');
  const actRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(activeRange, true)
    .setAllowInvalid(true)
    .build();
  cc.getRange(START_ROW, COL_ACTIVE, lastRow - START_ROW + 1, 1).setDataValidation(actRule);

  // Field Name (B) depends on Section per row
  for (let r = START_ROW; r <= lastRow; r++) {
    const section = String(cc.getRange(r, COL_SECTION).getValue() || '').trim();
    _cc_applyFieldValidationForRow_(ss, cc, r, section, COL_FIELD, /*allowBlank=*/true);
  }
}

/** onEdit: when Section (col A) changes, rewire Field Name (col B) for that row */
function onEdit(e) {
  try {
    if (!e || !e.range || !e.source) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== 'Control Center') return;

    const HEADER_ROW = 9, START_ROW = HEADER_ROW + 1;
    const COL_SECTION = 1, COL_FIELD = 2;

    if (e.range.getRow() < START_ROW || e.range.getColumn() !== COL_SECTION) return;

    const ss = e.source;
    const section = String(e.range.getValue() || '').trim();
    const fieldCell = sh.getRange(e.range.getRow(), COL_FIELD);

    // Clear Field first, then apply/clear validation
    fieldCell.setValue('');
    _cc_applyFieldValidationForRow_(ss, sh, e.range.getRow(), section, COL_FIELD, /*allowBlank=*/true);
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/* ==================== Discovery & Lookups ==================== */

/** Read a handful of JSONs (or API) and collect available sections/fields. */
function buildControlCenterLookupsFromJSON_() {
  const sections = new Set();
  const fieldsBySection = {};
  const useMock = (typeof EXTRACTOR_CFG === 'object' && !!EXTRACTOR_CFG.MOCK_MODE);
  const MAX_FILES = 12;

  if (useMock) {
    const fid = _sanitizeFolderIdInput_(_getMockFolderIdFromControlCenter_());
    if (!fid) throw new Error('Control Center!B7 (Mock Folder ID) is empty.');
    const folder = DriveApp.getFolderById(fid);

    // Iterate all files, filter *.json (Drive MIME types can be inconsistent)
    const it = folder.getFiles();
    let count = 0;
    while (it.hasNext() && count < MAX_FILES) {
      const f = it.next();
      if (!String(f.getName()).toLowerCase().endsWith('.json')) continue;
      try {
        const json = JSON.parse(f.getBlob().getDataAsString('utf-8'));
        accumulateSchema_(json, sections, fieldsBySection);
        count++;
      } catch (_) {}
    }
  } else {
    // API mode: sample first few tickers
    const ss = SpreadsheetApp.getActive();
    const tSheet = ss.getSheetByName(EXTRACTOR_CFG.TICKERS_SHEET);
    if (!tSheet) throw new Error(`Tickers sheet "${EXTRACTOR_CFG.TICKERS_SHEET}" not found.`);
    const tickers = tSheet.getRange(EXTRACTOR_CFG.TICKERS_RANGE).getValues()
      .map(r => (r[0] || '').toString().trim()).filter(Boolean).slice(0, MAX_FILES);

    tickers.forEach(t => {
      try {
        const tkr = t.includes('.') ? t : `${t}.PSE`;
        const json = _fetchFromApi_(tkr);
        accumulateSchema_(json, sections, fieldsBySection);
      } catch (_) {}
    });
  }

  return { sections: Array.from(sections).sort(), fieldsBySection };
}

/** Pull keys from Financials.<Section>.yearly/quarterly arrays. */
function accumulateSchema_(json, sections, fieldsBySection) {
  const fin = json && json.Financials;
  if (!fin) return;

  ['Income_Statement','Balance_Sheet','Cash_Flow'].forEach(sec => {
    const node = fin[sec];
    if (!node) return;

    sections.add(sec);
    const set = (fieldsBySection[sec] = fieldsBySection[sec] || new Set());
    ['yearly','quarterly'].forEach(period => {
      const arr = node[period];
      if (Array.isArray(arr)) {
        arr.forEach(obj => {
          if (obj && typeof obj === 'object') {
            Object.keys(obj).forEach(k => {
              if (k === 'date' || k === 'fiscalYear' || k === 'fiscalQuarter') return;
              set.add(k); // TitleCase field names
            });
          }
        });
      }
    });
  });
}

/* ==================== Lookups sheet & named ranges ==================== */

function ensureLookupSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('CC_Lookups');
  if (!sh) sh = ss.insertSheet('CC_Lookups');
  sh.clear();

  // Nuke old named ranges to avoid stale references
  ss.getNamedRanges().forEach(nr => {
    const n = nr.getName();
    if (n === 'SectionsRange' || n === 'PeriodTypeList' || n === 'ActiveList' || n.startsWith('Fields_')) {
      nr.remove();
    }
  });
  return sh;
}

function writeNamedLookups_(ss, sh, data) {
  // Column 1: Sections
  sh.getRange(1,1).setValue('Sections');
  if (data.sections.length) sh.getRange(2,1,data.sections.length,1).setValues(data.sections.map(v=>[v]));
  ss.setNamedRange('SectionsRange', sh.getRange(2,1,Math.max(1,data.sections.length),1));

  // Column 2: PeriodTypeList (Annual, Quarterly)
  const period = [['Annual'],['Quarterly']];
  sh.getRange(1,2).setValue('PeriodType');
  sh.getRange(2,2,period.length,1).setValues(period);
  ss.setNamedRange('PeriodTypeList', sh.getRange(2,2,period.length,1));

  // Column 3: ActiveList (Yes, No)
  const active = [['Yes'],['No']];
  sh.getRange(1,3).setValue('Active');
  sh.getRange(2,3,active.length,1).setValues(active);
  ss.setNamedRange('ActiveList', sh.getRange(2,3,active.length,1));

  // Per-section field lists starting at column 5 and spaced by 2
  let col = 5;
  data.sections.forEach(sec => {
    const fields = Array.from(data.fieldsBySection[sec] || new Set()).sort();
    sh.getRange(1,col).setValue(sec);
    if (fields.length) sh.getRange(2,col,fields.length,1).setValues(fields.map(v=>[v]));
    ss.setNamedRange(`Fields_${sec}`, sh.getRange(2,col,Math.max(1,fields.length),1));
    col += 2;
  });

  sh.hideSheet();
}

/* ==================== Per-row field validation ==================== */

function _cc_applyFieldValidationForRow_(ss, sh, row, section, COL_FIELD, allowBlank) {
  const cell = sh.getRange(row, COL_FIELD);
  cell.clearDataValidations();
  if (!section) return; // cleared section → leave field free

  const fieldsRange = ss.getRangeByName(`Fields_${section}`);
  if (fieldsRange) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(fieldsRange, true)
      .setAllowInvalid(allowBlank === true)
      .build();
    cell.setDataValidation(rule);
  }
}

/* ==================== Shared small helpers ==================== */

function _sanitizeFolderIdInput_(input) {
  const s = (input || '').trim();
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}
function _getMockFolderIdFromControlCenter_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Control Center');
    if (!sh) return '';
    return String(sh.getRange('B7').getValue() || '').trim(); // B7 holds the folder ID
  } catch (e) { return ''; }
}
