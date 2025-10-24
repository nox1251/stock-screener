// sync test: added locally

/***************************************************************
 * code.gs — Stock Screener master menu & utilities
 * -------------------------------------------------------------
 * Menus:
 *  - Extract Fundamentals: Annual / Quarterly / All / Splits & Dividends
 *  - Prices: Latest (snap) / History (N trading days)
 *  - Per Share: Build Per Share (Full) / Compute CAGRs
 *  - Screener: Run Screener
 *  - Settings: Set API Key / Set Mock Prices Folder / Mock Mode ON
 *              Check Config / Authorize Drive
 *
 * Depends on other files providing:
 *  run_ExtractAnnual, run_ExtractQuarterly, run_ExtractAll, run_ExtractSplitsDividends
 *  run_ExtractPricesLatest, run_ExtractPricesHistory
 *  buildPerShare_Full, computeAnnualOpsMetrics_AndUpsert
 *  runScreener
 *  setMockPricesFolderId_NOW, setMockMode_ON
 *  checkConfig (below), _authorizeDriveOnce (Drive auth helper)
 ***************************************************************/

/* ============================== Menu ============================== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Stock Screener')
    .addSubMenu(
      ui.createMenu('Extract Fundamentals')
        .addItem('Extract — Annual (All Tickers)', 'menu_ExtractAnnual')
        .addItem('Extract — Quarterly (All Tickers)', 'menu_ExtractQuarterly')
        .addItem('Extract — All (Annual + Quarterly)', 'menu_ExtractAll')
        .addItem('Extract — Splits & Dividends', 'menu_ExtractSplitsDividends')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Prices')
        .addItem('Extract — Latest (All Tickers)', 'menu_ExtractPricesLatest')
        .addItem('Extract — History 60d (All Tickers)', 'menu_ExtractPricesHistory60')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Per Share')
        .addItem('Build Per Share (Full)', 'menu_BuildPerShareFull')
        .addItem('Compute CAGRs', 'menu_ComputeCAGRs')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Screener')
        .addItem('Run Screener', 'menu_RunScreener')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Settings')
        .addItem('Set API Key (EODHD)', 'menu_SetApiKey')
        .addItem('Set Mock Prices Folder (Drive)', 'setMockPricesFolderId_NOW')
        .addItem('Mock Mode ON (Prices from Drive)', 'setMockMode_ON')
        .addSeparator()
        .addItem('Check Config', 'menu_CheckConfig')
        .addItem('Authorize Drive', 'menu_AuthorizeDrive')
    )
    .addToUi();
}

/* ===================== Extract Fundamentals ===================== */

function menu_ExtractAnnual() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractAnnual');
    run_ExtractAnnual();
  }, 'Extracting Annual…');
}

function menu_ExtractQuarterly() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractQuarterly');
    run_ExtractQuarterly();
  }, 'Extracting Quarterly…');
}

function menu_ExtractAll() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractAll');
    run_ExtractAll();
  }, 'Extracting Annual + Quarterly…');
}

function menu_ExtractSplitsDividends() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractSplitsDividends');
    run_ExtractSplitsDividends();
  }, 'Extracting Splits & Dividends…');
}

/* ============================ Prices ============================ */

function menu_ExtractPricesLatest() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractPricesLatest');
    run_ExtractPricesLatest();
  }, 'Extracting Prices — Latest…');
}

function menu_ExtractPricesHistory60() {
  _withUiSpinner_(() => {
    _requireFunction_('run_ExtractPricesHistory');
    run_ExtractPricesHistory(60);
  }, 'Extracting Prices — History 60d…');
}

/* ============================ Per Share ============================ */

function menu_BuildPerShareFull() {
  _withUiSpinner_(() => {
    _requireFunction_('buildPerShare_Full');
    buildPerShare_Full();
  }, 'Building Per Share (Full)…');
}

function menu_ComputeCAGRs() {
  _withUiSpinner_(() => {
    _requireFunction_('computeAnnualOpsMetrics_AndUpsert');
    computeAnnualOpsMetrics_AndUpsert();
  }, 'Computing 5Y/9Y CAGRs…');
}

/* ========================= Control Center ========================= */

function menu_CC_RefreshDropdowns() {
  _withUiSpinner_(() => {
    _requireFunction_('applyControlCenterValidations');
    applyControlCenterValidations();
  }, 'Refreshing Control Center dropdowns…');
}

/* ============================= Settings ============================= */

function menu_SetApiKey() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ui.prompt('EODHD API Key', 'Enter your EODHD api_token:', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const key = (res.getResponseText() || '').trim();
    if (!key) return ui.alert('No key provided.');
    // Store at script level so collaborators share the key
    PropertiesService.getScriptProperties().setProperty('EODHD_API_KEY', key);
    ui.alert('API key saved.');
  } catch (e) {
    ui.alert('Error saving API key:\n' + (e && e.message ? e.message : e));
  }
}

function menu_CheckConfig() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActive();
    const apiKey = PropertiesService.getScriptProperties().getProperty('EODHD_API_KEY') ? 'YES' : 'NO';
    const mockMode = (typeof isMockModeOn_ === 'function' && isMockModeOn_()) ? 'ON' : 'OFF';
    const mockFolderId = PropertiesService.getUserProperties().getProperty('MOCK_PRICES_FOLDER_ID') || '(not set)';

    const sheets = [
      'Per_Share', 'Calculated_Metrics', 'Prices_Stats',
      'Prices_Latest', 'Raw_Prices',
      'Raw_Fundamentals_Annual', 'Raw_Fundamentals_Quarterly', 'Raw_Splits_Dividends'
    ];
    const lines = [
      `API Key present: ${apiKey}`,
      `Mock Mode (Prices): ${mockMode}`,
      `Mock Prices Folder ID: ${mockFolderId}`,
      '',
      'Sheets:'
    ];
    for (const name of sheets) {
      lines.push(`- ${name}: ${ss.getSheetByName(name) ? 'OK' : 'MISSING'}`);
    }
    ui.alert(lines.join('\n'));
  } catch (e) {
    ui.alert('Check Config error:\n' + (e && e.message ? e.message : e));
  }
}

function menu_AuthorizeDrive() {
  try {
    _requireFunction_('_authorizeDriveOnce');
    _authorizeDriveOnce();
    SpreadsheetApp.getUi().alert('Drive access is authorized.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Authorization failed:\n' + (e && e.message ? e.message : e));
  }
}

/* ============================ Helpers ============================ */

/**
 * Simple UX wrapper: alerts to start, runs fn, then shows timing.
 */
function _withUiSpinner_(fn, label) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(label + '\n\nClick OK to start.');
  const t0 = Date.now();
  try {
    fn();
    const ms = Date.now() - t0;
    Logger.log(`✓ ${label} — ${(ms / 1000).toFixed(1)}s`);
    ui.alert('Done.\nElapsed: ' + (ms / 1000).toFixed(1) + 's');
  } catch (e) {
    Logger.log(`✗ ${label} — ${e && e.stack ? e.stack : e}`);
    ui.alert('Error:\n' + (e && e.stack ? e.stack : e));
    throw e;
  }
}

/**
 * Ensures a function exists in the project, otherwise throws
 * a helpful error pointing to the missing file/feature.
 */
function _requireFunction_(name) {
  const fn = this[name];
  if (typeof fn !== 'function') {
    throw new Error(`Required function "${name}" not found. Make sure the corresponding file is included.`);
  }
}
