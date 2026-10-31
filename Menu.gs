/**
 * PFMS - custom "PFMS" menu in the spreadsheet.
 *
 * Adds a menu next to Help when the spreadsheet is opened. It opens the app's
 * screens in a window over the spreadsheet, opens the full web app in a new
 * tab, and can create any PFMS tabs that are missing.
 *
 * The menu only appears when this Apps Script project is bound to the
 * spreadsheet (created from Extensions > Apps Script). In a standalone
 * project onOpen never runs, and the menu functions explain that.
 *
 * Uses PFMS_SCREENS, openTable_() and getPaymentTab_() from Code.gs.
 */

var PFMS_DIALOG_WIDTH = 1000;
var PFMS_DIALOG_HEIGHT = 720;
var PFMS_DIALOG_TITLES = {
  'receive-payment': 'PFMS - Receive payment',
  'make-payment': 'PFMS - Make payment',
  dashboard: 'PFMS - Dashboard',
  reports: 'PFMS - Reports',
  properties: 'PFMS - Properties',
  companies: 'PFMS - Companies',
  settings: 'PFMS - Settings'
};

/** Simple trigger: runs whenever the spreadsheet is opened. */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('PFMS')
      .addItem('Open PFMS app in a new tab', 'pfmsMenuOpenWebApp')
      .addSeparator()
      .addItem('Receive payment', 'pfmsMenuReceivePayment')
      .addItem('Make payment', 'pfmsMenuMakePayment')
      .addSeparator()
      .addItem('Dashboard', 'pfmsMenuDashboard')
      .addItem('Reports', 'pfmsMenuReports')
      .addSeparator()
      .addItem('Properties', 'pfmsMenuProperties')
      .addItem('Companies', 'pfmsMenuCompanies')
      .addItem('Settings', 'pfmsMenuSettings')
      .addSeparator()
      .addItem('Set up missing tabs', 'pfmsMenuSetUpTabs')
      .addToUi();
  } catch (e) {
    // Not opened from a spreadsheet (standalone project or web app): no menu.
  }
}

function pfmsMenuReceivePayment() { pfmsShowScreenDialog_('receive-payment'); }
function pfmsMenuMakePayment() { pfmsShowScreenDialog_('make-payment'); }
function pfmsMenuDashboard() { pfmsShowScreenDialog_('dashboard'); }
function pfmsMenuReports() { pfmsShowScreenDialog_('reports'); }
function pfmsMenuProperties() { pfmsShowScreenDialog_('properties'); }
function pfmsMenuCompanies() { pfmsShowScreenDialog_('companies'); }
function pfmsMenuSettings() { pfmsShowScreenDialog_('settings'); }

/** Opens the deployed web app in a new browser tab. */
function pfmsMenuOpenWebApp() {
  var ui = SpreadsheetApp.getUi();
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    ui.alert('PFMS', 'The web app has not been deployed yet. In Apps Script, use Deploy > New deployment > Web app first.', ui.ButtonSet.OK);
    return;
  }

  // Browsers only allow a script to open a tab right after a click, so the
  // dialog tries automatically and also shows a link to click.
  var safeUrl = url.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px;padding:8px">' +
      '<p><a id="pfms-link" href="' + safeUrl + '" target="_blank" rel="noopener" ' +
      'onclick="setTimeout(function(){google.script.host.close()},300)">Open PFMS</a></p>' +
      '<p style="color:#5f6368;font-size:12px">If a new tab didn\'t open, click the link.</p>' +
    '</div>' +
    '<script>var w = window.open(document.getElementById("pfms-link").href, "_blank");' +
    'if (w) { google.script.host.close(); }</script>'
  ).setWidth(320).setHeight(120);
  ui.showModalDialog(html, 'PFMS');
}

/** Creates any PFMS tabs that are missing and reports what it did. */
function pfmsMenuSetUpTabs() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var before = spreadsheet.getSheets().map(function (sheet) {
    return sheet.getName();
  });

  openTable_('properties', true);
  openTable_('companies', true);
  openTable_('categories', false);
  getPaymentTab_('income');
  getPaymentTab_('expense');

  var created = spreadsheet.getSheets().map(function (sheet) {
    return sheet.getName();
  }).filter(function (name) {
    return before.indexOf(name) === -1;
  });

  spreadsheet.toast(
    created.length ? 'Created: ' + created.join(', ') : 'All PFMS tabs were already there.',
    'PFMS',
    8
  );
  return created;
}

function pfmsShowScreenDialog_(screen) {
  var ui = SpreadsheetApp.getUi();
  var template = HtmlService.createTemplateFromFile('App');
  template.initialScreen = PFMS_SCREENS.indexOf(screen) !== -1 ? screen : 'home';

  var output = template.evaluate()
    .setWidth(PFMS_DIALOG_WIDTH)
    .setHeight(PFMS_DIALOG_HEIGHT);
  ui.showModalDialog(output, PFMS_DIALOG_TITLES[screen] || 'PFMS');
}
