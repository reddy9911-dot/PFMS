/**
 * PFMS - Settings screen.
 *
 * Reads and saves the Script Properties the other PFMS files use, so they
 * can be changed from the app instead of the Apps Script editor:
 *   PFMS_PEOPLE           names in the Entered By picker (People.gs)
 *   PFMS_INCOME_SHEETS    income tabs for Dashboard / Reports (Dashboard.gs)
 *   PFMS_EXPENSE_SHEETS   expense tabs for Dashboard / Reports
 *
 * Anyone who can open the web app can change these settings.
 */

var PFMS_MAX_PEOPLE = 10;
var PFMS_MAX_PERSON_NAME = 40;

/**
 * @return {Object} everything the Settings screen shows
 */
function getPfmsSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var spreadsheet = getPfmsSpreadsheet_();
  var incomeNames = splitSheetList_(scriptProperties.getProperty('PFMS_INCOME_SHEETS'));
  var expenseNames = splitSheetList_(scriptProperties.getProperty('PFMS_EXPENSE_SHEETS'));
  var autoDetect = incomeNames.length === 0 && expenseNames.length === 0;

  var sheets = spreadsheet.getSheets().map(function (sheet) {
    var name = sheet.getName();
    var lowerName = name.toLowerCase().trim();
    var configured = incomeNames.indexOf(lowerName) !== -1 ? 'income'
      : (expenseNames.indexOf(lowerName) !== -1 ? 'expense' : '');
    return {
      name: name,
      detected: detectSheetRole_(sheet, lowerName),
      configured: configured
    };
  });

  return {
    people: getPfmsPeople(),
    autoDetect: autoDetect,
    sheets: sheets,
    spreadsheet: { name: spreadsheet.getName(), url: spreadsheet.getUrl() },
    invoiceFolderUrl: findFolderUrl_(scriptProperties.getProperty('PFMS_INVOICE_FOLDER_ID'), PFMS_INVOICE_ROOT_FOLDER_NAME),
    reportsFolderUrl: findFolderUrl_('', PFMS_REPORTS_FOLDER_NAME),
    aiEnabled: !!getAnthropicApiKey_()
  };
}

/**
 * @param {string[]} names
 * @return {string[]} the names as saved
 */
function savePfmsPeople(names) {
  var seen = {};
  var cleaned = (names || [])
    .map(function (name) {
      return String(name || '').replace(/,/g, ' ').replace(/\s+/g, ' ').trim().substring(0, PFMS_MAX_PERSON_NAME);
    })
    .filter(function (name) {
      var key = name.toLowerCase();
      if (!name || seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });

  if (!cleaned.length) {
    throw new Error('Add at least one name.');
  }
  if (cleaned.length > PFMS_MAX_PEOPLE) {
    throw new Error('Use ' + PFMS_MAX_PEOPLE + ' names or fewer.');
  }

  PropertiesService.getScriptProperties().setProperty('PFMS_PEOPLE', cleaned.join(', '));
  return cleaned;
}

/**
 * @param {{autoDetect: boolean, income: string[], expense: string[]}} config
 * @return {Object} the refreshed settings
 */
function savePfmsSheetSettings(config) {
  config = config || {};
  var scriptProperties = PropertiesService.getScriptProperties();

  if (config.autoDetect) {
    scriptProperties.deleteProperty('PFMS_INCOME_SHEETS');
    scriptProperties.deleteProperty('PFMS_EXPENSE_SHEETS');
    return getPfmsSettings();
  }

  var existing = {};
  getPfmsSpreadsheet_().getSheets().forEach(function (sheet) {
    existing[sheet.getName().toLowerCase().trim()] = sheet.getName();
  });

  var pick = function (list) {
    return (list || []).map(function (name) {
      return existing[String(name || '').toLowerCase().trim()];
    }).filter(Boolean);
  };
  var income = pick(config.income);
  var expense = pick(config.expense).filter(function (name) {
    return income.indexOf(name) === -1;
  });

  if (!income.length && !expense.length) {
    throw new Error('Choose at least one income or expense tab, or turn automatic detection back on.');
  }
  if (income.concat(expense).some(function (name) { return name.indexOf(',') !== -1; })) {
    throw new Error('Tab names with commas can\'t be used. Rename the tab in the spreadsheet first.');
  }

  setOrDeleteProperty_(scriptProperties, 'PFMS_INCOME_SHEETS', income.join(', '));
  setOrDeleteProperty_(scriptProperties, 'PFMS_EXPENSE_SHEETS', expense.join(', '));
  return getPfmsSettings();
}

/* ---------- helpers ---------- */

// What automatic detection (Dashboard.gs) would do with this tab.
function detectSheetRole_(sheet, lowerName) {
  if (PFMS_DASHBOARD_SKIP_SHEETS.indexOf(lowerName) !== -1) {
    return 'system';
  }
  var lastColumn = sheet.getLastColumn();
  if (sheet.getLastRow() < 1 || lastColumn < 1) {
    return 'none';
  }
  var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var columns = mapDashboardColumns_(header);
  if (columns.date === -1 || columns.amount === -1) {
    return 'none';
  }
  return classifyDashboardSheet_(lowerName, header) || 'unknown';
}

function findFolderUrl_(folderId, name) {
  if (folderId) {
    try {
      var byId = DriveApp.getFolderById(folderId);
      if (!byId.isTrashed()) {
        return byId.getUrl();
      }
    } catch (e) {
      // Fall back to looking it up by name.
    }
  }
  var matches = DriveApp.getRootFolder().getFoldersByName(name);
  return matches.hasNext() ? matches.next().getUrl() : '';
}

function setOrDeleteProperty_(scriptProperties, key, value) {
  if (value) {
    scriptProperties.setProperty(key, value);
  } else {
    scriptProperties.deleteProperty(key);
  }
}
