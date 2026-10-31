/**
 * PFMS - Reports.
 *
 * Lists payments for a date range with optional property / type filters,
 * totals per currency (USD and CRC are never added together), totals per
 * category, and links to invoices from the Attachments sheet.
 *
 * Uses readDashboardTransactions_() and helpers from Dashboard.gs, and
 * getPfmsSpreadsheet_() / getOrCreateChildFolder_() from Uploads.gs.
 */

var PFMS_REPORT_MAX_ROWS = 5000;
var PFMS_REPORTS_FOLDER_NAME = 'PFMS Reports';

/**
 * @param {{from: string, to: string, property: string, type: string}} filters
 *   from/to as 'YYYY-MM-DD' (inclusive); property '' for all;
 *   type 'income', 'expense' or '' for both.
 * @return {Object} report
 */
function getReportData(filters) {
  var timeZone = Session.getScriptTimeZone();
  var f = normalizeReportFilters_(filters, timeZone);
  var source = readDashboardTransactions_(timeZone);
  var attachments = readAttachmentIndex_(timeZone);

  var allProperties = {};
  source.transactions.forEach(function (t) {
    if (t.property) {
      allProperties[t.property] = true;
    }
  });

  var matches = source.transactions.filter(function (t) {
    return t.date >= f.from && t.date <= f.to &&
      (!f.property || t.property === f.property) &&
      (!f.type || t.type === f.type);
  });

  matches.sort(function (a, b) {
    return a.date === b.date ? a.order - b.order : (a.date < b.date ? -1 : 1);
  });

  var totals = {};
  var categories = {};
  matches.forEach(function (t) {
    var c = t.currency;
    if (!totals[c]) {
      totals[c] = { currency: c, income: 0, expense: 0, count: 0 };
      categories[c] = {};
    }
    totals[c][t.type] += t.amount;
    totals[c].count++;

    var categoryKey = t.type + '|' + (t.category || 'Uncategorized');
    if (!categories[c][categoryKey]) {
      categories[c][categoryKey] = { type: t.type, category: t.category || 'Uncategorized', amount: 0, count: 0 };
    }
    categories[c][categoryKey].amount += t.amount;
    categories[c][categoryKey].count++;
  });

  var currencyList = Object.keys(totals).sort();

  return {
    filters: f,
    properties: Object.keys(allProperties).sort(),
    totals: currencyList.map(function (c) {
      var t = totals[c];
      t.net = t.income - t.expense;
      return t;
    }),
    categories: currencyList.map(function (c) {
      return {
        currency: c,
        rows: Object.keys(categories[c]).map(function (k) {
          return categories[c][k];
        }).sort(function (a, b) {
          return a.type === b.type ? b.amount - a.amount : (a.type === 'income' ? -1 : 1);
        })
      };
    }),
    rowCount: matches.length,
    truncated: matches.length > PFMS_REPORT_MAX_ROWS,
    rows: matches.slice(0, PFMS_REPORT_MAX_ROWS).map(function (t) {
      return {
        type: t.type,
        date: t.date,
        property: t.property,
        category: t.category,
        amount: t.amount,
        currency: t.currency,
        enteredBy: t.enteredBy,
        notes: t.notes,
        attachments: attachments[attachmentKey_(t.type, t.property, t.date, t.amount, t.currency)] || []
      };
    }),
    sources: source.sources,
    skippedSheets: source.skipped
  };
}

/**
 * Saves the report as a new Google Sheet in Drive > "PFMS Reports".
 * @return {{url: string, name: string}}
 */
function exportReportToSheet(filters) {
  var report = getReportData(filters);
  var f = report.filters;
  var name = 'PFMS Report ' + f.from + ' to ' + f.to +
    (f.property ? ' - ' + f.property : '') +
    (f.type ? ' - ' + (f.type === 'income' ? 'Income' : 'Expenses') : '');

  var spreadsheet = SpreadsheetApp.create(name);
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName('Payments');

  var header = ['Date', 'Type', 'Property', 'Category', 'Amount', 'Currency', 'Entered By', 'Notes', 'Invoices'];
  var rows = report.rows.map(function (r) {
    return [
      r.date,
      r.type === 'income' ? 'Income' : 'Expense',
      r.property,
      r.category,
      r.type === 'income' ? r.amount : -r.amount,
      r.currency,
      r.enteredBy,
      r.notes,
      r.attachments.map(function (a) {
        return a.url;
      }).join('\n')
    ];
  });

  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0.00');
  }
  sheet.setFrozenRows(1);

  var summary = spreadsheet.insertSheet('Summary');
  var summaryRows = [['Report', name], ['From', f.from], ['To', f.to],
    ['Property', f.property || 'All'], ['Type', f.type ? (f.type === 'income' ? 'Income' : 'Expenses') : 'All'], ['', '']];
  summaryRows.push(['Currency', 'Income', 'Expenses', 'Net', 'Payments']);
  report.totals.forEach(function (t) {
    summaryRows.push([t.currency, t.income, t.expense, t.net, t.count]);
  });
  var width = 5;
  summary.getRange(1, 1, summaryRows.length, width).setValues(summaryRows.map(function (r) {
    while (r.length < width) {
      r.push('');
    }
    return r;
  }));
  summary.getRange(7, 1, 1, width).setFontWeight('bold');

  var file = DriveApp.getFileById(spreadsheet.getId());
  var folder = getOrCreateChildFolder_(DriveApp.getRootFolder(), PFMS_REPORTS_FOLDER_NAME);
  file.moveTo(folder);

  return { url: spreadsheet.getUrl(), name: name };
}

/* ---------- helpers ---------- */

function normalizeReportFilters_(filters, timeZone) {
  filters = filters || {};
  var today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  var isoDate = /^\d{4}-\d{2}-\d{2}$/;
  var from = isoDate.test(filters.from || '') ? filters.from : today.substring(0, 8) + '01';
  var to = isoDate.test(filters.to || '') ? filters.to : today;
  if (from > to) {
    var swap = from;
    from = to;
    to = swap;
  }
  var type = filters.type === 'income' || filters.type === 'expense' ? filters.type : '';
  return { from: from, to: to, property: String(filters.property || ''), type: type };
}

function attachmentKey_(type, property, date, amount, currency) {
  return [type, String(property || '').trim(), date, Number(amount).toFixed(2), currency].join('|');
}

// Index of Attachments-sheet rows by payment, so each report row can link
// to its invoices. Returns {} when there is no Attachments sheet yet.
function readAttachmentIndex_(timeZone) {
  var index = {};
  var sheet = getPfmsSpreadsheet_().getSheetByName(PFMS_ATTACHMENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return index;
  }

  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var header = values[0].map(function (h) {
    return String(h || '').toLowerCase().trim();
  });
  var col = function (name) {
    return header.indexOf(name.toLowerCase());
  };
  var c = {
    type: col('Payment Type'),
    property: col('Property'),
    date: col('Payment Date'),
    amount: col('Amount'),
    currency: col('Currency'),
    name: col('File Name'),
    url: col('File Link')
  };
  if (c.type === -1 || c.date === -1 || c.amount === -1 || c.url === -1) {
    return index;
  }

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var type = String(row[c.type]).toLowerCase() === 'income' ? 'income' : 'expense';
    var date = toIsoDate_(row[c.date], timeZone);
    var amount = toAmount_(row[c.amount]);
    var url = String(row[c.url] || '');
    if (!date || amount === null || !url) {
      continue;
    }
    var key = attachmentKey_(
      type,
      c.property === -1 ? '' : row[c.property],
      date,
      Math.abs(amount),
      normalizeCurrency_(c.currency === -1 ? '' : row[c.currency])
    );
    if (!index[key]) {
      index[key] = [];
    }
    index[key].push({ name: c.name === -1 ? 'Invoice' : String(row[c.name] || 'Invoice'), url: url });
  }
  return index;
}
