/**
 * PFMS - Dashboard data.
 *
 * Reads the income and expense sheets and returns the totals the Dashboard
 * screen shows. Amounts are never converted between currencies: USD and CRC
 * are totalled separately.
 *
 * Finding the sheets
 *   By default every tab is scanned and a tab is used when row 1 has a Date
 *   and an Amount column. A tab counts as income when it has an "Income Type"
 *   column or its name mentions income/receive/ingreso, and as expenses when
 *   it has an "Expense Category" column or its name mentions
 *   expense/gasto/egreso/payment. Tabs that can't be classified are skipped and
 *   listed in the result so the screen can say why.
 *
 *   To name the tabs explicitly, set Script Properties (comma-separated for
 *   more than one tab):
 *     PFMS_INCOME_SHEETS   e.g.  Income
 *     PFMS_EXPENSE_SHEETS  e.g.  Expenses
 *
 * Uses getPfmsSpreadsheet_() from Uploads.gs.
 */

var PFMS_DASHBOARD_TREND_MONTHS = 6;
var PFMS_DASHBOARD_RECENT_COUNT = 10;
var PFMS_DASHBOARD_SKIP_SHEETS = ['attachments', 'properties', 'companies', 'expense categories', 'categories', 'settings'];

var PFMS_DASHBOARD_HEADERS = {
  date: ['date', 'payment date', 'fecha', 'transaction date'],
  amount: ['amount', 'monto', 'total', 'importe'],
  currency: ['currency', 'moneda'],
  property: ['property', 'property name', 'propiedad'],
  incomeType: ['income type', 'tipo de ingreso'],
  expenseCategory: ['expense category', 'categoria', 'categoría', 'category'],
  notes: ['notes', 'notas', 'description', 'descripcion', 'descripción'],
  enteredBy: ['entered by', 'ingresado por']
};

/**
 * @param {{month: string}=} options month as 'YYYY-MM'; defaults to this month.
 * @return {Object} dashboard summary
 */
function getDashboardData(options) {
  options = options || {};
  var timeZone = Session.getScriptTimeZone();
  var month = /^\d{4}-\d{2}$/.test(options.month || '')
    ? options.month
    : Utilities.formatDate(new Date(), timeZone, 'yyyy-MM');

  var source = readDashboardTransactions_(timeZone);
  var transactions = source.transactions;
  var trendMonths = getTrailingMonths_(month, PFMS_DASHBOARD_TREND_MONTHS);
  var year = month.substring(0, 4);

  var currencies = {};
  transactions.forEach(function (t) {
    currencies[t.currency] = true;
  });

  var result = {
    month: month,
    trendMonths: trendMonths,
    currencies: Object.keys(currencies).sort(),
    totals: {},
    yearToDate: {},
    trend: {},
    byProperty: {},
    byCategory: {},
    recent: [],
    sources: source.sources,
    skippedSheets: source.skipped,
    transactionCount: transactions.length
  };

  result.currencies.forEach(function (currency) {
    result.totals[currency] = { income: 0, expense: 0, count: 0 };
    result.yearToDate[currency] = { income: 0, expense: 0 };
    result.trend[currency] = trendMonths.map(function (m) {
      return { month: m, income: 0, expense: 0 };
    });
    result.byProperty[currency] = {};
    result.byCategory[currency] = {};
  });

  transactions.forEach(function (t) {
    var key = t.type === 'income' ? 'income' : 'expense';
    var c = t.currency;

    var trendIndex = trendMonths.indexOf(t.month);
    if (trendIndex !== -1) {
      result.trend[c][trendIndex][key] += t.amount;
    }

    if (t.month.substring(0, 4) === year && t.month <= month) {
      result.yearToDate[c][key] += t.amount;
    }

    if (t.month !== month) {
      return;
    }

    result.totals[c][key] += t.amount;
    result.totals[c].count++;

    var property = t.property || 'No property';
    if (!result.byProperty[c][property]) {
      result.byProperty[c][property] = { property: property, income: 0, expense: 0 };
    }
    result.byProperty[c][property][key] += t.amount;

    if (key === 'expense') {
      var category = t.category || 'Uncategorized';
      result.byCategory[c][category] = (result.byCategory[c][category] || 0) + t.amount;
    }
  });

  result.currencies.forEach(function (c) {
    result.byProperty[c] = Object.keys(result.byProperty[c])
      .map(function (name) {
        var row = result.byProperty[c][name];
        row.net = row.income - row.expense;
        return row;
      })
      .sort(function (a, b) {
        return (b.income + b.expense) - (a.income + a.expense);
      });

    result.byCategory[c] = Object.keys(result.byCategory[c])
      .map(function (name) {
        return { category: name, amount: result.byCategory[c][name] };
      })
      .sort(function (a, b) {
        return b.amount - a.amount;
      });
  });

  result.recent = transactions
    .filter(function (t) {
      return t.month <= month;
    })
    .sort(function (a, b) {
      return b.date === a.date ? b.order - a.order : (b.date < a.date ? -1 : 1);
    })
    .slice(0, PFMS_DASHBOARD_RECENT_COUNT)
    .map(function (t) {
      return {
        type: t.type,
        date: t.date,
        property: t.property,
        category: t.category,
        amount: t.amount,
        currency: t.currency,
        enteredBy: t.enteredBy
      };
    });

  return result;
}

/* ---------- reading the sheets ---------- */

function readDashboardTransactions_(timeZone) {
  var spreadsheet = getPfmsSpreadsheet_();
  var scriptProperties = PropertiesService.getScriptProperties();
  var incomeNames = splitSheetList_(scriptProperties.getProperty('PFMS_INCOME_SHEETS'));
  var expenseNames = splitSheetList_(scriptProperties.getProperty('PFMS_EXPENSE_SHEETS'));
  var configured = incomeNames.length > 0 || expenseNames.length > 0;

  var transactions = [];
  var sources = [];
  var skipped = [];
  var order = 0;

  spreadsheet.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    var lowerName = name.toLowerCase().trim();

    if (configured) {
      if (incomeNames.indexOf(lowerName) === -1 && expenseNames.indexOf(lowerName) === -1) {
        return;
      }
    } else if (PFMS_DASHBOARD_SKIP_SHEETS.indexOf(lowerName) !== -1) {
      return;
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 1 || lastColumn < 1) {
      return;
    }

    var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    var columns = mapDashboardColumns_(values[0]);

    if (columns.date === -1 || columns.amount === -1) {
      if (configured) {
        skipped.push({ sheet: name, reason: 'No Date or Amount column in row 1' });
      }
      return;
    }

    var type = configured
      ? (incomeNames.indexOf(lowerName) !== -1 ? 'income' : 'expense')
      : classifyDashboardSheet_(lowerName, values[0]);

    if (!type) {
      skipped.push({ sheet: name, reason: 'Could not tell if it holds income or expenses' });
      return;
    }

    var rowsUsed = 0;
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var date = toIsoDate_(row[columns.date], timeZone);
      var amount = toAmount_(row[columns.amount]);
      if (!date || amount === null || amount === 0) {
        continue;
      }

      var categoryColumn = type === 'income' ? columns.incomeType : columns.expenseCategory;
      if (categoryColumn === -1) {
        categoryColumn = type === 'income' ? columns.expenseCategory : columns.incomeType;
      }

      transactions.push({
        type: type,
        date: date,
        month: date.substring(0, 7),
        amount: Math.abs(amount),
        currency: normalizeCurrency_(columns.currency === -1 ? '' : row[columns.currency]),
        property: cellText_(row, columns.property),
        category: cellText_(row, categoryColumn),
        enteredBy: cellText_(row, columns.enteredBy),
        order: order++
      });
      rowsUsed++;
    }

    sources.push({ sheet: name, type: type, rows: rowsUsed });
  });

  return { transactions: transactions, sources: sources, skipped: skipped };
}

function mapDashboardColumns_(headerRow) {
  var normalized = headerRow.map(function (h) {
    return String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
  });

  var columns = {};
  Object.keys(PFMS_DASHBOARD_HEADERS).forEach(function (field) {
    var names = PFMS_DASHBOARD_HEADERS[field];
    columns[field] = -1;
    for (var i = 0; i < names.length && columns[field] === -1; i++) {
      columns[field] = normalized.indexOf(names[i]);
    }
  });
  return columns;
}

// Exact "Income Type" / "Expense Category" headers decide first; otherwise
// the tab name does.
function classifyDashboardSheet_(lowerName, headerRow) {
  var headers = headerRow.map(function (h) {
    return String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
  });
  var hasIncomeType = headers.indexOf('income type') !== -1;
  var hasExpenseCategory = headers.indexOf('expense category') !== -1;

  if (hasIncomeType && !hasExpenseCategory) {
    return 'income';
  }
  if (hasExpenseCategory && !hasIncomeType) {
    return 'expense';
  }
  if (/income|receiv|ingreso/.test(lowerName)) {
    return 'income';
  }
  if (/expense|gasto|egreso|payment|pago/.test(lowerName)) {
    return 'expense';
  }
  return '';
}

function splitSheetList_(value) {
  return String(value || '')
    .split(',')
    .map(function (name) {
      return name.toLowerCase().trim();
    })
    .filter(function (name) {
      return name.length > 0;
    });
}

function cellText_(row, column) {
  return column === -1 ? '' : String(row[column] === null || row[column] === undefined ? '' : row[column]).trim();
}

function toIsoDate_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (iso) {
    return iso[1] + '-' + pad2_(iso[2]) + '-' + pad2_(iso[3]);
  }
  // Day-first (31/10/2026), the usual order in Costa Rica.
  var dmy = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(text);
  if (dmy) {
    return dmy[3] + '-' + pad2_(dmy[2]) + '-' + pad2_(dmy[1]);
  }
  return '';
}

function toAmount_(value) {
  if (typeof value === 'number') {
    return isFinite(value) ? value : null;
  }
  var text = String(value || '').replace(/[^\d.,-]/g, '');
  if (!text) {
    return null;
  }
  // "1.234,56" -> 1234.56 ; "1,234.56" -> 1234.56
  if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else {
    text = text.replace(/,/g, '');
  }
  var number = parseFloat(text);
  return isNaN(number) ? null : number;
}

function normalizeCurrency_(value) {
  var text = String(value || '').toUpperCase().trim();
  if (!text || text === '$' || text === 'US$' || text === 'DOLLARS') {
    return 'USD';
  }
  if (text === '₡' || text === 'COLONES' || text === 'CRC₡') {
    return 'CRC';
  }
  return text;
}

function getTrailingMonths_(month, count) {
  var year = parseInt(month.substring(0, 4), 10);
  var monthIndex = parseInt(month.substring(5, 7), 10) - 1;
  var months = [];
  for (var i = count - 1; i >= 0; i--) {
    var d = new Date(year, monthIndex - i, 1);
    months.push(d.getFullYear() + '-' + pad2_(d.getMonth() + 1));
  }
  return months;
}

function pad2_(value) {
  return ('0' + value).slice(-2);
}
