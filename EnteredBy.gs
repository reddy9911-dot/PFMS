/**
 * PFMS - "Entered By" on the payment rows.
 *
 * saveReceivePayment / saveMakePayment (in Code.gs) write the payment row
 * but don't store who entered it. Right after a save, the app calls
 * recordEnteredBy(), which finds the row that was just written and puts the
 * name in an "Entered By" column, adding that column to the tab's header
 * the first time.
 *
 * The row is the bottom-most one in an income (or expense) tab whose Date,
 * Amount, Property and Currency match the payment and whose Entered By cell
 * is still empty. If Code.gs is later changed to write payment.enteredBy
 * itself, this finds nothing empty to fill and does nothing.
 *
 * Uses listPaymentSheets_() and helpers from Dashboard.gs.
 */

var PFMS_ENTERED_BY_HEADER = 'Entered By';

/**
 * @param {string} paymentType 'Income' or 'Expense'
 * @param {Object} payment the object passed to saveReceivePayment / saveMakePayment
 * @return {{updated: boolean, sheet: string=, row: number=, reason: string=}}
 */
function recordEnteredBy(paymentType, payment) {
  payment = payment || {};
  var name = String(payment.enteredBy || '').trim().substring(0, 80);
  if (!name) {
    return { updated: false, reason: 'No name given' };
  }

  var timeZone = Session.getScriptTimeZone();
  var type = String(paymentType || '').toLowerCase() === 'income' ? 'income' : 'expense';
  var date = toIsoDate_(payment.date, timeZone);
  var amount = toAmount_(payment.amount);
  var currency = normalizeCurrency_(payment.currency);
  var property = String(payment.property || '').trim();

  if (!date || amount === null) {
    return { updated: false, reason: 'Payment has no date or amount' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheets = listPaymentSheets_().sheets.filter(function (entry) {
      return entry.type === type;
    });

    for (var s = 0; s < sheets.length; s++) {
      var entry = sheets[s];
      var values = entry.values;
      var columns = entry.columns;

      for (var r = values.length - 1; r >= 1; r--) {
        var row = values[r];
        if (columns.enteredBy !== -1 && cellText_(row, columns.enteredBy)) {
          continue;
        }
        if (toIsoDate_(row[columns.date], timeZone) !== date) {
          continue;
        }
        var rowAmount = toAmount_(row[columns.amount]);
        if (rowAmount === null || Math.abs(Math.abs(rowAmount) - Math.abs(amount)) > 0.005) {
          continue;
        }
        if (columns.property !== -1 && cellText_(row, columns.property) !== property) {
          continue;
        }
        if (columns.currency !== -1 && normalizeCurrency_(row[columns.currency]) !== currency) {
          continue;
        }

        var column = ensureEnteredByColumn_(entry);
        entry.sheet.getRange(r + 1, column + 1).setValue(name);
        return { updated: true, sheet: entry.name, row: r + 1 };
      }
    }
  } finally {
    lock.releaseLock();
  }

  return { updated: false, reason: 'The saved row was not found' };
}

// Returns the zero-based Entered By column, adding the header after the
// last used header cell if the tab doesn't have one yet.
function ensureEnteredByColumn_(entry) {
  if (entry.columns.enteredBy !== -1) {
    return entry.columns.enteredBy;
  }

  var header = entry.values[0];
  var lastUsed = -1;
  for (var c = 0; c < header.length; c++) {
    if (String(header[c] || '').trim()) {
      lastUsed = c;
    }
  }

  var column = lastUsed + 1;
  var hasDataBelow = entry.values.some(function (row, index) {
    return index > 0 && column < row.length && String(row[column] === null ? '' : row[column]).trim() !== '';
  });
  if (hasDataBelow) {
    // A blank header over data: don't write over it, use a new column.
    column = header.length;
  }

  var headerCell = entry.sheet.getRange(1, column + 1);
  headerCell.setValue(PFMS_ENTERED_BY_HEADER);
  if (column > 0) {
    // Match the look of the existing header.
    entry.sheet.getRange(1, column).copyTo(headerCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  entry.columns.enteredBy = column;
  return column;
}
