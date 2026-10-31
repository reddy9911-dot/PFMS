/**
 * PFMS - Invoice / factura attachments.
 *
 * Add this file to the Apps Script project next to the existing code.
 * It is self-contained: it does not change saveReceivePayment or
 * saveMakePayment.
 *
 * Where things go:
 *   - Files:  Google Drive > "PFMS Invoices" / <year> / <property> /
 *             (owned by whoever the web app runs as, usually you)
 *   - Links:  an "Attachments" tab in the PFMS spreadsheet, one row per
 *             file, with the payment details and who entered it.
 *
 * Optional Script Properties (Project Settings > Script Properties):
 *   PFMS_SPREADSHEET_ID     - needed only if this script is NOT bound to
 *                             the PFMS spreadsheet.
 *   PFMS_INVOICE_FOLDER_ID  - use an existing Drive folder instead of
 *                             creating "PFMS Invoices" in My Drive.
 *                             Set automatically the first time.
 *
 * The first upload asks for Google Drive permission.
 */

var PFMS_INVOICE_ROOT_FOLDER_NAME = 'PFMS Invoices';
var PFMS_ATTACHMENTS_SHEET_NAME = 'Attachments';
var PFMS_MAX_UPLOAD_BYTES = 20 * 1024 * 1024;
var PFMS_ATTACHMENT_HEADERS = [
  'Uploaded At',
  'Payment Type',
  'Property',
  'Category / Income Type',
  'Payment Date',
  'Amount',
  'Currency',
  'File Name',
  'File Link',
  'File ID',
  'Entered By'
];

/**
 * Saves one uploaded file to Drive.
 *
 * @param {{name: string, mimeType: string, base64: string}} upload
 * @param {{paymentType: string, property: string, date: string, enteredBy: string}} meta
 * @return {{fileId: string, name: string, url: string, mimeType: string, size: number}}
 */
function uploadPaymentAttachment(upload, meta) {
  if (!upload || !upload.base64) {
    throw new Error('No file was received.');
  }
  meta = meta || {};

  var bytes = Utilities.base64Decode(upload.base64);
  if (!bytes.length) {
    throw new Error('The file is empty.');
  }
  if (bytes.length > PFMS_MAX_UPLOAD_BYTES) {
    throw new Error('The file is larger than 20 MB.');
  }

  var originalName = sanitizeDriveName_(upload.name) || 'invoice';
  var paymentDate = normalizePaymentDate_(meta.date);
  var paymentType = sanitizeDriveName_(meta.paymentType) || 'Payment';
  var fileName = [paymentDate, paymentType, originalName].join(' - ');

  var blob = Utilities.newBlob(bytes, upload.mimeType || 'application/octet-stream', fileName);
  var folder = getInvoiceFolder_(paymentDate.substring(0, 4), meta.property);
  var file = folder.createFile(blob);

  var enteredBy = resolveEnteredBy_(meta.enteredBy);
  file.setDescription(
    'PFMS ' + paymentType + ' | ' + (meta.property || '') + ' | ' + paymentDate +
    (enteredBy ? ' | entered by ' + enteredBy : '')
  );

  return {
    fileId: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    mimeType: file.getMimeType(),
    size: bytes.length
  };
}

/**
 * Records saved files in the Attachments sheet. Called after the payment
 * itself has been saved.
 *
 * @param {string} paymentType 'Income' or 'Expense'
 * @param {Object} payment the same object passed to saveReceivePayment / saveMakePayment
 * @param {Array<{fileId: string, name: string, url: string}>} attachments
 * @return {{logged: number}}
 */
function logPaymentAttachments(paymentType, payment, attachments) {
  if (!attachments || !attachments.length) {
    return { logged: 0 };
  }
  payment = payment || {};

  var sheet = getAttachmentsSheet_();
  var now = new Date();
  var enteredBy = resolveEnteredBy_(payment.enteredBy);
  var category = payment.expenseCategory || payment.incomeType || '';

  var rows = attachments.map(function (attachment) {
    return [
      now,
      paymentType || '',
      payment.property || '',
      category,
      payment.date || '',
      payment.amount === undefined ? '' : payment.amount,
      payment.currency || '',
      attachment.name || '',
      attachment.url || '',
      attachment.fileId || '',
      enteredBy
    ];
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PFMS_ATTACHMENT_HEADERS.length)
      .setValues(rows);
  } finally {
    lock.releaseLock();
  }

  return { logged: rows.length };
}

/* ---------- helpers ---------- */

function getPfmsSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('PFMS_SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error(
      'PFMS spreadsheet not found. Add a Script Property named PFMS_SPREADSHEET_ID ' +
      'with the spreadsheet ID.'
    );
  }
  return active;
}

function getAttachmentsSheet_() {
  var spreadsheet = getPfmsSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(PFMS_ATTACHMENTS_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(PFMS_ATTACHMENTS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, PFMS_ATTACHMENT_HEADERS.length)
      .setValues([PFMS_ATTACHMENT_HEADERS])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getInvoiceFolder_(year, propertyName) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var root = getInvoiceRootFolder_();
    var yearFolder = getOrCreateChildFolder_(root, year || String(new Date().getFullYear()));
    return getOrCreateChildFolder_(yearFolder, sanitizeDriveName_(propertyName) || 'No Property');
  } finally {
    lock.releaseLock();
  }
}

function getInvoiceRootFolder_() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var folderId = scriptProperties.getProperty('PFMS_INVOICE_FOLDER_ID');

  if (folderId) {
    try {
      var existing = DriveApp.getFolderById(folderId);
      if (!existing.isTrashed()) {
        return existing;
      }
    } catch (e) {
      // Folder was deleted or is no longer accessible; fall through and recreate.
    }
  }

  var matches = DriveApp.getRootFolder().getFoldersByName(PFMS_INVOICE_ROOT_FOLDER_NAME);
  var folder = matches.hasNext()
    ? matches.next()
    : DriveApp.getRootFolder().createFolder(PFMS_INVOICE_ROOT_FOLDER_NAME);

  scriptProperties.setProperty('PFMS_INVOICE_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateChildFolder_(parent, name) {
  var matches = parent.getFoldersByName(name);
  return matches.hasNext() ? matches.next() : parent.createFolder(name);
}

function sanitizeDriveName_(value) {
  return String(value || '')
    .replace(/[\/\\:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 120);
}

function normalizePaymentDate_(value) {
  var text = String(value || '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Prefer the name picked on the form; fall back to the Google account
// email (often blank for other people on regular Gmail accounts).
function resolveEnteredBy_(pickedName) {
  var name = String(pickedName || '').trim().substring(0, 80);
  return name || getCurrentUserEmail_();
}

function getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}
