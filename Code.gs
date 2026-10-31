/**
 * PFMS - core server functions: opening the app, properties, companies,
 * expense categories, and saving payments.
 *
 * Works with the spreadsheet's existing tabs instead of assuming a layout:
 * each tab is found by its name (English or Spanish), each field by its
 * header in row 1 (accents and case ignored), and values are written only
 * into matching columns. A missing column is added after the last header;
 * a missing tab is created with standard headers.
 *
 * Every value returned to the browser is a string, number or boolean.
 * google.script.run returns null for objects containing Date values, so
 * dates are sent as 'YYYY-MM-DD'.
 *
 * Uses getPfmsSpreadsheet_() from Uploads.gs and listPaymentSheets_() /
 * toIsoDate_() / toAmount_() / normalizeCurrency_() from Dashboard.gs.
 */

var PFMS_SCREENS = ['home', 'receive-payment', 'make-payment', 'properties', 'companies', 'dashboard', 'reports', 'settings'];

var PFMS_TABLES = {
  properties: {
    tabNames: ['Properties', 'Propiedades', 'Property', 'Propiedad'],
    idPrefix: 'PROP-',
    defaultHeaders: ['Property ID', 'Name', 'Company ID', 'Company', 'Address', 'Purchase Date', 'Purchase Price',
      'Currency', 'Status', 'Notes', 'Created At', 'Updated At'],
    fields: {
      id: ['property id', 'propertyid', 'id', 'id propiedad'],
      name: ['name', 'property name', 'property', 'nombre', 'propiedad'],
      companyId: ['company id', 'companyid', 'id compania', 'id empresa'],
      companyName: ['company', 'company name', 'compania', 'empresa', 'sociedad'],
      cedulaJuridica: ['cedula juridica', 'cedula'],
      address: ['address', 'direccion', 'ubicacion'],
      purchaseDate: ['purchase date', 'fecha de compra', 'fecha compra'],
      purchasePrice: ['purchase price', 'precio de compra', 'precio compra', 'price', 'precio'],
      currency: ['currency', 'moneda'],
      status: ['status', 'estado'],
      notes: ['notes', 'notas', 'observaciones'],
      createdAt: ['created at', 'created', 'fecha creacion'],
      updatedAt: ['updated at', 'updated', 'fecha actualizacion']
    }
  },
  companies: {
    tabNames: ['Companies', 'Companias', 'Compañías', 'Empresas', 'Sociedades', 'Company'],
    idPrefix: 'COMP-',
    defaultHeaders: ['Company ID', 'Name', 'Cedula Juridica', 'Address', 'Phone', 'Email', 'Bank', 'Status', 'Notes',
      'Created At', 'Updated At'],
    fields: {
      id: ['company id', 'companyid', 'id', 'id compania', 'id empresa'],
      name: ['name', 'company name', 'company', 'nombre', 'compania', 'empresa', 'sociedad'],
      cedulaJuridica: ['cedula juridica', 'cedula', 'tax id', 'identificacion'],
      address: ['address', 'direccion'],
      phone: ['phone', 'telefono', 'tel'],
      email: ['email', 'e-mail', 'correo', 'correo electronico'],
      bank: ['bank', 'banco', 'bank account', 'cuenta bancaria', 'cuenta'],
      status: ['status', 'estado'],
      notes: ['notes', 'notas', 'observaciones'],
      createdAt: ['created at', 'created', 'fecha creacion'],
      updatedAt: ['updated at', 'updated', 'fecha actualizacion']
    }
  },
  categories: {
    tabNames: ['Expense Categories', 'Categories', 'Categorias', 'Categorías', 'Categorias de Gastos', 'Categorías de Gastos'],
    defaultHeaders: ['Category', 'Status'],
    defaultRows: ['Repairs', 'Maintenance', 'Utilities', 'Property Tax', 'HOA / Condo Fees', 'Insurance', 'Cleaning',
      'Management Fees', 'Mortgage', 'Other'],
    fields: {
      name: ['category', 'expense category', 'name', 'categoria', 'nombre'],
      status: ['status', 'estado', 'active', 'activo']
    }
  }
};

var PFMS_PAYMENT_TABS = {
  income: {
    name: 'Income',
    headers: ['Date', 'Property', 'Income Type', 'Currency', 'Amount', 'Notes', 'Entered By', 'Created At']
  },
  expense: {
    name: 'Expenses',
    headers: ['Date', 'Property', 'Expense Category', 'Currency', 'Amount', 'Notes', 'Entered By', 'Created At']
  }
};

/* ---------- web app ---------- */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var screen = String(params.screen || params.page || 'home');

  var template = HtmlService.createTemplateFromFile('App');
  template.initialScreen = PFMS_SCREENS.indexOf(screen) !== -1 ? screen : 'home';

  return template.evaluate()
    .setTitle('PFMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- properties ---------- */

/** Active properties, for the payment forms. */
function getProperties() {
  return getAllProperties().filter(function (p) {
    return !isArchivedStatus_(p.status);
  }).map(function (p) {
    return { propertyId: p.propertyId, name: p.name };
  });
}

/** All properties, for the Properties screen. */
function getAllProperties() {
  var table = openTable_('properties', true);
  var companies = companyLookup_();

  return table.rows.map(function (row) {
    var companyId = cellString_(row.values, table.columns.companyId);
    var companyName = cellString_(row.values, table.columns.companyName);
    var company = companies.byId[companyId] || companies.byName[companyName.toLowerCase()] || null;

    return {
      propertyId: cellString_(row.values, table.columns.id),
      name: cellString_(row.values, table.columns.name),
      companyId: company ? company.companyId : companyId,
      companyName: company ? company.name : companyName,
      cedulaJuridica: (company && company.cedulaJuridica) || cellString_(row.values, table.columns.cedulaJuridica),
      address: cellString_(row.values, table.columns.address),
      purchaseDate: cellDate_(row.values, table.columns.purchaseDate),
      purchasePrice: cellNumberOrString_(row.values, table.columns.purchasePrice),
      currency: normalizeCurrency_(cellString_(row.values, table.columns.currency)),
      status: normalizeStatus_(cellString_(row.values, table.columns.status)),
      notes: cellString_(row.values, table.columns.notes)
    };
  }).filter(function (p) {
    return p.name;
  });
}

/**
 * @param {Object} data fields from the property form
 * @return {{success: boolean, errors: string[]=, propertyId: string=}}
 */
function saveProperty(data) {
  data = data || {};
  var name = cleanText_(data.name, 120);
  if (!name) {
    return { success: false, errors: ['Property name is required.'] };
  }

  var purchasePrice = String(data.purchasePrice || '').trim();
  if (purchasePrice && toAmount_(purchasePrice) === null) {
    return { success: false, errors: ['Purchase price must be a number.'] };
  }
  var purchaseDate = String(data.purchaseDate || '').trim();
  if (purchaseDate && !/^\d{4}-\d{2}-\d{2}$/.test(purchaseDate)) {
    return { success: false, errors: ['Purchase date is not valid.'] };
  }

  return withLock_(function () {
    var table = openTable_('properties', true);
    var id = String(data.propertyId || '').trim();
    var existing = id ? findRowById_(table, id) : null;
    if (id && !existing) {
      return { success: false, errors: ['That property no longer exists. Reload and try again.'] };
    }

    var duplicate = table.rows.some(function (row) {
      return row !== existing &&
        cellString_(row.values, table.columns.name).toLowerCase() === name.toLowerCase() &&
        !isArchivedStatus_(cellString_(row.values, table.columns.status));
    });
    if (duplicate) {
      return { success: false, errors: ['Another active property is already called "' + name + '".'] };
    }

    var company = companyLookup_().byId[String(data.companyId || '').trim()] || null;
    var values = {
      name: name,
      companyId: company ? company.companyId : '',
      companyName: company ? company.name : '',
      address: cleanText_(data.address, 300),
      purchaseDate: purchaseDate ? isoToDate_(purchaseDate) : '',
      purchasePrice: purchasePrice ? toAmount_(purchasePrice) : '',
      currency: data.currency === 'CRC' ? 'CRC' : 'USD',
      status: data.status === 'Archived' ? 'Archived' : 'Active',
      notes: cleanNotes_(data.notes),
      updatedAt: new Date()
    };

    if (existing) {
      writeFields_(table, existing.rowNumber, values);
      return { success: true, propertyId: id };
    }

    values.id = nextId_(table);
    values.createdAt = new Date();
    appendFields_(table, values);
    return { success: true, propertyId: values.id };
  });
}

/** @return {{success: boolean, errors: string[]=}} */
function archiveProperty(propertyId) {
  return archiveRow_('properties', propertyId, 'property');
}

/* ---------- companies ---------- */

/** Active companies, for the property form's dropdown. */
function getCompanies() {
  return getAllCompanies().filter(function (c) {
    return !isArchivedStatus_(c.status);
  });
}

/** All companies, for the Companies screen. */
function getAllCompanies() {
  var table = openTable_('companies', true);
  return table.rows.map(function (row) {
    return {
      companyId: cellString_(row.values, table.columns.id),
      name: cellString_(row.values, table.columns.name),
      cedulaJuridica: cellString_(row.values, table.columns.cedulaJuridica),
      address: cellString_(row.values, table.columns.address),
      phone: cellString_(row.values, table.columns.phone),
      email: cellString_(row.values, table.columns.email),
      bank: cellString_(row.values, table.columns.bank),
      status: normalizeStatus_(cellString_(row.values, table.columns.status)),
      notes: cellString_(row.values, table.columns.notes)
    };
  }).filter(function (c) {
    return c.name;
  });
}

/**
 * @param {Object} data fields from the company form
 * @return {{success: boolean, errors: string[]=, companyId: string=}}
 */
function saveCompany(data) {
  data = data || {};
  var name = cleanText_(data.name, 160);
  if (!name) {
    return { success: false, errors: ['Company name is required.'] };
  }
  var email = cleanText_(data.email, 200);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, errors: ['Email address is not valid.'] };
  }

  return withLock_(function () {
    var table = openTable_('companies', true);
    var id = String(data.companyId || '').trim();
    var existing = id ? findRowById_(table, id) : null;
    if (id && !existing) {
      return { success: false, errors: ['That company no longer exists. Reload and try again.'] };
    }

    var values = {
      name: name,
      cedulaJuridica: cleanText_(data.cedulaJuridica, 40),
      address: cleanText_(data.address, 300),
      phone: cleanText_(data.phone, 60),
      email: email,
      bank: cleanText_(data.bank, 200),
      status: data.status === 'Archived' ? 'Archived' : 'Active',
      notes: cleanNotes_(data.notes),
      updatedAt: new Date()
    };

    if (existing) {
      writeFields_(table, existing.rowNumber, values);
      return { success: true, companyId: id };
    }

    values.id = nextId_(table);
    values.createdAt = new Date();
    appendFields_(table, values);
    return { success: true, companyId: values.id };
  });
}

/** @return {{success: boolean, errors: string[]=}} */
function archiveCompany(companyId) {
  return archiveRow_('companies', companyId, 'company');
}

/* ---------- expense categories ---------- */

/** Active expense categories, as names, in sheet order. */
function getExpenseCategories() {
  var table = openTable_('categories', false);
  var nameColumn = table.columns.name !== -1 ? table.columns.name : 0;
  var seen = {};

  return table.rows.map(function (row) {
    return {
      name: cellString_(row.values, nameColumn),
      status: cellString_(row.values, table.columns.status)
    };
  }).filter(function (c) {
    var key = c.name.toLowerCase();
    if (!c.name || seen[key] || isArchivedStatus_(c.status)) {
      return false;
    }
    seen[key] = true;
    return true;
  }).map(function (c) {
    return c.name;
  });
}

/* ---------- payments ---------- */

/** Saves a payment from the Receive Payment form. */
function saveReceivePayment(payment) {
  return savePayment_('income', payment);
}

/** Saves a payment from the Make Payment form. */
function saveMakePayment(payment) {
  return savePayment_('expense', payment);
}

function savePayment_(type, payment) {
  payment = payment || {};
  var property = cleanText_(payment.property, 120);
  var amount = toAmount_(payment.amount);
  var date = String(payment.date || '').trim();

  if (!property) {
    throw new Error('Choose a property.');
  }
  if (amount === null || amount <= 0) {
    throw new Error('Enter an amount greater than zero.');
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    throw new Error('The date is not valid.');
  }

  var category = type === 'income'
    ? cleanText_(payment.incomeType, 80)
    : cleanText_(payment.expenseCategory, 80);

  return withLock_(function () {
    var target = getPaymentTab_(type);
    var columns = target.columns;
    var sheet = target.sheet;

    var categoryField = type === 'income' ? 'incomeType' : 'expenseCategory';
    var categoryHeader = type === 'income' ? 'Income Type' : 'Expense Category';
    if (columns[categoryField] === -1) {
      // e.g. an income tab with a generic "Category" column.
      columns[categoryField] = columns[type === 'income' ? 'expenseCategory' : 'incomeType'];
    }

    var fields = [
      ['date', 'Date', isoToDate_(date)],
      ['property', 'Property', property],
      [categoryField, categoryHeader, category],
      ['currency', 'Currency', payment.currency === 'CRC' ? 'CRC' : 'USD'],
      ['amount', 'Amount', Math.round(amount * 100) / 100],
      ['notes', 'Notes', cleanNotes_(payment.notes)],
      ['enteredBy', 'Entered By', cleanText_(payment.enteredBy, 80)],
      ['createdAt', 'Created At', new Date()]
    ];

    var row = [];
    fields.forEach(function (f) {
      var column = columns[f[0]];
      if (column === undefined || column === -1) {
        column = addHeader_(sheet, target.header, f[1]);
        columns[f[0]] = column;
      }
      while (row.length <= column) {
        row.push('');
      }
      row[column] = f[2];
    });

    var rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    return { success: true, sheet: sheet.getName(), row: rowNumber };
  });
}

// The income or expense tab the Dashboard reads (listPaymentSheets_ in
// Dashboard.gs). Creates "Income" / "Expenses" if there is none yet.
function getPaymentTab_(type) {
  var match = listPaymentSheets_().sheets.filter(function (entry) {
    return entry.type === type;
  })[0];

  var sheet;
  var header;
  if (match) {
    sheet = match.sheet;
    header = match.values[0].slice();
  } else {
    var spec = PFMS_PAYMENT_TABS[type];
    var spreadsheet = getPfmsSpreadsheet_();
    sheet = spreadsheet.getSheetByName(spec.name) || spreadsheet.insertSheet(spec.name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  var mapped = mapDashboardColumns_(header);
  var normalized = header.map(normalizeHeader_);
  mapped.createdAt = firstIndex_(normalized, ['created at', 'created', 'fecha creacion']);
  return { sheet: sheet, header: header, columns: mapped };
}

/* ---------- table helpers ---------- */

function openTable_(kind, needsId) {
  var spec = PFMS_TABLES[kind];
  var spreadsheet = getPfmsSpreadsheet_();
  var sheet = findTab_(spreadsheet, spec.tabNames);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(spec.tabNames[0]);
    sheet.getRange(1, 1, 1, spec.defaultHeaders.length).setValues([spec.defaultHeaders]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (spec.defaultRows) {
      sheet.getRange(2, 1, spec.defaultRows.length, 2).setValues(spec.defaultRows.map(function (name) {
        return [name, 'Active'];
      }));
    }
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, lastColumn).getValues() : [[]];
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, spec.defaultHeaders.length).setValues([spec.defaultHeaders]).setFontWeight('bold');
    values = [spec.defaultHeaders.slice()];
  }

  var header = values[0];
  var normalized = header.map(normalizeHeader_);
  var columns = {};
  Object.keys(spec.fields).forEach(function (field) {
    columns[field] = firstIndex_(normalized, spec.fields[field]);
  });
  // "Company" means the company's name in a property tab only when no other
  // column already holds the property name.
  if (kind === 'properties' && columns.companyName !== -1 && columns.companyName === columns.name) {
    columns.companyName = -1;
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var isBlank = values[r].every(function (v) {
      return v === '' || v === null;
    });
    if (!isBlank) {
      rows.push({ rowNumber: r + 1, values: values[r] });
    }
  }

  var table = { kind: kind, spec: spec, sheet: sheet, header: header, columns: columns, rows: rows };
  if (needsId) {
    ensureIds_(table);
  }
  return table;
}

function findTab_(spreadsheet, names) {
  var wanted = names.map(normalizeHeader_);
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (wanted.indexOf(normalizeHeader_(sheets[i].getName())) !== -1) {
      return sheets[i];
    }
  }
  return null;
}

// The app selects, edits and archives rows by ID. Adds an ID column if the
// tab has none and fills in IDs for rows that are missing one.
function ensureIds_(table) {
  var missing = table.rows.filter(function (row) {
    return table.columns.id === -1 || !cellString_(row.values, table.columns.id);
  });
  if (!missing.length && table.columns.id !== -1) {
    return;
  }

  withLock_(function () {
    if (table.columns.id === -1) {
      var headerText = table.kind === 'companies' ? 'Company ID' : 'Property ID';
      table.columns.id = addHeader_(table.sheet, table.header, headerText);
    }
    missing.forEach(function (row) {
      var id = nextId_(table);
      table.sheet.getRange(row.rowNumber, table.columns.id + 1).setValue(id);
      while (row.values.length <= table.columns.id) {
        row.values.push('');
      }
      row.values[table.columns.id] = id;
    });
  });
}

// Continues the tab's existing ID pattern (e.g. PROP-007 -> PROP-008),
// or starts PROP-0001 / COMP-0001.
function nextId_(table) {
  var prefix = table.spec.idPrefix;
  var width = 4;
  var max = 0;
  var counts = {};

  table.rows.forEach(function (row) {
    var match = /^(.*?)(\d+)$/.exec(cellString_(row.values, table.columns.id));
    if (match) {
      counts[match[1]] = (counts[match[1]] || 0) + 1;
    }
  });
  var usual = Object.keys(counts).sort(function (a, b) {
    return counts[b] - counts[a];
  })[0];
  if (usual !== undefined) {
    prefix = usual;
  }

  var widthSeen = 0;
  table.rows.forEach(function (row) {
    var match = /^(.*?)(\d+)$/.exec(cellString_(row.values, table.columns.id));
    if (match && match[1] === prefix) {
      max = Math.max(max, parseInt(match[2], 10));
      widthSeen = Math.max(widthSeen, match[2].length);
    }
  });
  if (widthSeen) {
    width = widthSeen;
  }

  var number = String(max + 1);
  while (number.length < width) {
    number = '0' + number;
  }
  var id = prefix + number;
  table.rows.push({ rowNumber: -1, values: idRowStub_(table, id) });
  return id;
}

function idRowStub_(table, id) {
  var stub = [];
  while (stub.length <= table.columns.id) {
    stub.push('');
  }
  stub[table.columns.id] = id;
  return stub;
}

function findRowById_(table, id) {
  for (var i = 0; i < table.rows.length; i++) {
    if (table.rows[i].rowNumber > 0 && cellString_(table.rows[i].values, table.columns.id) === id) {
      return table.rows[i];
    }
  }
  return null;
}

function writeFields_(table, rowNumber, values) {
  Object.keys(values).forEach(function (field) {
    if (values[field] === undefined) {
      return;
    }
    var column = columnForWrite_(table, field, values[field]);
    if (column !== -1) {
      table.sheet.getRange(rowNumber, column + 1).setValue(values[field]);
    }
  });
}

function appendFields_(table, values) {
  var row = [];
  Object.keys(values).forEach(function (field) {
    var column = columnForWrite_(table, field, values[field]);
    if (column === -1) {
      return;
    }
    while (row.length <= column) {
      row.push('');
    }
    row[column] = values[field];
  });
  var rowNumber = table.sheet.getLastRow() + 1;
  table.sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return rowNumber;
}

// Columns to add when a field has a value but nowhere to go. Optional extras
// such as a property's company name or created/updated stamps are only
// written when the tab already has a column for them.
var PFMS_ADDABLE_HEADERS = {
  properties: { name: 'Name', companyId: 'Company ID', address: 'Address', purchaseDate: 'Purchase Date',
    purchasePrice: 'Purchase Price', currency: 'Currency', status: 'Status', notes: 'Notes' },
  companies: { name: 'Name', cedulaJuridica: 'Cedula Juridica', address: 'Address', phone: 'Phone', email: 'Email',
    bank: 'Bank', status: 'Status', notes: 'Notes' }
};

function columnForWrite_(table, field, value) {
  var column = table.columns[field];
  if (column !== undefined && column !== -1) {
    return column;
  }
  var addable = (PFMS_ADDABLE_HEADERS[table.kind] || {})[field];
  // Only add a column when there is something to put in it. A tab with no
  // Status column already reads as all Active.
  var isEmpty = value === '' || value === null || value === undefined || (field === 'status' && value === 'Active');
  if (!addable || isEmpty) {
    return -1;
  }
  column = addHeader_(table.sheet, table.header, addable);
  table.columns[field] = column;
  return column;
}

// Adds a header after the last used header cell (or after all columns if
// that cell has data below it) and returns its zero-based index.
function addHeader_(sheet, header, text) {
  var lastUsed = -1;
  for (var c = 0; c < header.length; c++) {
    if (String(header[c] === null ? '' : header[c]).trim()) {
      lastUsed = c;
    }
  }
  var column = lastUsed + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1 && column < sheet.getLastColumn()) {
    var below = sheet.getRange(2, column + 1, lastRow - 1, 1).getValues();
    var hasData = below.some(function (r) {
      return String(r[0] === null ? '' : r[0]).trim() !== '';
    });
    if (hasData) {
      column = Math.max(sheet.getLastColumn(), header.length);
    }
  }

  var cell = sheet.getRange(1, column + 1);
  cell.setValue(text);
  if (column > 0) {
    sheet.getRange(1, column).copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  while (header.length <= column) {
    header.push('');
  }
  header[column] = text;
  return column;
}

function archiveRow_(kind, id, label) {
  id = String(id || '').trim();
  if (!id) {
    return { success: false, errors: ['Choose a ' + label + ' first.'] };
  }
  return withLock_(function () {
    var table = openTable_(kind, true);
    var row = findRowById_(table, id);
    if (!row) {
      return { success: false, errors: ['That ' + label + ' no longer exists. Reload and try again.'] };
    }
    writeFields_(table, row.rowNumber, { status: 'Archived', updatedAt: new Date() });
    return { success: true };
  });
}

function companyLookup_() {
  var byId = {};
  var byName = {};
  getAllCompanies().forEach(function (c) {
    byId[c.companyId] = c;
    byName[c.name.toLowerCase()] = c;
  });
  return { byId: byId, byName: byName };
}

/* ---------- small helpers ---------- */

// Re-entrant: nested calls (e.g. ensureIds_ inside saveProperty) reuse the
// lock already held.
var pfmsLockDepth_ = 0;
function withLock_(fn) {
  if (pfmsLockDepth_ > 0) {
    return fn();
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  pfmsLockDepth_++;
  try {
    return fn();
  } finally {
    pfmsLockDepth_--;
    lock.releaseLock();
  }
}

function normalizeHeader_(value) {
  return String(value === null || value === undefined ? '' : value)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[_\s]+/g, ' ').trim();
}

function firstIndex_(normalizedHeader, names) {
  for (var i = 0; i < names.length; i++) {
    var index = normalizedHeader.indexOf(normalizeHeader_(names[i]));
    if (index !== -1) {
      return index;
    }
  }
  return -1;
}

function cellString_(values, column) {
  if (column === -1 || column === undefined || column >= values.length) {
    return '';
  }
  var value = values[column];
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

function cellDate_(values, column) {
  if (column === -1 || column === undefined || column >= values.length) {
    return '';
  }
  return toIsoDate_(values[column], Session.getScriptTimeZone()) || cellString_(values, column);
}

function cellNumberOrString_(values, column) {
  if (column === -1 || column === undefined || column >= values.length) {
    return '';
  }
  var value = values[column];
  if (typeof value === 'number') {
    return value;
  }
  return cellString_(values, column);
}

function normalizeStatus_(value) {
  return isArchivedStatus_(value) ? 'Archived' : 'Active';
}

function isArchivedStatus_(value) {
  var text = normalizeHeader_(value);
  return ['archived', 'archivado', 'archivada', 'inactive', 'inactivo', 'inactiva', 'no', 'false', 'sold', 'vendido',
    'vendida'].indexOf(text) !== -1;
}

function cleanText_(value, maxLength) {
  return String(value === null || value === undefined ? '' : value).replace(/\s+/g, ' ').trim().substring(0, maxLength);
}

function cleanNotes_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/\r\n?/g, '\n').trim().substring(0, 2000);
}

function isoToDate_(iso) {
  var parts = iso.split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}
