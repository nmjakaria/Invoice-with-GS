/**
 * INVOICE MAKER - Google Apps Script backend
 * Sheets used:
 *  - "Invoices":  one row per saved invoice
 *  - "Config":    stores the last invoice number (auto-created)
 *  - "Settings":  company name, currency, invoice prefix, etc. (auto-created)
 */

const SHEET_NAME = 'Invoices';
const CONFIG_SHEET = 'Config';
const SETTINGS_SHEET = 'Settings';
const EXPENSE_SHEET = 'Expenses';
const EXPENSE_HEADERS = [
  'Expense ID', 'Date', 'Category', 'Description', 'Amount', 'Notes', 'Created At'
];
const INVOICE_HEADERS = [
  'Invoice ID', 'Date', 'Due Date', 'Customer Name', 'Customer Email',
  'Customer Phone', 'Customer Address', 'Items JSON', 'Subtotal',
  'Tax %', 'Tax Amount', 'Discount', 'Total', 'Status', 'Notes', 'Created At'
];

const DEFAULT_SETTINGS = {
  'Company Name': 'Ledger Studio',
  'Tagline': 'Invoicing for modern business',
  'Currency Symbol': '$',
  'Company Email': 'billing@ledger.studio',
  'Company Address': '128 Commerce Street, Springfield',
  'Invoice Prefix': 'INV-',
  'Footer': 'Thank you for your business',
  'Spreadsheet ID': ''
};

/** Serves the web app UI */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Invoice Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** The spreadsheet this script is attached to (stores Settings + Config). */
function getActiveSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Extracts a spreadsheet ID from a URL or raw ID. */
function extractSpreadsheetId(ref) {
  ref = String(ref || '').trim();
  if (!ref) return '';
  const m = ref.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(ref)) return ref;
  return ref;
}

/** Reads the optional linked spreadsheet ID from the attached spreadsheet's Settings tab. */
function getCustomSpreadsheetId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET);
    if (!sheet) return '';
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === 'Spreadsheet ID' && data[i][1]) {
        return String(data[i][1]).trim();
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

/** The spreadsheet that holds invoice data:
 *  the linked one from Settings if provided (and accessible), otherwise
 *  the spreadsheet this script is attached to. */
function getSpreadsheet() {
  const customId = getCustomSpreadsheetId();
  if (customId) {
    const id = extractSpreadsheetId(customId);
    if (id) {
      try {
        return SpreadsheetApp.openById(id);
      } catch (e) {
        // Not accessible or invalid — fall back to the attached spreadsheet.
      }
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Finds the sheet that holds invoice data.
 *  Prefers a tab named "Invoices", but also auto-detects any tab whose
 *  header row looks like an invoice sheet (e.g. it was renamed), or whose
 *  first column contains INV-xxxx identifiers. Creates "Invoices" with
 *  headers only if nothing matches. */
function getInvoiceSheet() {
  const ss = getSpreadsheet();

  function looksLikeInvoiceSheet(s) {
    const colCount = Math.min(s.getLastColumn(), 20);
    if (colCount === 0) return false;
    const headers = s.getRange(1, 1, 1, colCount).getValues()[0]
      .map(function (h) { return String(h).replace(/\s+/g, '').toLowerCase(); });
    if (headers.indexOf('invoiceid') >= 0 || headers.indexOf('customer') >= 0) return true;
    const lastRow = Math.min(s.getLastRow(), 300);
    if (lastRow < 2) return false;
    const colA = s.getRange(1, 1, lastRow, 1).getValues();
    for (let i = 0; i < colA.length; i++) {
      if (/^INV-[\w-]+$/i.test(String(colA[i][0]).trim())) return true;
    }
    return false;
  }

  const candidates = ss.getSheets().filter(looksLikeInvoiceSheet);
  const exact = ss.getSheetByName(SHEET_NAME);

  if (exact) {
    const exactHasData = exact.getLastRow() > 1;
    if (exactHasData || candidates.length === 0) return exact;
  }

  if (candidates.length > 0) {
    candidates.sort(function (a, b) { return b.getLastRow() - a.getLastRow(); });
    return candidates[0];
  }

  const sheet = ss.insertSheet(SHEET_NAME);
  sheet.appendRow(INVOICE_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

/** Returns diagnostic info about where the app is reading invoices from.
 *  Shown in the UI when the list looks empty so problems are obvious. */
function getDataInfo() {
  try {
    const ss = getSpreadsheet();
    const sheet = getInvoiceSheet();
    const data = sheet.getDataRange().getValues();
    const headers = (data[0] || []).map(function (h) { return String(h); });
    const firstCell = String(data.length ? (data[0][0] || '') : '').trim();
    const hasHeader = !/^INV-/i.test(firstCell);
    return {
      sheetName: sheet.getName(),
      sheetCount: ss.getSheets().length,
      headerRow: headers,
      dataRows: data.length ? (hasHeader ? data.length - 1 : data.length) : 0
    };
  } catch (err) {
    return { sheetName: 'error', sheetCount: 0, headerRow: [], dataRows: 0, error: err.message };
  }
}

/** Thread-safe auto-incrementing invoice number, e.g. INV-1001 */
function getNextInvoiceNumber() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet();
    let configSheet = ss.getSheetByName(CONFIG_SHEET);
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET);
      configSheet.getRange('A1').setValue('LastInvoiceNumber');
      configSheet.getRange('B1').setValue(1000);
      configSheet.hideSheet();
    }
    const cell = configSheet.getRange('B1');
    const nextNum = Number(cell.getValue()) + 1;
    cell.setValue(nextNum);
    return getInvoicePrefix() + nextNum;
  } finally {
    lock.releaseLock();
  }
}

/** Preview the next invoice number WITHOUT consuming it (for display only) */
function previewNextInvoiceNumber() {
  const ss = getSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) return getInvoicePrefix() + '1001';
  return getInvoicePrefix() + (Number(configSheet.getRange('B1').getValue()) + 1);
}

function getInvoicePrefix() {
  try {
    return String(getSettings()['Invoice Prefix'] || 'INV-');
  } catch (e) {
    return 'INV-';
  }
}

/**
 * Saves an invoice submitted from the client.
 * invoiceData: {
 *   date, dueDate, customerName, customerEmail, customerPhone, customerAddress,
 *   items: [{description, qty, price}, ...],
 *   taxPercent, discount, status, notes
 * }
 */
function saveInvoice(invoiceData) {
  try {
    if (!invoiceData || !invoiceData.customerName) {
      throw new Error('Customer name is required.');
    }
    const items = invoiceData.items || [];
    if (items.length === 0) {
      throw new Error('Add at least one line item.');
    }

    // Always calculate totals on the server - never trust client-side math
    let subtotal = 0;
    items.forEach(function (it) {
      const qty = parseFloat(it.qty) || 0;
      const price = parseFloat(it.price) || 0;
      subtotal += qty * price;
    });
    const taxPercent = parseFloat(invoiceData.taxPercent) || 0;
    const discount = parseFloat(invoiceData.discount) || 0;
    const taxAmount = subtotal * taxPercent / 100;
    const total = subtotal + taxAmount - discount;

    const invoiceId = getNextInvoiceNumber();
    const sheet = getInvoiceSheet();

    sheet.appendRow([
      invoiceId,
      invoiceData.date || new Date(),
      invoiceData.dueDate || '',
      invoiceData.customerName,
      invoiceData.customerEmail || '',
      invoiceData.customerPhone || '',
      invoiceData.customerAddress || '',
      JSON.stringify(items),
      subtotal,
      taxPercent,
      taxAmount,
      discount,
      total,
      invoiceData.status || 'Unpaid',
      invoiceData.notes || '',
      new Date()
    ]);

    return {
      success: true,
      invoiceId: invoiceId,
      subtotal: subtotal,
      taxAmount: taxAmount,
      total: total
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Returns all invoices as an array of objects, most recent first.
 *  Maps positions using INVOICE_HEADERS, and works even when the sheet
 *  has no header row (detected by looking for an INV-xxxx id in row 1). */
/** Internal: returns invoices as a real array (for server-side use only) */
function getAllInvoicesRaw() {
  const sheet = getInvoiceSheet();
  const data = sheet.getDataRange().getValues();
  if (!data.length) return [];
  const firstCell = String(data[0][0] || '').trim();
  const hasHeader = !/^INV-/i.test(firstCell);
  const rows = hasHeader ? data.slice(1) : data;
  if (rows.length === 0) return [];
  const mapped = rows.map(function (row) {
    const obj = {};
    INVOICE_HEADERS.forEach(function (h, i) { obj[h] = row[i]; });
    try { obj['Items'] = JSON.parse(obj['Items JSON']); } catch (e) { obj['Items'] = []; }
    return obj;
  });
  return mapped.reverse();
}

/** Client-facing: same data as a JSON string (see note on google.script.run
 *  silently returning null for arrays containing nested arrays-of-objects). */
function getAllInvoices() {
  return JSON.stringify(getAllInvoicesRaw());
}

/** Looks up a single invoice by ID (for viewing/reprinting) */
function getInvoiceById(invoiceId) {
  const invoices = getAllInvoicesRaw();   // আগে ছিল getAllInvoices()
  return invoices.find(function (inv) { return inv['Invoice ID'] === invoiceId; }) || null;
}

/** Flips an invoice's status (e.g. "Mark as Paid" from the dashboard) */
function updateInvoiceStatus(invoiceId, status) {
  try {
    const sheet = getInvoiceSheet();
    const data = sheet.getDataRange().getValues();
    const idCol = INVOICE_HEADERS.indexOf('Invoice ID');
    const statusCol = INVOICE_HEADERS.indexOf('Status');
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][idCol]) === String(invoiceId)) {
        sheet.getRange(r + 1, statusCol + 1).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: 'Invoice not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ---------------- Expenses ---------------- */

/** Finds (or creates) the tab that holds expense rows. */
function getExpenseSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(EXPENSE_SHEET);
  if (sheet) return sheet;
  const created = ss.insertSheet(EXPENSE_SHEET);
  created.appendRow(EXPENSE_HEADERS);
  created.setFrozenRows(1);
  return created;
}

/** Auto-incrementing expense number, e.g. EXP-1001 (shares the Config tab) */
function getNextExpenseNumber() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet();
    let configSheet = ss.getSheetByName(CONFIG_SHEET);
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET);
      configSheet.getRange('A1').setValue('LastInvoiceNumber');
      configSheet.getRange('B1').setValue(1000);
      configSheet.hideSheet();
    }
    const data = configSheet.getDataRange().getValues();
    let row = -1, next = 1000;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === 'LastExpenseNumber') {
        row = i + 1;
        next = Number(data[i][1]) || 1000;
      }
    }
    next += 1;
    if (row > 0) configSheet.getRange(row, 2).setValue(next);
    else configSheet.appendRow(['LastExpenseNumber', next]);
    return 'EXP-' + next;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves an expense submitted from the client.
 * expenseData: { date, category, description, amount, notes }
 */
function saveExpense(expenseData) {
  try {
    if (!expenseData || !expenseData.category || !expenseData.description) {
      throw new Error('Category and description are required.');
    }
    const amount = parseFloat(expenseData.amount);
    if (isNaN(amount) || amount < 0) {
      throw new Error('Enter a valid amount.');
    }

    const expenseId = getNextExpenseNumber();
    getExpenseSheet().appendRow([
      expenseId,
      expenseData.date || new Date(),
      expenseData.category,
      expenseData.description,
      amount,
      expenseData.notes || '',
      new Date()
    ]);

    return { success: true, expenseId: expenseId, amount: amount };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Internal: returns expenses as a real array (for server-side use only) */
function getAllExpensesRaw() {
  const sheet = getExpenseSheet();
  const data = sheet.getDataRange().getValues();
  if (!data.length) return [];
  const firstCell = String(data[0][0] || '').trim();
  const hasHeader = !/^EXP-/i.test(firstCell);
  const rows = hasHeader ? data.slice(1) : data;
  if (rows.length === 0) return [];
  return rows.map(function (row) {
    const obj = {};
    EXPENSE_HEADERS.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  }).reverse();
}

/** Client-facing: same data as a JSON string */
function getAllExpenses() {
  return JSON.stringify(getAllExpensesRaw());
}

/** Removes an expense row by ID */
function deleteExpense(expenseId) {
  try {
    const sheet = getExpenseSheet();
    const data = sheet.getDataRange().getValues();
    const idCol = EXPENSE_HEADERS.indexOf('Expense ID');
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][idCol]) === String(expenseId)) {
        sheet.deleteRow(r + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Expense not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ---------------- Settings (company name, currency, etc.) ---------------- */

function getSettingsSheet() {
  const ss = getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(['Key', 'Value']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Reads saved settings and merges with defaults */
function getSettings() {
  const sheet = getSettingsSheet();
  const data = sheet.getDataRange().getValues();
  const saved = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) saved[String(data[i][0])] = data[i][1];
  }
  const out = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    out[k] = saved[k] !== undefined && saved[k] !== '' ? saved[k] : DEFAULT_SETTINGS[k];
  });
  return out;
}

/** Saves any provided settings keys (unknown keys are ignored) */
function saveSettings(settings) {
  try {
    const sheet = getSettingsSheet();
    const data = sheet.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) rowByKey[String(data[i][0])] = i + 1; // 1-based sheet row
    }
    Object.keys(DEFAULT_SETTINGS).forEach(function (key) {
      const val = settings[key] !== undefined ? String(settings[key]) : '';
      if (rowByKey[key]) {
        sheet.getRange(rowByKey[key], 2).setValue(val);
      } else {
        sheet.appendRow([key, val]);
      }
    });
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Aggregated numbers + monthly revenue series + daily income/expense for the dashboard */
function getDashboardStats() {
  const invoices = getAllInvoicesRaw();   // আগে ছিল getAllInvoices()
  const expenses = getAllExpensesRaw();
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  let totalRevenue = 0, outstanding = 0, overdueCount = 0, overdueAmount = 0;
  const monthly = {};

  invoices.forEach(function (inv) {
    const total = Number(inv['Total']) || 0;
    const status = inv['Status'];
    const dueDate = inv['Due Date'] ? new Date(inv['Due Date']) : null;
    const createdAt = inv['Created At'] ? new Date(inv['Created At']) : new Date(inv['Date']);

    if (status === 'Paid') {
      totalRevenue += total;
    } else {
      outstanding += total;
      if (dueDate && !isNaN(dueDate) && dueDate < now) {
        overdueCount++;
        overdueAmount += total;
      }
    }

    if (createdAt && !isNaN(createdAt)) {
      const key = Utilities.formatDate(createdAt, tz, 'yyyy-MM');
      monthly[key] = (monthly[key] || 0) + total;
    }
  });

  const series = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = Utilities.formatDate(d, tz, 'yyyy-MM');
    series.push({
      month: Utilities.formatDate(d, tz, 'MMM'),
      total: monthly[key] || 0
    });
  }

  /* Daily income = paid invoices dated that day; expense = expenses dated that day */
  const daily = {};

  invoices.forEach(function (inv) {
    if (inv['Status'] !== 'Paid') return;
    const d = inv['Date'] ? new Date(inv['Date']) : new Date(inv['Created At']);
    if (d && !isNaN(d)) {
      const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      daily[key] = daily[key] || { income: 0, expense: 0 };
      daily[key].income += Number(inv['Total']) || 0;
    }
  });

  expenses.forEach(function (exp) {
    const d = exp['Date'] ? new Date(exp['Date']) : null;
    if (d && !isNaN(d)) {
      const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      daily[key] = daily[key] || { income: 0, expense: 0 };
      daily[key].expense += Number(exp['Amount']) || 0;
    }
  });

  const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const today = daily[todayKey] || { income: 0, expense: 0 };

  const dailySeries = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i);
    const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const day = daily[key] || { income: 0, expense: 0 };
    dailySeries.push({
      day: Utilities.formatDate(d, tz, 'EEE dd'),
      income: day.income,
      expense: day.expense
    });
  }

  return {
    totalRevenue: totalRevenue,
    outstanding: outstanding,
    overdueCount: overdueCount,
    overdueAmount: overdueAmount,
    invoiceCount: invoices.length,
    monthlySeries: series,
    todayIncome: today.income,
    todayExpense: today.expense,
    dailySeries: dailySeries
  };
}