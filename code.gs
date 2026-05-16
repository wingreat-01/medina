// ══════════════════════════════════════════════════════════════════
//  AJ Medina POS — Google Apps Script Backend (code.gs)
//  Logs POS orders to a Google Sheet + User Authentication
// ══════════════════════════════════════════════════════════════════
//
//  SETUP GUIDE:
//  1. Open Google Sheets → Extensions → Apps Script
//  2. Paste this entire file as code.gs
//  3. Click Deploy → New deployment → Web App
//     - Execute as:  Me
//     - Who has access:  Anyone
//  4. Copy the Web App URL
//  5. Paste it into index.html where it says:
//     const GAS_URL = 'YOUR_GAS_WEB_APP_URL_HERE';
//
//  FIRST-TIME SETUP:
//  The admin account is auto-created on first run:
//    Username: admin
//    Password: admin123
//  Change it immediately after first login via the Admin panel.
//
// ══════════════════════════════════════════════════════════════════

// ── Config ──────────────────────────────────────────────────────
var SHEET_NAME_ORDERS    = 'Orders';
var SHEET_NAME_ITEMS     = 'Order Items';
var SHEET_NAME_INVENTORY = 'Inventory';
var SHEET_NAME_USERS     = 'Users';

// ── Entry Points ─────────────────────────────────────────────────

/**
 * GET  → health check / simple ping, or fetch inventory/transactions/users
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

  if (action === 'inventory') {
    return getInventory();
  }
  if (action === 'transactions') {
    return getTransactions();
  }
  if (action === 'getUsers') {
    return getUsers();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', app: 'AJ Medina POS', time: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST → receive order or inventory action or auth action from POS frontend
 */
function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    // ── Auth routes ──
    if (data.action === 'login')      return handleLogin(data);
    if (data.action === 'createUser') return createUser(data);
    if (data.action === 'updateUser') return updateUser(data);
    if (data.action === 'deleteUser') return deleteUser(data);

    // ── Inventory routes ──
    if (data.action === 'addInventory')    return addInventoryRow(data);
    if (data.action === 'updateInventory') return updateInventoryRow(data);
    if (data.action === 'deleteInventory') return deleteInventoryRow(data);

    // Default: save order
    saveOrder(data);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderNum: data.orderNum }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Auth / User Management ────────────────────────────────────────

/**
 * Ensures the Users sheet exists with the default admin account.
 */
function ensureUsersSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var header = ['Username', 'Password', 'Role', 'Created At', 'Last Login'];
    sheet.appendRow(header);
    styleHeaderRow(sheet, header.length);
    // Seed default admin
    sheet.appendRow(['admin', 'admin123', 'admin', new Date().toISOString(), '']);
  }
}

/**
 * Login: validate username + password, return role.
 */
function handleLogin(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS);
  ensureUsersSheet(sheet);

  var last = sheet.getLastRow();
  if (last < 2) {
    return jsonResponse({ success: false, error: 'No users found.' });
  }

  var rows = sheet.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    var uname = String(rows[i][0]).trim();
    var upass = String(rows[i][1]).trim();
    var urole = String(rows[i][2]).trim();
    if (uname === String(data.username).trim() && upass === String(data.password).trim()) {
      // Update last login
      sheet.getRange(i + 2, 5).setValue(new Date().toISOString());
      return jsonResponse({ success: true, role: urole, username: uname });
    }
  }
  return jsonResponse({ success: false, error: 'Invalid username or password.' });
}

/**
 * Returns all users (passwords redacted).
 */
function getUsers() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS);
  ensureUsersSheet(sheet);

  var last = sheet.getLastRow();
  if (last < 2) {
    return jsonResponse({ success: true, users: [] });
  }

  var rows = sheet.getRange(2, 1, last - 1, 5).getValues();
  var users = rows.map(function(r, i) {
    return {
      rowIndex:   i,
      username:   r[0],
      role:       r[2],
      createdAt:  r[3],
      lastLogin:  r[4]
    };
  }).filter(function(u) { return u.username !== ''; });

  return jsonResponse({ success: true, users: users });
}

/**
 * Create a new user (admin only — enforced on frontend).
 */
function createUser(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS);
  ensureUsersSheet(sheet);

  // Check for duplicate username
  var last = sheet.getLastRow();
  if (last >= 2) {
    var rows = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(data.username).trim().toLowerCase()) {
        return jsonResponse({ success: false, error: 'Username already exists.' });
      }
    }
  }

  sheet.appendRow([
    data.username || '',
    data.password || '',
    data.role     || 'user',
    new Date().toISOString(),
    ''
  ]);
  return jsonResponse({ success: true });
}

/**
 * Update an existing user row.
 */
function updateUser(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS);

  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return jsonResponse({ success: false, error: 'Invalid row index.' });
  }

  var existing = sheet.getRange(sheetRow, 1, 1, 5).getValues()[0];
  sheet.getRange(sheetRow, 1, 1, 5).setValues([[
    data.username || existing[0],
    data.password || existing[1],   // keep old password if blank
    data.role     || existing[2],
    existing[3],
    existing[4]
  ]]);
  return jsonResponse({ success: true });
}

/**
 * Delete a user row.
 */
function deleteUser(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS);

  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return jsonResponse({ success: false, error: 'Invalid row index.' });
  }

  // Protect: never delete the last admin
  var username = sheet.getRange(sheetRow, 1).getValue();
  var role     = sheet.getRange(sheetRow, 3).getValue();
  if (role === 'admin') {
    // Count other admins
    var last = sheet.getLastRow();
    var adminCount = 0;
    if (last >= 2) {
      var roles = sheet.getRange(2, 3, last - 1, 1).getValues();
      roles.forEach(function(r) { if (r[0] === 'admin') adminCount++; });
    }
    if (adminCount <= 1) {
      return jsonResponse({ success: false, error: 'Cannot delete the last admin account.' });
    }
  }

  sheet.deleteRow(sheetRow);
  return jsonResponse({ success: true });
}

// ── Inventory Logic ───────────────────────────────────────────────

function getInventory() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);
  ensureInventoryHeader(sheet);

  var last = sheet.getLastRow();
  if (last < 2) {
    return jsonResponse({ success: true, inventory: [] });
  }

  var rows = sheet.getRange(2, 1, last - 1, 6).getValues();
  var inventory = rows.map(function(r) {
    return { item: r[0], unit: r[1], beginQty: r[2], withdrawal: r[3], balance: r[4], remarks: r[5] };
  }).filter(function(r) { return r.item !== ''; });

  return jsonResponse({ success: true, inventory: inventory });
}

function addInventoryRow(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);
  ensureInventoryHeader(sheet);

  sheet.appendRow([
    data.item       || '',
    data.unit       || '',
    data.beginQty   != null ? data.beginQty   : 0,
    data.withdrawal != null ? data.withdrawal : 0,
    data.balance    != null ? data.balance    : 0,
    data.remarks    || ''
  ]);
  return jsonResponse({ success: true });
}

function updateInventoryRow(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);
  ensureInventoryHeader(sheet);

  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return jsonResponse({ success: false, error: 'Invalid row index' });
  }

  sheet.getRange(sheetRow, 1, 1, 6).setValues([[
    data.item       || '',
    data.unit       || '',
    data.beginQty   != null ? data.beginQty   : 0,
    data.withdrawal != null ? data.withdrawal : 0,
    data.balance    != null ? data.balance    : 0,
    data.remarks    || ''
  ]]);
  return jsonResponse({ success: true });
}

function deleteInventoryRow(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);

  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return jsonResponse({ success: false, error: 'Invalid row index' });
  }

  sheet.deleteRow(sheetRow);
  return jsonResponse({ success: true });
}

function ensureInventoryHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    var header = ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS'];
    sheet.appendRow(header);
    styleHeaderRow(sheet, header.length);
  }
}

// ── Transactions Fetch ────────────────────────────────────────────

function getTransactions() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ success: true, transactions: [] });
  }

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var txns = rows.map(function(r) {
    return { orderNum: r[0], time: r[1], items: r[2], total: r[4], paid: 0, change: 0 };
  });
  return jsonResponse({ success: true, transactions: txns });
}

// ── Core Order Logic ─────────────────────────────────────────────

function saveOrder(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ordersSheet = getOrCreateSheet(ss, SHEET_NAME_ORDERS);
  ensureOrdersHeader(ordersSheet);

  var orderRow = [
    data.orderNum                             || '',
    formatDateTime(data.timestamp)            || '',
    data.items ? data.items.length : 0,
    data.items ? data.items.reduce(function(s, i){ return s + i.qty; }, 0) : 0,
    data.total != null ? data.total : 0
  ];
  ordersSheet.appendRow(orderRow);

  var itemsSheet = getOrCreateSheet(ss, SHEET_NAME_ITEMS);
  ensureItemsHeader(itemsSheet);

  if (data.items && data.items.length > 0) {
    data.items.forEach(function(item) {
      itemsSheet.appendRow([
        data.orderNum     || '',
        formatDateTime(data.timestamp) || '',
        item.name         || '',
        item.qty          || 0,
        item.price        || 0,
        item.subtotal     || 0
      ]);
    });
  }

  autoFormatSheets(ordersSheet, itemsSheet);
}

// ── Sheet Helpers ─────────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureOrdersHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    var header = ['Order #', 'Date & Time', 'Line Items', 'Total Qty', 'Grand Total (₱)'];
    sheet.appendRow(header);
    styleHeaderRow(sheet, header.length);
  }
}

function ensureItemsHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    var header = ['Order #', 'Date & Time', 'Product', 'Qty', 'Unit Price (₱)', 'Subtotal (₱)'];
    sheet.appendRow(header);
    styleHeaderRow(sheet, header.length);
  }
}

function styleHeaderRow(sheet, numCols) {
  var range = sheet.getRange(1, 1, 1, numCols);
  range
    .setBackground('#1a1a2e')
    .setFontColor('#f5c842')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function autoFormatSheets(ordersSheet, itemsSheet) {
  try {
    ordersSheet.autoResizeColumns(1, 5);
    itemsSheet.autoResizeColumns(1, 6);
  } catch(e) { /* ignore */ }
}

// ── Helpers ───────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateTime(isoString) {
  if (!isoString) return new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });
  try {
    var d = new Date(isoString);
    return Utilities.formatDate(d, 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss');
  } catch(e) {
    return isoString;
  }
}

// ══════════════════════════════════════════════════════════════════
//  OPTIONAL: Daily Summary Email
// ══════════════════════════════════════════════════════════════════

/*
function sendDailySummary() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return;

  var today = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
  var data   = sheet.getDataRange().getValues();
  var todayOrders = data.slice(1).filter(function(row) {
    return String(row[1]).startsWith(today);
  });

  var totalRevenue = todayOrders.reduce(function(s, r) { return s + (r[4] || 0); }, 0);

  var body = 'AJ Medina POS Daily Summary\n\n' +
             'Date: ' + today + '\n' +
             'Orders: ' + todayOrders.length + '\n' +
             'Revenue: ₱' + totalRevenue.toFixed(2) + '\n\n' +
             'View full report: ' + ss.getUrl();

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'AJ Medina POS — Daily Summary ' + today,
    body: body
  });
}
*/
