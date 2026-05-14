// ══════════════════════════════════════════════════════════════════
//  AJ Medina POS — Google Apps Script Backend (code.gs)
//  Logs POS orders to a Google Sheet
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
// ══════════════════════════════════════════════════════════════════

// ── Config ──────────────────────────────────────────────────────
var SHEET_NAME_ORDERS    = 'Orders';
var SHEET_NAME_ITEMS     = 'Order Items';
var SHEET_NAME_INVENTORY = 'Inventory';

// ── Entry Points ─────────────────────────────────────────────────

/**
 * GET  → health check / simple ping, or fetch inventory/transactions
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';

  if (action === 'inventory') {
    return getInventory();
  }
  if (action === 'transactions') {
    return getTransactions();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', app: 'AJ Medina POS', time: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST → receive order or inventory action from POS frontend
 */
function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    // Route by action field
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

// ── Inventory Logic ───────────────────────────────────────────────

/**
 * Returns all rows from the Inventory sheet as JSON.
 */
function getInventory() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);
  ensureInventoryHeader(sheet);

  var last = sheet.getLastRow();
  if (last < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, inventory: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var rows = sheet.getRange(2, 1, last - 1, 6).getValues();
  var inventory = rows.map(function(r) {
    return {
      item:       r[0],
      unit:       r[1],
      beginQty:   r[2],
      withdrawal: r[3],
      balance:    r[4],
      remarks:    r[5]
    };
  }).filter(function(r) { return r.item !== ''; });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, inventory: inventory }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Appends a new row to the Inventory sheet.
 */
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

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Updates an existing row in the Inventory sheet (0-based rowIndex from frontend).
 */
function updateInventoryRow(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);
  ensureInventoryHeader(sheet);

  // rowIndex is 0-based from the frontend (maps to sheet row = rowIndex + 2)
  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Invalid row index' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  sheet.getRange(sheetRow, 1, 1, 6).setValues([[
    data.item       || '',
    data.unit       || '',
    data.beginQty   != null ? data.beginQty   : 0,
    data.withdrawal != null ? data.withdrawal : 0,
    data.balance    != null ? data.balance    : 0,
    data.remarks    || ''
  ]]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Deletes a row from the Inventory sheet (0-based rowIndex from frontend).
 */
function deleteInventoryRow(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY);

  var sheetRow = parseInt(data.rowIndex) + 2;
  if (isNaN(sheetRow) || sheetRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Invalid row index' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  sheet.deleteRow(sheetRow);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureInventoryHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    var header = ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS'];
    sheet.appendRow(header);
    styleHeaderRow(sheet, header.length);
  }
}

// ── Transactions Fetch ────────────────────────────────────────────

/**
 * Returns rows from the Orders sheet as JSON for the Transactions tab.
 */
function getTransactions() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, transactions: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var txns = rows.map(function(r) {
    return {
      orderNum: r[0],
      time:     r[1],
      items:    r[2],
      total:    r[4],
      paid:     0,
      change:   0
    };
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, transactions: txns }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Core Logic ───────────────────────────────────────────────────

/**
 * Writes the order to both the Orders and Order Items sheets.
 * Creates and formats sheets if they don't exist yet.
 */
function saveOrder(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Orders sheet ──
  var ordersSheet = getOrCreateSheet(ss, SHEET_NAME_ORDERS);
  ensureOrdersHeader(ordersSheet);

  var orderRow = [
    data.orderNum                             || '',          // A: Order #
    formatDateTime(data.timestamp)            || '',          // B: Date/Time
    data.items ? data.items.length : 0,                       // C: # of line items
    data.items ? data.items.reduce(function(s, i){ return s + i.qty; }, 0) : 0, // D: total qty
    data.total != null ? data.total : 0                       // E: Grand Total
  ];
  ordersSheet.appendRow(orderRow);

  // ── Order Items sheet ──
  var itemsSheet = getOrCreateSheet(ss, SHEET_NAME_ITEMS);
  ensureItemsHeader(itemsSheet);

  if (data.items && data.items.length > 0) {
    data.items.forEach(function(item) {
      itemsSheet.appendRow([
        data.orderNum     || '',     // A: Order #
        formatDateTime(data.timestamp) || '',  // B: Date/Time
        item.name         || '',     // C: Product
        item.qty          || 0,      // D: Qty
        item.price        || 0,      // E: Unit Price
        item.subtotal     || 0       // F: Subtotal
      ]);
    });
  }

  // ── Auto-format the sheets (first time) ──
  autoFormatSheets(ordersSheet, itemsSheet);
}

// ── Sheet Helpers ─────────────────────────────────────────────────

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
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
  // Auto-resize columns
  try {
    ordersSheet.autoResizeColumns(1, 5);
    itemsSheet.autoResizeColumns(1, 6);
  } catch(e) { /* ignore */ }
}

// ── Date Formatting ───────────────────────────────────────────────

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
//  Uncomment and set up a time-driven trigger → onOpen()
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
