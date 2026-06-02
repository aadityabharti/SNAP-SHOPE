/**
 * SnapShope — Google Apps Script Backend
 * ═══════════════════════════════════════════════════════════════
 * This script turns your Google Sheet into a live database for
 * the SnapShope e-commerce website.
 *
 * SETUP (one-time, ~5 minutes):
 * ──────────────────────────────────────────────────────────────
 * 1. Open Google Sheets → create a new spreadsheet
 *    Name it: "SnapShope Database"
 *
 * 2. In the spreadsheet, go to:
 *    Extensions → Apps Script
 *
 * 3. Delete everything in the editor and paste THIS ENTIRE FILE
 *
 * 4. Click the floppy-disk icon (Save) — name the project "SnapShope DB"
 *
 * 5. Click "Deploy" (top right) → "New Deployment"
 *    - Type: Web App
 *    - Description: SnapShope Live DB v1
 *    - Execute as: Me (your Google account)
 *    - Who has access: Anyone          ← IMPORTANT
 *    Click "Deploy" → copy the Web App URL
 *
 * 6. Open your SnapShope website → Profile → Overview
 *    Paste the URL into the "Google Sheets Database" box → click Connect
 *
 * Done! Every signup, login, order and cart update will now
 * appear in your Google Sheet in real time.
 * ═══════════════════════════════════════════════════════════════
 */

// ── Sheet names ────────────────────────────────────────────────
const SHEET_NAMES = {
  Users:         'Users',
  Orders:        'Orders',
  CartSnapshots: 'CartSnapshots',
  Logs:          'Logs',
};

// ── Column definitions ─────────────────────────────────────────
const COLUMNS = {
  Users: [
    'User ID', 'First Name', 'Last Name', 'Email', 'Phone',
    'Date of Birth', 'Joined At', 'Last Event', 'Last Event At',
    'Total Orders', 'Total Spent (₹)', 'Cart Items',
    'Cart Summary', 'Saved Addresses',
  ],
  Orders: [
    'Order ID', 'Order Date', 'User ID', 'Customer Name',
    'Email', 'Phone', 'Product ID', 'Product Name', 'Category',
    'Unit Price (₹)', 'Quantity', 'Item Total (₹)',
    'Order Total (₹)', 'Payment Method', 'Status',
    'Delivery Name', 'Delivery City', 'Delivery State', 'Delivery PIN',
  ],
  CartSnapshots: [
    'Snapshot Time', 'User ID', 'Customer Name', 'Email',
    'Items', 'Total Items', 'Cart Value (₹)',
  ],
  Logs: [
    'Timestamp', 'Sheet', 'Action', 'Email / Key', 'Status', 'Notes',
  ],
};

// ── Field → column mapping ─────────────────────────────────────
const FIELD_MAP = {
  Users: {
    userId:'User ID', firstName:'First Name', lastName:'Last Name',
    email:'Email', phone:'Phone', dob:'Date of Birth',
    joinedAt:'Joined At', lastEvent:'Last Event', lastEventAt:'Last Event At',
    totalOrders:'Total Orders', totalSpent:'Total Spent (₹)',
    cartItemCount:'Cart Items', cartSummary:'Cart Summary',
    savedAddresses:'Saved Addresses',
  },
  Orders: {
    orderId:'Order ID', orderDate:'Order Date', userId:'User ID',
    customerName:'Customer Name', email:'Email', phone:'Phone',
    productId:'Product ID', productName:'Product Name', category:'Category',
    unitPrice:'Unit Price (₹)', quantity:'Quantity', itemTotal:'Item Total (₹)',
    orderTotal:'Order Total (₹)', paymentMethod:'Payment Method', status:'Status',
    deliveryName:'Delivery Name', deliveryCity:'Delivery City',
    deliveryState:'Delivery State', deliveryPin:'Delivery PIN',
  },
  CartSnapshots: {
    snapshotTime:'Snapshot Time', userId:'User ID', userName:'Customer Name',
    email:'Email', items:'Items', totalItems:'Total Items', cartValue:'Cart Value (₹)',
  },
};

// ══════════════════════════════════════════════════════════════════
// ENTRY POINT — handles all POST requests from the website
// ══════════════════════════════════════════════════════════════════
function doPost(e) {
  const cors = ContentService
    .createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const payload = JSON.parse(e.postData.contents);
    const { sheet: sheetName, action, key, data } = payload;

    // Ping / test request
    if (action === 'ping') {
      cors.setContent(JSON.stringify({ ok: true, msg: 'SnapShope DB is live!' }));
      return cors;
    }

    // Validate
    if (!sheetName || !action || !data) throw new Error('Missing sheet, action or data');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, sheetName);

    let result;
    if (action === 'upsert') {
      result = upsertRow(sheet, sheetName, key, data);
    } else if (action === 'append') {
      result = appendRow(sheet, sheetName, data);
    } else {
      throw new Error('Unknown action: ' + action);
    }

    appendLog(ss, sheetName, action, data[key] || data.email || '—', 'OK', '');
    cors.setContent(JSON.stringify({ ok: true, result }));

  } catch (err) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      appendLog(ss, 'Error', 'error', '—', 'FAIL', err.message);
    } catch(_) {}
    cors.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return cors;
}

// Also allow GET for quick ping from browser address bar
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'SnapShope DB — use POST to send data' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════
// SHEET HELPERS
// ══════════════════════════════════════════════════════════════════

/**
 * Get sheet by name, or create it with the correct header row.
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = COLUMNS[sheetName];
    if (headers) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      styleHeaderRow(headerRange, sheetName);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Upsert: find a row by `key` column value and update it,
 * or append a new row if not found.
 */
function upsertRow(sheet, sheetName, key, data) {
  const headers = getHeaders(sheet);
  const colMap  = buildColMap(headers);
  const keyCol  = colMap[FIELD_MAP[sheetName]?.[key] || key];

  if (keyCol === undefined) {
    // Can't find key column — just append
    return appendRow(sheet, sheetName, data);
  }

  const lastRow = sheet.getLastRow();
  let targetRow = -1;

  if (lastRow >= 2) {
    const keyValues = sheet.getRange(2, keyCol + 1, lastRow - 1, 1).getValues();
    const searchVal = String(data[key] || '').toLowerCase();
    for (let r = 0; r < keyValues.length; r++) {
      if (String(keyValues[r][0]).toLowerCase() === searchVal) {
        targetRow = r + 2; // +2: 1-indexed + header row
        break;
      }
    }
  }

  const rowData = buildRowArray(headers, sheetName, data);

  if (targetRow > 0) {
    // Update existing row
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    styleDataRow(sheet, targetRow, lastRow);
    return 'updated row ' + targetRow;
  } else {
    // Append new row
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    styleDataRow(sheet, newRow, newRow);
    return 'appended row ' + newRow;
  }
}

/**
 * Always append a new row.
 */
function appendRow(sheet, sheetName, data) {
  const headers = getHeaders(sheet);
  const rowData = buildRowArray(headers, sheetName, data);
  const newRow  = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  styleDataRow(sheet, newRow, newRow);
  return 'appended row ' + newRow;
}

/**
 * Map data object to a flat array matching the sheet's header columns.
 */
function buildRowArray(headers, sheetName, data) {
  const fieldMap = FIELD_MAP[sheetName] || {};
  // Build reverse map: header label → field name in data
  const reverseMap = {};
  Object.entries(fieldMap).forEach(([field, label]) => { reverseMap[label] = field; });

  return headers.map(header => {
    const field = reverseMap[header] || Object.keys(data).find(k => k === header);
    const val   = field ? data[field] : '';
    return val !== undefined && val !== null ? val : '';
  });
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function buildColMap(headers) {
  const m = {};
  headers.forEach((h, i) => { m[h] = i; });
  return m;
}

// ══════════════════════════════════════════════════════════════════
// LOGGING
// ══════════════════════════════════════════════════════════════════
function appendLog(ss, sheetName, action, key, status, notes) {
  const logSheet = getOrCreateSheet(ss, 'Logs');
  logSheet.appendRow([
    new Date().toLocaleString('en-IN'),
    sheetName, action, key, status, notes
  ]);
}

// ══════════════════════════════════════════════════════════════════
// STYLING
// ══════════════════════════════════════════════════════════════════
const HEADER_COLORS = {
  Users:         '#E8180A',
  Orders:        '#1a7a4a',
  CartSnapshots: '#d97706',
  Logs:          '#374151',
};

function styleHeaderRow(range, sheetName) {
  const bg = HEADER_COLORS[sheetName] || '#0A0A0A';
  range.setBackground(bg)
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setFontFamily('Arial')
       .setFontSize(9)
       .setHorizontalAlignment('center')
       .setWrap(true);
}

function styleDataRow(sheet, rowNum, lastRow) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const bg = rowNum % 2 === 0 ? '#FAFAFA' : '#FFFFFF';
  const range = sheet.getRange(rowNum, 1, 1, lastCol);
  range.setBackground(bg)
       .setFontFamily('Arial')
       .setFontSize(9)
       .setFontColor('#0A0A0A');

  // Auto-resize columns periodically (every 10 rows to avoid slow ops)
  if (lastRow % 10 === 0) {
    for (let c = 1; c <= lastCol; c++) {
      sheet.autoResizeColumn(c);
    }
  }
}
