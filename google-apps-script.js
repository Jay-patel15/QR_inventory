/**
 * BoxTrack — Google Apps Script Backend (No-SKU Version)
 * ======================================================
 * Deploy as a Web App:
 *   1. Paste this into script.google.com (or Extensions > Apps Script)
 *   2. Click Save (Ctrl+S)
 *   3. Deploy → Manage deployments → Edit (Pencil) → Version: New version → Deploy
 *      (Or Deploy → New deployment → Web app → Execute as: Me | Access: Anyone)
 *   4. Make sure the exec URL matches CONFIG.GAS_URL in index.html
 *
 * Google Sheet: Tab named "Inventory"
 * Recommended Columns: A=Box ID | B=Location | C=Product Name | D=Qty
 * (Also auto-detects legacy 5-column sheets with SKU if present)
 */

// ─── CONFIG ───────────────────────────────────────────────────────────────────
var SHEET_NAME = 'Inventory';

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found. Please rename your tab to "' + SHEET_NAME + '" or run seedSampleData().');
  }
  return sheet;
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Dynamically resolves column positions from Row 1 headers.
 * Supports both 4-column (Box ID, Location, Product Name, Qty)
 * and legacy 5-column (Box ID, Location, SKU, Product Name, Qty).
 */
function getColumnIndices(sheet) {
  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 4)).getValues()[0];
  var cols = { box: 1, location: 2, name: 3, qty: 4, hasSku: false, sku: -1 };

  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i] || '').trim().toLowerCase();
    if (h.includes('box')) cols.box = i + 1;
    else if (h.includes('loc')) cols.location = i + 1;
    else if (h.includes('sku')) { cols.hasSku = true; cols.sku = i + 1; }
    else if (h.includes('name') || h.includes('product') || h.includes('item')) cols.name = i + 1;
    else if (h.includes('qty') || h.includes('quant') || h.includes('count') || h.includes('stock')) cols.qty = i + 1;
  }
  return cols;
}

// ─── GET ROUTER ───────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? String(e.parameter.action).trim().toLowerCase() : null;

    // ── GET ?action=all  →  returns every product (used by client-side search)
    if (action === 'all') return handleGetAll();

    // ── GET ?box=BOX-01  →  returns products for a specific box
    var boxId = (e.parameter && e.parameter.box)
      ? String(e.parameter.box).trim().toUpperCase()
      : null;

    if (boxId) return handleGetBox(boxId);

    return jsonOut({ error: 'Missing parameter: use ?box=BOX-01 or ?action=all' });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

/**
 * GET ?box=BOX-01
 * Returns: { boxId, location, products: [{ name, qty }] }
 */
function handleGetBox(boxId) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndices(sheet);
  var products = [];
  var location = '';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowBox = String(row[cols.box - 1] || '').trim().toUpperCase();
    if (rowBox === boxId) {
      if (!location) location = String(row[cols.location - 1] || '').trim();
      var prodName = String(row[cols.name - 1] || '').trim();
      if (prodName) {
        products.push({
          name: prodName,
          qty:  Number(row[cols.qty - 1]) || 0,
        });
      }
    }
  }

  return jsonOut({ boxId: boxId, location: location || 'Unknown', products: products });
}

/**
 * GET ?action=all
 * Returns: { rows: [{ boxId, location, name, qty }] }
 */
function handleGetAll() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndices(sheet);
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowBox = String(row[cols.box - 1] || '').trim().toUpperCase();
    var prodName = String(row[cols.name - 1] || '').trim();
    if (!rowBox || !prodName) continue;

    rows.push({
      boxId:    rowBox,
      location: String(row[cols.location - 1] || '').trim(),
      name:     prodName,
      qty:      Number(row[cols.qty - 1]) || 0,
    });
  }

  return jsonOut({ rows: rows });
}

// ─── POST ROUTER ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ error: 'Empty POST request body' });
    }

    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'update';

    if (action === 'update')     return handleUpdate(body);
    if (action === 'addProduct') return handleAddProduct(body);

    return jsonOut({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

/**
 * POST { action:'update', boxId, updates:[{ name, newQty }] }
 * Updates the Qty column matching boxId and product name.
 * Returns: { success:true, updated:N }
 */
function handleUpdate(body) {
  var boxId = String(body.boxId || '').trim().toUpperCase();
  var updates = body.updates || [];
  if (!boxId || updates.length === 0) {
    return jsonOut({ error: 'Invalid payload: boxId and updates array required' });
  }

  var updateMap = {};
  updates.forEach(function (u) {
    var key = String(u.name || '').trim().toLowerCase();
    if (key) updateMap[key] = Number(u.newQty);
  });

  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndices(sheet);
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowBox = String(row[cols.box - 1] || '').trim().toUpperCase();
    var rowName = String(row[cols.name - 1] || '').trim().toLowerCase();

    if (rowBox === boxId && updateMap.hasOwnProperty(rowName)) {
      sheet.getRange(i + 1, cols.qty).setValue(updateMap[rowName]);
      count++;
    }
  }

  return jsonOut({ success: true, updated: count });
}

/**
 * POST { action:'addProduct', boxId, location, name, qty }
 * Appends a new product row.
 * Returns: { success:true }
 */
function handleAddProduct(body) {
  var boxId    = String(body.boxId    || '').trim().toUpperCase();
  var location = String(body.location || '').trim();
  var name     = String(body.name     || '').trim();
  var qty      = Number(body.qty) || 0;

  if (!boxId || !name) {
    return jsonOut({ error: 'boxId and name are required' });
  }

  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var cols = getColumnIndices(sheet);

  // Reject duplicate product name in the same box
  var nameLower = name.toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowBox = String(data[i][cols.box - 1] || '').trim().toUpperCase();
    var rowName = String(data[i][cols.name - 1] || '').trim().toLowerCase();
    if (rowBox === boxId && rowName === nameLower) {
      return jsonOut({ error: 'Product "' + name + '" already exists in ' + boxId });
    }
    // Infer location if not passed
    if (!location && rowBox === boxId) {
      location = String(data[i][cols.location - 1] || '').trim();
    }
  }

  // Construct new row according to sheet columns
  if (cols.hasSku) {
    var newRow = [];
    newRow[cols.box - 1] = boxId;
    newRow[cols.location - 1] = location;
    newRow[cols.sku - 1] = '-';
    newRow[cols.name - 1] = name;
    newRow[cols.qty - 1] = qty;
    sheet.appendRow(newRow);
  } else {
    sheet.appendRow([boxId, location, name, qty]);
  }

  return jsonOut({ success: true });
}

// ─── SEED SAMPLE DATA (run once from Apps Script editor) ──────────────────────
/**
 * Run this function from Apps Script editor if starting fresh.
 * Creates clean 4-column Inventory sheet: Box ID | Location | Product Name | Qty
 */
function seedSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clearContents();

  var headers = ['Box ID', 'Location', 'Product Name', 'Qty'];
  var rows = [
    ['BOX-01', 'Shelf A1', 'Widget Alpha',       42],
    ['BOX-01', 'Shelf A1', 'Widget Beta',         3],
    ['BOX-01', 'Shelf A1', 'Connector Type-C',   18],
    ['BOX-02', 'Shelf B3', 'Gadget Pro',         17],
    ['BOX-02', 'Shelf B3', 'Gadget Lite',         0],
    ['BOX-03', 'Shelf C2', 'Sensor Module X',    25],
    ['BOX-03', 'Shelf C2', 'Sensor Module Y',     4],
    ['BOX-04', 'Shelf D1', 'Power Adapter 12V',  10],
    ['BOX-04', 'Shelf D1', 'USB Hub 4-Port',      7],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4f46e5');
  headerRange.setFontColor('#ffffff');

  for (var c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);

  Logger.log('seedSampleData: ' + rows.length + ' rows created in "' + SHEET_NAME + '".');
}
