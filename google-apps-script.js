/**
 * BoxTrack — Google Apps Script Backend  (v2)
 * =============================================
 * Deploy as a Web App:
 *   1. Paste this into script.google.com
 *   2. Deploy → New deployment → Web app
 *      Execute as: Me | Access: Anyone
 *   3. Copy the exec URL → paste into index.html CONFIG.GAS_URL
 *
 * Google Sheet: one tab named "Inventory"
 * Columns: A=BoxID | B=Location | C=SKU | D=ProductName | E=Qty
 * Row 1 = header (skipped automatically)
 */

// ─── CONFIG ───────────────────────────────────────────────────────────────────
var SHEET_NAME = 'Inventory';
var COL = { BOX: 1, LOCATION: 2, SKU: 3, NAME: 4, QTY: 5 }; // 1-indexed

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found. Run seedSampleData() first or rename your tab to "Inventory".');
  return sheet;
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET ROUTER ───────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : null;

    // ── GET ?action=all  →  return every row (used by client-side search)
    if (action === 'all') return handleGetAll();

    // ── GET ?box=BOX-01  →  return products for a specific box
    var boxId = (e.parameter && e.parameter.box)
      ? e.parameter.box.toString().trim().toUpperCase()
      : null;

    if (boxId) return handleGetBox(boxId);

    return jsonOut({ error: 'Missing parameter: use ?box=BOX-01 or ?action=all' });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

/**
 * GET ?box=BOX-01
 * Returns: { boxId, location, products: [{sku, name, qty}] }
 */
function handleGetBox(boxId) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var products = [];
  var location = '';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.BOX - 1]) continue; // skip empty rows
    var rowBox = String(row[COL.BOX - 1]).trim().toUpperCase();
    if (rowBox === boxId) {
      if (!location) location = String(row[COL.LOCATION - 1]).trim();
      products.push({
        sku:  String(row[COL.SKU  - 1]).trim(),
        name: String(row[COL.NAME - 1]).trim(),
        qty:  Number(row[COL.QTY  - 1]) || 0,
      });
    }
  }

  return jsonOut({ boxId: boxId, location: location || 'Unknown', products: products });
}

/**
 * GET ?action=all
 * Returns: { rows: [{boxId, location, sku, name, qty}] }
 * Used by the client-side search to avoid per-keystroke server calls.
 */
function handleGetAll() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.BOX - 1]) continue;
    rows.push({
      boxId:    String(row[COL.BOX      - 1]).trim().toUpperCase(),
      location: String(row[COL.LOCATION - 1]).trim(),
      sku:      String(row[COL.SKU      - 1]).trim(),
      name:     String(row[COL.NAME     - 1]).trim(),
      qty:      Number(row[COL.QTY      - 1]) || 0,
    });
  }

  return jsonOut({ rows: rows });
}

// ─── POST ROUTER ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'update'; // backwards-compatible default

    if (action === 'update')     return handleUpdate(body);
    if (action === 'addProduct') return handleAddProduct(body);

    return jsonOut({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

/**
 * POST { action:'update', boxId, updates:[{sku, newQty}] }
 * Updates the Qty column for existing rows.
 * Returns: { success:true, updated:N }
 */
function handleUpdate(body) {
  var boxId = String(body.boxId).trim().toUpperCase();
  var updates = body.updates || [];
  if (!boxId || updates.length === 0) return jsonOut({ error: 'Invalid payload' });

  var updateMap = {};
  updates.forEach(function (u) { updateMap[String(u.sku).trim()] = Number(u.newQty); });

  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowBox = String(row[COL.BOX - 1]).trim().toUpperCase();
    var rowSku = String(row[COL.SKU - 1]).trim();
    if (rowBox === boxId && updateMap.hasOwnProperty(rowSku)) {
      sheet.getRange(i + 1, COL.QTY).setValue(updateMap[rowSku]);
      count++;
    }
  }

  return jsonOut({ success: true, updated: count });
}

/**
 * POST { action:'addProduct', boxId, location, sku, name, qty }
 * Appends a new row for the product.
 * Returns: { success:true }
 */
function handleAddProduct(body) {
  var boxId    = String(body.boxId   || '').trim().toUpperCase();
  var location = String(body.location|| '').trim();
  var sku      = String(body.sku     || '').trim().toUpperCase();
  var name     = String(body.name    || '').trim();
  var qty      = Number(body.qty) || 0;

  if (!boxId || !sku || !name) return jsonOut({ error: 'boxId, sku and name are required' });

  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();

  // Reject duplicate SKU within same box
  for (var i = 1; i < data.length; i++) {
    var rowBox = String(data[i][COL.BOX - 1]).trim().toUpperCase();
    var rowSku = String(data[i][COL.SKU - 1]).trim().toUpperCase();
    if (rowBox === boxId && rowSku === sku) {
      return jsonOut({ error: 'SKU ' + sku + ' already exists in ' + boxId });
    }
    // Infer location from existing rows for this box if not provided
    if (!location && rowBox === boxId) {
      location = String(data[i][COL.LOCATION - 1]).trim();
    }
  }

  // Append new row
  sheet.appendRow([boxId, location, sku, name, qty]);

  return jsonOut({ success: true });
}

// ─── SEED SAMPLE DATA (run once from the editor) ──────────────────────────────
/**
 * Run this function ONCE from the Apps Script editor to create sample data.
 * It will:
 *   1. Create (or clear) the "Inventory" sheet
 *   2. Populate it with demo boxes and products
 */
function seedSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clearContents();

  var headers = ['Box ID', 'Location', 'SKU', 'Product Name', 'Qty'];
  var rows = [
    ['BOX-01', 'Shelf A1', 'SKU-1001', 'Widget Alpha',       42],
    ['BOX-01', 'Shelf A1', 'SKU-1002', 'Widget Beta',         3],
    ['BOX-01', 'Shelf A1', 'SKU-1003', 'Connector Type-C',   18],
    ['BOX-02', 'Shelf B3', 'SKU-2001', 'Gadget Pro',         17],
    ['BOX-02', 'Shelf B3', 'SKU-2002', 'Gadget Lite',         0],
    ['BOX-03', 'Shelf C2', 'SKU-3001', 'Sensor Module X',    25],
    ['BOX-03', 'Shelf C2', 'SKU-3002', 'Sensor Module Y',     4],
    ['BOX-04', 'Shelf D1', 'SKU-4001', 'Power Adapter 12V',  10],
    ['BOX-04', 'Shelf D1', 'SKU-4002', 'USB Hub 4-Port',      7],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4f46e5');
  headerRange.setFontColor('#ffffff');

  for (var c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);

  Logger.log('seedSampleData: ' + rows.length + ' rows written to "' + SHEET_NAME + '".');
}
