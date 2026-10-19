/**
 * BoxTrack — Google Apps Script Backend
 * ======================================
 * Deploy as a Web App:
 *   1. Open Apps Script (script.google.com), paste this file.
 *   2. Click "Deploy" → "New deployment" → type: "Web app"
 *   3. Execute as: Me | Who has access: Anyone
 *   4. Copy the deployment URL and paste it into index.html → CONFIG.GAS_URL
 *
 * Google Sheet structure (one sheet named "Inventory"):
 * ┌──────────┬──────────────┬─────────────┬────────────────────────┬─────┐
 * │  Box ID  │  Shelf/Loc   │  SKU        │  Product Name          │ Qty │
 * ├──────────┼──────────────┼─────────────┼────────────────────────┼─────┤
 * │  BOX-01  │  Shelf A1    │  SKU-1001   │  Widget Alpha          │  42 │
 * │  BOX-01  │  Shelf A1    │  SKU-1002   │  Widget Beta           │   3 │
 * │  BOX-02  │  Shelf B3    │  SKU-2001   │  Gadget Pro            │  17 │
 * └──────────┴──────────────┴─────────────┴────────────────────────┴─────┘
 *
 * Row 1 = header row (not used in data processing, just for your reference).
 * Columns: A=BoxID, B=Location, C=SKU, D=ProductName, E=Qty
 */

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const SHEET_NAME = 'Inventory';
const COL = { BOX: 1, LOCATION: 2, SKU: 3, NAME: 4, QTY: 5 }; // 1-indexed

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function jsonResponse(data, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function corsHeaders() {
  // Note: GAS doesn't fully support custom response headers for CORS,
  // but returning JSON with Content-Type is usually sufficient for
  // fetch() calls with mode:'cors' or no-cors.
  return {};
}

// ─── GET — Retrieve all products for a Box ID ─────────────────────────────────
/**
 * Called by the web app when the user scans a QR code:
 *   GET ?box=BOX-01
 *
 * Returns:
 * {
 *   "boxId": "BOX-01",
 *   "location": "Shelf A1",
 *   "products": [
 *     { "sku": "SKU-1001", "name": "Widget Alpha", "qty": 42 },
 *     { "sku": "SKU-1002", "name": "Widget Beta",  "qty": 3  }
 *   ]
 * }
 */
function doGet(e) {
  try {
    var boxId = (e.parameter && e.parameter.box)
      ? e.parameter.box.toString().trim().toUpperCase()
      : null;

    if (!boxId) {
      return jsonResponse({ error: 'Missing ?box= parameter' });
    }

    var sheet = getSheet();
    if (!sheet) {
      return jsonResponse({ error: 'Sheet "' + SHEET_NAME + '" not found' });
    }

    var data = sheet.getDataRange().getValues();
    var products = [];
    var location = '';

    // Skip header row (index 0)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
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

    return jsonResponse({
      boxId:    boxId,
      location: location || 'Unknown',
      products: products,
    });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── POST — Update stock quantities ──────────────────────────────────────────
/**
 * Called when the user taps "Confirm Changes".
 *
 * Request body (JSON string, Content-Type: text/plain to avoid CORS preflight):
 * {
 *   "boxId": "BOX-01",
 *   "updates": [
 *     { "sku": "SKU-1001", "newQty": 40 },
 *     { "sku": "SKU-1002", "newQty": 5  }
 *   ]
 * }
 *
 * Returns:
 * { "success": true, "updated": 2 }
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var boxId  = String(body.boxId).trim().toUpperCase();
    var updates = body.updates || []; // [{ sku, newQty }]

    if (!boxId || updates.length === 0) {
      return jsonResponse({ error: 'Invalid payload' });
    }

    // Build a lookup map from SKU → newQty
    var updateMap = {};
    updates.forEach(function(u) {
      updateMap[String(u.sku).trim()] = Number(u.newQty);
    });

    var sheet = getSheet();
    if (!sheet) {
      return jsonResponse({ error: 'Sheet "' + SHEET_NAME + '" not found' });
    }

    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowBox = String(row[COL.BOX - 1]).trim().toUpperCase();
      var rowSku = String(row[COL.SKU - 1]).trim();

      if (rowBox === boxId && updateMap.hasOwnProperty(rowSku)) {
        var newQty = updateMap[rowSku];
        // Write to column E (COL.QTY = 5), row i+1 (1-indexed, skip header offset)
        sheet.getRange(i + 1, COL.QTY).setValue(newQty);
        updatedCount++;
      }
    }

    return jsonResponse({ success: true, updated: updatedCount });

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ─── OPTIONAL: Seed sample data ───────────────────────────────────────────────
/**
 * Run this function once manually from the Apps Script editor to create
 * a sample "Inventory" sheet with demo data. Safe to skip if you already
 * have your sheet set up.
 */
function seedSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  sheet.clearContents();

  var headers = ['Box ID', 'Location', 'SKU', 'Product Name', 'Qty'];
  var rows = [
    ['BOX-01', 'Shelf A1', 'SKU-1001', 'Widget Alpha',      42],
    ['BOX-01', 'Shelf A1', 'SKU-1002', 'Widget Beta',        3],
    ['BOX-01', 'Shelf A1', 'SKU-1003', 'Connector Type-C',  18],
    ['BOX-02', 'Shelf B3', 'SKU-2001', 'Gadget Pro',        17],
    ['BOX-02', 'Shelf B3', 'SKU-2002', 'Gadget Lite',        0],
    ['BOX-03', 'Shelf C2', 'SKU-3001', 'Sensor Module X',   25],
    ['BOX-03', 'Shelf C2', 'SKU-3002', 'Sensor Module Y',    4],
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4f46e5');
  headerRange.setFontColor('#ffffff');

  // Auto-resize columns
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }

  Logger.log('Sample data seeded successfully.');
}
