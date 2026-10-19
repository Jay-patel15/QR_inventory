# BoxTrack — QR Inventory System

A zero-dependency, mobile-first web app for tracking physical box inventory via QR codes, backed by Google Sheets.

---

## Project Structure

```
QR inventory/
├── index.html              ← Single-page app (Scan Mode + Generator Mode)
├── google-apps-script.js   ← Backend code — deploy to Google Apps Script
├── vercel.json             ← Vercel deployment config
└── README.md
```

---

## Quick Setup (5 steps)

### Step 1 — Create your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet.
2. Rename the default sheet tab to **`Inventory`** (exact name, case-sensitive).
3. The sheet columns should be:

| A — Box ID | B — Location | C — SKU | D — Product Name | E — Qty |
|------------|--------------|---------|------------------|---------|
| BOX-01     | Shelf A1     | SKU-001 | Widget Alpha     | 42      |

> **Shortcut:** In Apps Script, run `seedSampleData()` once to auto-populate demo data.

---

### Step 2 — Deploy the Google Apps Script

1. Open [script.google.com](https://script.google.com) → **New project**.
2. Paste the entire contents of `google-apps-script.js` into the editor.
3. Click **Extensions → Apps Script** (or just use the link above).
4. Click **Deploy → New deployment**:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy** and copy the **Web app URL** (it looks like `https://script.google.com/macros/s/ABC123.../exec`).

> ⚠️ Every time you edit the script, you must create a **new deployment** (not re-use the old one) for changes to take effect.

---

### Step 3 — Wire the URL into the app

Open `index.html` and find this section near the top of the `<script>` block:

```javascript
const CONFIG = {
  GAS_URL: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
  LOW_STOCK_THRESHOLD: 5,
};
```

Replace `YOUR_SCRIPT_ID` with your actual script ID from Step 2.

---

### Step 4 — Deploy to Vercel

```bash
# Install Vercel CLI (one-time)
npm i -g vercel

# From inside the project folder:
vercel

# Follow the prompts — choose default settings.
# After deploy, your live URL will be printed, e.g.:
#   https://qr-inventory-xyz.vercel.app
```

Or use the [Vercel Dashboard](https://vercel.com/new) → **Import Git Repository**.

---

### Step 5 — Generate QR codes

1. Visit your Vercel URL (no `?box=` param → **Generator Mode**).
2. Type a Box ID (e.g. `BOX-01`) and tap **Generate QR Code**.
3. Tap **Print QR Label** — it outputs a 62×38mm label (Brother QL / Dymo compatible).
4. Stick the label on the physical box.

---

## How It Works

### Scan Mode (`?box=BOX-01`)
- Triggered automatically when a QR code is scanned with any camera app.
- Fetches live data from Google Sheets via the Apps Script endpoint.
- Shows Box ID, shelf location, all products with current quantities.
- Amber = low stock (≤ 5), Red = out of stock.
- Tap **Checkout / Update Stock** to enter edit mode — tap `+`/`−` per product.
- Tap **Confirm Changes** → POSTs a JSON delta to Apps Script → updates Sheet in real time.

### Generator Mode (no `?box=` param)
- Enter any Box ID to generate its QR code.
- The QR encodes `https://your-app.vercel.app?box=BOX-01`.
- Print-optimised label output (62×38mm, high error-correction).

---

## Sheet Column Reference

| Column | Field        | Notes                                    |
|--------|--------------|------------------------------------------|
| A      | Box ID       | Uppercase, e.g. `BOX-01`                |
| B      | Location     | Shelf/aisle label, e.g. `Shelf A1`      |
| C      | SKU          | Unique product identifier                |
| D      | Product Name | Human-readable name                      |
| E      | Qty          | Numeric quantity (updated by app)        |

Row 1 is the header row and is skipped by the script.

---

## Customization

| What                        | Where                          |
|-----------------------------|--------------------------------|
| Low-stock threshold         | `CONFIG.LOW_STOCK_THRESHOLD`   |
| Sheet name                  | `SHEET_NAME` in GAS script     |
| Column layout               | `COL` object in GAS script     |
| Label size                  | Print CSS in `index.html`      |
| Accent color                | `--accent` CSS variable        |

---

## CORS Note

The POST request uses `Content-Type: text/plain` intentionally. This avoids CORS preflight (`OPTIONS`) requests, which Google Apps Script does not handle. The GAS endpoint reads `e.postData.contents` which works with any Content-Type.
