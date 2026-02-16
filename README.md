# PO Reconciler

Excel Add-in that automates purchase order reconciliation against ERP exports. Compares customer PO prices with your system prices and flags exceptions — replacing manual VLOOKUP workflows.

## What it does

1. **Upload a customer PO** (Excel, CSV, or PDF) — auto-detects the header row even in messy order forms with instructions/legends at the top
2. **Select your ERP data** range in Excel
3. **Click Reconcile** — creates a formatted results sheet with:
   - Color-coded rows (green = match, yellow = within tolerance, red = exception)
   - Summary section with total exposure
   - Pre-filled email draft for your team

## Key features

- **Smart header detection** — scans all sheets and rows to find the actual data table, skipping cover pages, instructions, and status legends
- **Prefix SKU matching** — matches customer core numbers (e.g. `1234`) to full ERP variant SKUs (e.g. `1234V012`)
- **Column auto-detection** — recognizes common column names (SKU, Item #, Ordered Item, Price, Unit Cost, etc.)
- **Manual fallback** — if auto-detection fails, dropdown selectors let the user pick columns
- **Configurable tolerance** — set a $ threshold for acceptable price differences

## Installation

### For development

```bash
npm install
npm start        # https://localhost:3000
```

### Sideload in Excel

1. Open Excel (desktop or Excel Online)
2. Insert → My Add-ins → Upload My Add-in
3. Upload `manifest.xml`
4. The "PO Reconciler" tab appears in the ribbon

### Deploy via GitHub Pages

The add-in is served from `https://thetomhub.github.io/po-reconciler/`. The manifest.xml points to this URL for production use.

## Test data

Generate sample files that mimic real-world order forms:

```bash
node test-data/generate.js        # Simple ERP export + customer PO
node test-data/generate-oxo.js    # Multi-sheet order form with messy layout
```

## Tech stack

- **Office.js** — Excel Add-in APIs
- **Webpack** — bundling with content-hash filenames
- **Papa Parse** — CSV parsing
- **SheetJS** — Excel file reading
- **PDF.js** — PDF text extraction
