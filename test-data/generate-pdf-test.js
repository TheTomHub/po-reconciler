/**
 * Generate a test PO PDF and verify the positional parser can extract it.
 * Run: node test-data/generate-pdf-test.js
 */
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");

const outPath = path.join(__dirname, "tesco-po-2025-0247.pdf");

// Same data as the CSV test file (subset)
const rows = [
  ["1001V001", "Premium Organic Granola 500g",    "12", "4.29", "EA"],
  ["1002V001", "Free Range Eggs 12pk",             "40", "3.15", "DZ"],
  ["1003V001", "Sourdough Bread Loaf 800g",        "25", "2.89", "EA"],
  ["1005V001", "Greek Yoghurt 500g",               "30", "1.99", "EA"],
  ["1008V001", "Smoked Salmon 200g",                "8", "5.49", "EA"],
  ["1010V001", "Avocados Ripe & Ready 4pk",        "35", "2.75", "PK"],
  ["1015V001", "Italian Extra Virgin Olive Oil 1L", "6", "8.25", "EA"],
  ["1020V001", "Cornish Clotted Cream 227g",       "18", "2.99", "EA"],
  ["1025V001", "Wild Caught Tuna Steaks 2pk",       "5", "6.50", "PK"],
  ["1030V001", "British Chicken Breast 500g",      "22", "4.75", "EA"],
  ["1041V001", "Mature Cheddar 400g",              "28", "3.20", "EA"],
  ["1046",     "Vitamin C 1000mg 90 tabs",          "4", "7.99", "EA"],
  ["1047",     "Omega-3 Fish Oil 60 caps",          "6", "9.49", "EA"],
  ["1050",     "Zinc 15mg 120 tabs",                "3", "5.29", "EA"],
];

// Column X positions (points from left margin)
const colX = [72, 160, 380, 420, 480];
const headers = ["Item Code", "Description", "Qty", "Unit Price", "UOM"];

const doc = new PDFDocument({ size: "A4", margin: 50 });
const stream = fs.createWriteStream(outPath);
doc.pipe(stream);

// Title
doc.fontSize(16).font("Helvetica-Bold").text("PURCHASE ORDER", 72, 50);
doc.fontSize(10).font("Helvetica")
  .text("PO Number: PO-2025-0247", 72, 75)
  .text("Date: 15/01/2025", 72, 90)
  .text("Customer: Tesco Stores Ltd", 72, 105)
  .text("Delivery: 22/01/2025", 72, 120);

// Table header
const tableTop = 155;
doc.font("Helvetica-Bold").fontSize(9);
headers.forEach((h, i) => {
  doc.text(h, colX[i], tableTop, { width: 100 });
});

// Horizontal line
doc.moveTo(72, tableTop + 15).lineTo(540, tableTop + 15).stroke();

// Data rows
doc.font("Helvetica").fontSize(9);
rows.forEach((row, r) => {
  const y = tableTop + 22 + r * 16;
  row.forEach((cell, c) => {
    doc.text(cell, colX[c], y, { width: 200 });
  });
});

// Footer
const footerY = tableTop + 22 + rows.length * 16 + 20;
doc.font("Helvetica-Bold").fontSize(9)
  .text("Total Lines: " + rows.length, 72, footerY)
  .text("Currency: GBP", 72, footerY + 15);

doc.end();

stream.on("finish", () => {
  console.log("PDF written to", outPath);
  console.log("Size:", fs.statSync(outPath).size, "bytes");

  // Now test parsing with pdfjs-dist
  testParse(outPath);
});

async function testParse(pdfPath) {
  // Use the same pdfjs-dist that the add-in uses
  const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.mjs");

  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdf = await pdfjsLib.getDocument({ data }).promise;

  console.log("\n--- PDF Parse Test ---");
  console.log("Pages:", pdf.numPages);

  const allItems = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();

    for (const item of textContent.items) {
      const text = item.str;
      if (!text || !text.trim()) continue;
      allItems.push({
        x: item.transform[4],
        y: viewport.height - item.transform[5],
        text: text.trim(),
        width: item.width || 0,
      });
    }
  }

  console.log("Total text items:", allItems.length);

  // Group into rows
  const Y_TOL = 3;
  const sorted = [...allItems].sort((a, b) => a.y - b.y || a.x - b.x);
  const posRows = [];
  for (const item of sorted) {
    const last = posRows[posRows.length - 1];
    if (last && Math.abs(last.y - item.y) <= Y_TOL) {
      last.items.push(item);
    } else {
      posRows.push({ y: item.y, items: [item] });
    }
  }
  console.log("Rows detected:", posRows.length);

  // Merge items into cells per row
  const ITEM_GAP = 8;
  const rowCells = posRows.map((row) => {
    const s = [...row.items].sort((a, b) => a.x - b.x);
    if (s.length === 0) return [];
    const cells = [];
    let text = s[0].text, x = s[0].x, right = s[0].x + s[0].width;
    for (let i = 1; i < s.length; i++) {
      if (s[i].x - right > ITEM_GAP) {
        cells.push({ text, x });
        text = s[i].text;
        x = s[i].x;
      } else {
        text += " " + s[i].text;
      }
      right = Math.max(right, s[i].x + s[i].width);
    }
    cells.push({ text, x });
    return cells;
  });

  // Cluster columns
  const COL_TOL = 20;
  const columns = [];
  for (const cells of rowCells) {
    for (const cell of cells) {
      const match = columns.find((c) => Math.abs(c.center - cell.x) <= COL_TOL);
      if (match) {
        match.count++;
        match.center += (cell.x - match.center) / match.count;
      } else {
        columns.push({ center: cell.x, count: 1 });
      }
    }
  }
  columns.sort((a, b) => a.center - b.center);
  console.log("Columns detected:", columns.length);

  // Build boundaries
  const bounds = [-Infinity];
  for (let i = 0; i < columns.length - 1; i++) {
    bounds.push((columns[i].center + columns[i + 1].center) / 2);
  }
  bounds.push(Infinity);

  // Assign to columns
  const rawRows = rowCells.map((cells) => {
    const row = new Array(columns.length).fill("");
    for (const cell of cells) {
      for (let c = 0; c < columns.length; c++) {
        if (cell.x >= bounds[c] && cell.x < bounds[c + 1]) {
          row[c] = row[c] ? row[c] + " " + cell.text : cell.text;
          break;
        }
      }
    }
    return row;
  });

  console.log("\n--- Extracted Table ---");
  let passed = 0;
  let failed = 0;

  function assert(cond, msg) {
    if (cond) { passed++; console.log("  PASS:", msg); }
    else { failed++; console.log("  FAIL:", msg); }
  }

  // Find the header row (should contain "Item Code", "Description", etc.)
  const headerIdx = rawRows.findIndex((r) =>
    r.some((c) => c.toLowerCase().includes("item")) &&
    r.some((c) => c.toLowerCase().includes("price"))
  );
  assert(headerIdx >= 0, "Found header row at index " + headerIdx);

  if (headerIdx >= 0) {
    const hdr = rawRows[headerIdx];
    console.log("  Header:", JSON.stringify(hdr));

    assert(hdr.some((c) => c.includes("Item Code")), "Header has 'Item Code'");
    assert(hdr.some((c) => c.includes("Description")), "Header has 'Description'");
    assert(hdr.some((c) => c.includes("Qty")), "Header has 'Qty'");
    assert(hdr.some((c) => c.includes("Unit Price")), "Header has 'Unit Price'");
    assert(hdr.some((c) => c.includes("UOM")), "Header has 'UOM'");

    // Check data rows
    const dataRows = rawRows.slice(headerIdx + 1).filter((r) =>
      r.some((c) => c.trim() !== "")
    );

    // Filter to only rows that look like product data (have a SKU-like value)
    const productRows = dataRows.filter((r) =>
      r.some((c) => /^\d{4}/.test(c.trim()))
    );

    assert(productRows.length === rows.length,
      `Found ${productRows.length} product rows (expected ${rows.length})`);

    // Check first and last data rows
    if (productRows.length > 0) {
      const first = productRows[0];
      assert(first.some((c) => c.includes("1001V001")), "First row has SKU 1001V001");
      assert(first.some((c) => c.includes("4.29")), "First row has price 4.29");

      const last = productRows[productRows.length - 1];
      assert(last.some((c) => c.includes("1050")), "Last row has SKU 1050");
      assert(last.some((c) => c.includes("5.29")), "Last row has price 5.29");
    }
  }

  console.log(`\n${passed} passed, ${failed} failed`);
  process.exit(failed > 0 ? 1 : 0);
}
