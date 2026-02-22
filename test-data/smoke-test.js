#!/usr/bin/env node
/**
 * Chandlr — End-to-end smoke test
 *
 * Runs the full pipeline against the Tesco PO + ERP test files:
 *   CSV parse → Column detection → Extraction → Validation →
 *   Reconciliation → Credit note → Re-invoice → Email draft
 *
 * All core logic is inlined (same algorithms as src/) to avoid ESM/CJS issues.
 * This tests data flow correctness, not Office.js rendering.
 */

const fs = require("fs");
const path = require("path");

// ── Helpers (from src/utils/format.js) ──

function parseNumber(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") return isNaN(value) ? null : value;
  const cleaned = String(value).replace(/[$£€,\s]/g, "").trim();
  if (cleaned === "") return null;
  const parenMatch = cleaned.match(/^\((.+)\)$/);
  if (parenMatch) {
    const num = parseFloat(parenMatch[1]);
    return isNaN(num) ? null : -num;
  }
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

let activeCurrency = "GBP";
const CURRENCY_CONFIG = {
  GBP: { locale: "en-GB", currency: "GBP" },
};

function formatCurrency(value) {
  if (value == null) return "—";
  const num = typeof value === "number" ? value : parseFloat(value);
  if (isNaN(num)) return "—";
  const cfg = CURRENCY_CONFIG[activeCurrency];
  return new Intl.NumberFormat(cfg.locale, {
    style: "currency",
    currency: cfg.currency,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(num);
}

function round(n) {
  return Math.round(n * 100) / 100;
}

// ── CSV Parser ──

function parseCSV(text) {
  const lines = text.split("\n").filter((l) => l.trim());
  if (lines.length === 0) return { headers: [], rows: [] };

  // Find the header row (first row with 3+ comma-separated columns)
  let headerIdx = -1;
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const cols = lines[i].split(",");
    if (cols.length >= 3) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) throw new Error("No header row found");

  // Extract metadata from rows above header
  const metadata = {};
  for (let i = 0; i < headerIdx; i++) {
    const parts = lines[i].split(",");
    if (parts.length >= 2) {
      metadata[parts[0].trim()] = parts[1].trim();
    }
  }

  const headers = lines[headerIdx].split(",").map((h) => h.trim());
  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const vals = lines[i].split(",").map((v) => v.trim());
    if (vals.length < 2) continue;
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = vals[j] || "";
    }
    rows.push(row);
  }
  return { headers, rows, metadata };
}

// ── Column Detector (from src/reconcile/detector.js + src/capture/extractor.js) ──

const SKU_ALIASES = ["sku", "item number", "product code", "part number", "item no", "item #", "material", "product id", "article"];
const PRICE_ALIASES = ["price", "unit price", "cost", "amount", "unit cost", "net price", "list price"];
const NAME_ALIASES = ["product name", "description", "item description", "product description", "name"];
const QTY_ALIASES = ["qty", "quantity", "order qty", "ordered qty", "units", "order quantity"];
const UOM_ALIASES = ["uom", "unit of measure", "unit", "pack size"];
const LINE_TOTAL_ALIASES = ["line total", "total", "ext amount", "line amount", "net amount", "value"];

function findColumn(headers, aliases) {
  const normalized = headers.map((h) => h.toLowerCase().trim());
  for (const alias of aliases) {
    const idx = normalized.indexOf(alias);
    if (idx !== -1) return headers[idx];
  }
  for (const alias of aliases) {
    const idx = normalized.findIndex((h) => h.includes(alias));
    if (idx !== -1) return headers[idx];
  }
  for (const alias of aliases) {
    const idx = normalized.findIndex((h) => alias.includes(h) && h.length >= 3);
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function detectColumns(headers) {
  return {
    sku: findColumn(headers, SKU_ALIASES),
    price: findColumn(headers, PRICE_ALIASES),
    name: findColumn(headers, NAME_ALIASES),
    qty: findColumn(headers, QTY_ALIASES),
  };
}

function detectAllColumns(headers) {
  const base = detectColumns(headers);
  return {
    ...base,
    uom: findColumn(headers, UOM_ALIASES),
    lineTotal: findColumn(headers, LINE_TOTAL_ALIASES),
  };
}

// ── Extractor (from src/capture/extractor.js) ──

function extractPOData(parsedData, detectedColumns) {
  const cols = detectedColumns || detectAllColumns(parsedData.headers);
  if (!cols.sku) throw new Error("Cannot find SKU column. Headers: " + parsedData.headers.join(", "));

  const stagingRows = [];
  const warnings = [];

  for (let i = 0; i < parsedData.rows.length; i++) {
    const raw = parsedData.rows[i];
    const lineNum = i + 1;
    const sku = (raw[cols.sku] || "").trim();
    if (!sku) { warnings.push({ line: lineNum, field: "SKU", message: "Empty SKU" }); continue; }

    const price = cols.price ? parseNumber(raw[cols.price]) : null;
    const qty = cols.qty ? parseNumber(raw[cols.qty]) : null;
    const name = cols.name ? (raw[cols.name] || "").trim() : "";
    const uom = cols.uom ? (raw[cols.uom] || "").trim() : "";

    const rawLineTotal = cols.lineTotal ? raw[cols.lineTotal] : null;
    let lineTotal = rawLineTotal != null ? parseNumber(rawLineTotal) : null;
    if (lineTotal === null && price !== null && qty !== null) {
      lineTotal = round(price * qty);
    }

    // Cross-check line total
    if (lineTotal !== null && price !== null && qty !== null) {
      const expected = round(price * qty);
      if (Math.abs(lineTotal - expected) > 0.02) {
        warnings.push({ line: lineNum, field: "Line Total", message: `Mismatch: ${lineTotal} vs expected ${expected}` });
      }
    }

    if (qty !== null && qty <= 0) {
      warnings.push({ line: lineNum, field: "Qty", message: `Zero or negative: ${qty}` });
    }

    stagingRows.push({ lineNum, sku, name, qty: qty ?? 1, price: price ?? 0, uom, lineTotal: lineTotal ?? 0, status: "Pending" });
  }

  // Quantity outlier check
  const quantities = stagingRows.map((r) => r.qty).filter((q) => q > 0).sort((a, b) => a - b);
  if (quantities.length >= 5) {
    const median = quantities[Math.floor(quantities.length / 2)];
    const threshold = median * 3;
    for (const row of stagingRows) {
      if (row.qty > threshold) {
        warnings.push({ line: row.lineNum, field: "Qty", message: `Outlier: ${row.qty} (median: ${median}, threshold: ${threshold})` });
      }
    }
  }

  const totalValue = round(stagingRows.reduce((sum, r) => sum + r.lineTotal, 0));
  return {
    stagingRows,
    metadata: { lineCount: stagingRows.length, totalValue, detectedFields: Object.entries(cols).filter(([, v]) => v).map(([k]) => k), warningCount: warnings.length },
    warnings,
  };
}

// ── Reconciliation Engine (from src/reconcile/reconcile.js) ──

function normalizeSku(sku) { return String(sku).trim().toUpperCase(); }

function findPrefixMatches(normPoSku, erpMap) {
  const results = [];
  for (const [erpSku, entry] of erpMap) {
    if (erpSku.startsWith(normPoSku) && erpSku !== normPoSku) {
      results.push({ sku: erpSku, entry });
    }
  }
  return results;
}

function reconcile({ poData, poColumns, erpData, erpColumns, tolerance }) {
  const erpMap = new Map();
  const erpDuplicates = new Set();
  for (const row of erpData.rows) {
    const rawSku = row[erpColumns.sku];
    if (!rawSku) continue;
    const normSku = normalizeSku(rawSku);
    if (erpMap.has(normSku)) erpDuplicates.add(normSku);
    erpMap.set(normSku, {
      price: parseNumber(row[erpColumns.price]),
      name: erpColumns.name ? row[erpColumns.name] || "" : "",
      qty: erpColumns.qty ? parseNumber(row[erpColumns.qty]) || 1 : 1,
      originalRow: row,
      matched: false,
    });
  }

  const poSkuCounts = new Map();
  for (const row of poData.rows) {
    const rawSku = row[poColumns.sku];
    if (!rawSku) continue;
    poSkuCounts.set(normalizeSku(rawSku), (poSkuCounts.get(normalizeSku(rawSku)) || 0) + 1);
  }
  const poDuplicates = new Set();
  for (const [sku, count] of poSkuCounts) { if (count > 1) poDuplicates.add(sku); }

  const resultRows = [];
  let matches = 0, tolerances = 0, exceptions = 0, exposure = 0, warnings = 0;

  for (const row of poData.rows) {
    const rawSku = row[poColumns.sku];
    if (!rawSku) continue;
    const normSku = normalizeSku(rawSku);
    const poPrice = parseNumber(row[poColumns.price]);
    const poName = poColumns.name ? row[poColumns.name] || "" : "";
    const poQty = poColumns.qty ? parseNumber(row[poColumns.qty]) || 1 : 1;
    const isDuplicate = poDuplicates.has(normSku) || erpDuplicates.has(normSku);

    if (poPrice === null) {
      warnings++;
      resultRows.push({ status: "Warning", sku: rawSku, name: poName, erpPrice: null, poPrice: null, diff: null, pctDiff: null, action: "Non-numeric price", duplicate: isDuplicate, poQty, erpQty: null, lineTotal: null });
      continue;
    }

    let oracle = erpMap.get(normSku);
    let matchType = "exact";
    if (!oracle) {
      const prefixMatches = findPrefixMatches(normSku, erpMap);
      if (prefixMatches.length === 1) { oracle = prefixMatches[0].entry; matchType = "prefix"; }
      else if (prefixMatches.length > 1) {
        warnings++;
        resultRows.push({ status: "Warning", sku: rawSku, name: poName, erpPrice: null, poPrice, diff: null, pctDiff: null, action: `Multiple ERP matches`, duplicate: isDuplicate, poQty, erpQty: null, lineTotal: round(poPrice * poQty) });
        continue;
      }
    }

    if (!oracle) {
      exceptions++;
      resultRows.push({ status: "Not in ERP", sku: rawSku, name: poName, erpPrice: null, poPrice, diff: null, pctDiff: null, action: "SKU not found in ERP", duplicate: isDuplicate, poQty, erpQty: null, lineTotal: round(poPrice * poQty) });
      continue;
    }

    oracle.matched = true;
    const diff = round(poPrice - oracle.price);
    const absDiff = Math.abs(diff);
    const pctDiff = oracle.price !== 0 ? round((diff / oracle.price) * 100) : (diff !== 0 ? 100 : 0);

    let status, action;
    if (absDiff === 0) { status = "Match"; action = "OK"; matches++; }
    else if (absDiff <= tolerance) { status = "Tolerance"; action = "Within tolerance"; tolerances++; }
    else { status = "Exception"; action = "Review pricing"; exceptions++; exposure += absDiff; }

    resultRows.push({
      status, sku: rawSku, erpSku: matchType === "prefix" ? findPrefixMatches(normSku, erpMap)[0]?.sku : null,
      name: oracle.name || poName, erpPrice: oracle.price, poPrice, diff, pctDiff, action,
      duplicate: isDuplicate, poQty, erpQty: oracle.qty, lineTotal: round(poPrice * poQty),
    });
  }

  // Unmatched ERP rows
  for (const [normSku, oracle] of erpMap) {
    if (oracle.matched) continue;
    resultRows.push({
      status: "Not in PO", sku: normSku, name: oracle.name, erpPrice: oracle.price, poPrice: null,
      diff: null, pctDiff: null, action: "SKU not in PO", duplicate: erpDuplicates.has(normSku),
      poQty: null, erpQty: oracle.qty, lineTotal: null,
    });
  }

  const statusOrder = { Exception: 0, "Not in ERP": 1, "Not in PO": 2, Warning: 3, Tolerance: 4, Match: 5 };
  resultRows.sort((a, b) => {
    const sa = statusOrder[a.status] ?? 99, sb = statusOrder[b.status] ?? 99;
    if (sa !== sb) return sa - sb;
    return (b.diff != null ? Math.abs(b.diff) : 0) - (a.diff != null ? Math.abs(a.diff) : 0);
  });

  return {
    summary: { total: resultRows.length, matches, tolerances, exceptions, exposure: round(exposure), warnings, timestamp: new Date().toISOString() },
    rows: resultRows,
  };
}

// ── Credit Note / Re-Invoice (from src/reconcile/creditnote.js) ──

function generateCreditNote(results) {
  const creditRows = [];
  let totalCredit = 0;
  for (const row of results.rows) {
    if (row.status !== "Exception" && row.status !== "Tolerance") continue;
    if (row.poPrice == null) continue;
    const qty = row.poQty || 1;
    const lineTotal = round(row.poPrice * qty);
    const creditAmount = round(-lineTotal);
    creditRows.push({ sku: row.sku, name: row.name || "", qty, originalPrice: row.poPrice, erpPrice: row.erpPrice, diff: row.diff, lineTotal, creditAmount });
    totalCredit = round(totalCredit + creditAmount);
  }
  return { creditRows, totals: { lineCount: creditRows.length, totalCredit } };
}

function generateCorrectedInvoice(results) {
  const invoiceRows = [];
  let totalInvoice = 0;
  for (const row of results.rows) {
    if (row.status !== "Exception" && row.status !== "Tolerance") continue;
    if (row.erpPrice == null) continue;
    const qty = row.poQty || 1;
    const correctedPrice = row.erpPrice;
    const lineTotal = round(correctedPrice * qty);
    invoiceRows.push({ sku: row.sku, name: row.name || "", qty, originalPrice: row.poPrice, correctedPrice, lineTotal, diff: row.diff });
    totalInvoice = round(totalInvoice + lineTotal);
  }
  return { invoiceRows, totals: { lineCount: invoiceRows.length, totalInvoice } };
}

// ── Email Draft (from src/email/email.js) ──

function generateEmailDraft(results, filename) {
  const stem = filename.replace(/\.[^.]+$/, "");
  const poMatch = stem.match(/(?:PO[-_ ]?)(\d+[-\w]*)/i);
  const poNumber = poMatch ? `PO-${poMatch[1]}` : stem;
  const { summary, rows } = results;
  const exceptionRows = rows.filter((r) => r.status === "Exception" || r.status === "Not in ERP" || r.status === "Not in PO");
  const topExceptions = exceptionRows.slice(0, 3);
  const subject = `PO ${poNumber} — Reconciliation: ${summary.exceptions} exception${summary.exceptions !== 1 ? "s" : ""} found`;
  const topLines = topExceptions.map((r) => {
    if (r.status === "Not in ERP") return `  - SKU ${r.sku}: Not found in ERP`;
    if (r.status === "Not in PO") return `  - SKU ${r.sku}: In ERP but not on PO`;
    return `  - SKU ${r.sku}: PO ${formatCurrency(r.poPrice)} vs ERP ${formatCurrency(r.erpPrice)} (diff: ${formatCurrency(r.diff)})`;
  });
  const body = `PO reconciliation for ${poNumber}: ${summary.total} items, ${summary.matches} matches, ${summary.tolerances} tolerance, ${summary.exceptions} exceptions. Exposure: ${formatCurrency(summary.exposure)}.\n\nTop exceptions:\n${topLines.join("\n")}`;
  return { subject, body };
}

// ═══════════════════════════════════════════════════════
// SMOKE TEST
// ═══════════════════════════════════════════════════════

let passed = 0;
let failed = 0;

function assert(condition, label) {
  if (condition) {
    passed++;
    console.log(`  ✓ ${label}`);
  } else {
    failed++;
    console.log(`  ✗ ${label}`);
  }
}

function assertEq(actual, expected, label) {
  if (actual === expected) {
    passed++;
    console.log(`  ✓ ${label}`);
  } else {
    failed++;
    console.log(`  ✗ ${label}: expected ${expected}, got ${actual}`);
  }
}

console.log("═══════════════════════════════════════════════════════");
console.log("  CHANDLR — End-to-end Smoke Test");
console.log("═══════════════════════════════════════════════════════\n");

// ── Step 1: Parse CSV files ──

console.log("Step 1: Parse CSV files");

const poCSV = fs.readFileSync(path.join(__dirname, "tesco-po-2025-0247.csv"), "utf-8");
const erpCSV = fs.readFileSync(path.join(__dirname, "erp-price-list-2025-02.csv"), "utf-8");

const poData = parseCSV(poCSV);
const erpData = parseCSV(erpCSV);

assertEq(poData.rows.length, 50, `PO has 50 data rows`);
assert(erpData.rows.length >= 48, `ERP has ${erpData.rows.length} data rows (≥48)`);
assert(poData.metadata["Purchase Order"] === "PO-2025-0247", `PO metadata: ref = PO-2025-0247`);
assert(poData.metadata["Customer"] === "Tesco Stores Ltd", `PO metadata: customer = Tesco Stores Ltd`);

// ── Step 2: Column detection ──

console.log("\nStep 2: Column detection");

const poCols = detectAllColumns(poData.headers);
assertEq(poCols.sku, "Item Number", `PO SKU column: "${poCols.sku}"`);
assertEq(poCols.price, "Unit Price", `PO Price column: "${poCols.price}"`);
assertEq(poCols.name, "Description", `PO Name column: "${poCols.name}"`);
assertEq(poCols.qty, "Order Qty", `PO Qty column: "${poCols.qty}"`);
assertEq(poCols.uom, "UOM", `PO UOM column: "${poCols.uom}"`);
assertEq(poCols.lineTotal, "Line Total", `PO Line Total column: "${poCols.lineTotal}"`);

const erpCols = detectColumns(erpData.headers);
assertEq(erpCols.sku, "SKU", `ERP SKU column: "${erpCols.sku}"`);
assertEq(erpCols.price, "List Price", `ERP Price column: "${erpCols.price}"`);
assertEq(erpCols.name, "Product Description", `ERP Name column: "${erpCols.name}"`);

// ── Step 3: Extract PO data ──

console.log("\nStep 3: Extract PO data (Capture module)");

const extraction = extractPOData(poData, poCols);
assertEq(extraction.stagingRows.length, 50, `Extracted 50 staging rows`);
assert(extraction.metadata.totalValue > 5000, `Total value: £${extraction.metadata.totalValue.toFixed(2)} (>£5000)`);
assert(extraction.metadata.detectedFields.includes("sku"), `Detected field: sku`);
assert(extraction.metadata.detectedFields.includes("price"), `Detected field: price`);
assert(extraction.metadata.detectedFields.includes("lineTotal"), `Detected field: lineTotal`);

// Verify line total cross-checks pass (no mismatch warnings)
const lineTotalWarnings = extraction.warnings.filter((w) => w.field === "Line Total");
assertEq(lineTotalWarnings.length, 0, `No line total mismatch warnings`);

// Verify quantity outlier detection fires
const qtyOutlierWarnings = extraction.warnings.filter((w) => w.message.includes("Outlier"));
assert(qtyOutlierWarnings.length >= 1, `Quantity outlier warnings: ${qtyOutlierWarnings.length} (Semi-Skimmed Milk 200, Chopped Tomatoes 150)`);

// Spot-check a row
const tea = extraction.stagingRows.find((r) => r.sku === "1001V001");
assertEq(tea.name, "Organic Green Tea 250g", `Row 1: name correct`);
assertEq(tea.qty, 48, `Row 1: qty = 48`);
assertEq(tea.price, 2.49, `Row 1: price = 2.49`);
assertEq(tea.lineTotal, 119.52, `Row 1: lineTotal = 119.52`);
assertEq(tea.uom, "EA", `Row 1: UOM = EA`);

// ── Step 4: Validate ──

console.log("\nStep 4: Validate extracted data");

// Inline validator (from src/validate/validator.js)
function validate(rows) {
  const all = [];

  // No data
  if (rows.length === 0) {
    all.push({ severity: "error", rule: "no-data", message: "No data rows found." });
  }

  // Duplicate SKUs
  const skuMap = new Map();
  for (const row of rows) {
    const norm = row.sku.toUpperCase().trim();
    if (!skuMap.has(norm)) skuMap.set(norm, []);
    skuMap.get(norm).push(row.lineNum);
  }
  for (const [sku, lines] of skuMap) {
    if (lines.length > 1) {
      all.push({ severity: "warning", rule: "duplicate-sku", field: "SKU", lines, message: `Duplicate SKU "${sku}" on lines ${lines.join(", ")}` });
    }
  }

  // Missing prices
  const missingPrices = rows.filter((r) => r.price === 0 || r.price === null);
  if (missingPrices.length === rows.length) {
    all.push({ severity: "error", rule: "no-prices", message: "No rows have a valid price." });
  } else {
    for (const r of missingPrices) {
      all.push({ severity: "error", rule: "missing-price", field: "Price", lines: [r.lineNum], message: `Line ${r.lineNum} (${r.sku}): missing or zero price` });
    }
  }

  // Zero/negative quantities
  for (const row of rows) {
    if (row.qty <= 0) {
      all.push({ severity: "error", rule: "invalid-qty", field: "Qty", lines: [row.lineNum], message: `Line ${row.lineNum} (${row.sku}): quantity is ${row.qty}` });
    }
  }

  // Price range
  for (const row of rows) {
    if (row.price > 0 && (row.price < 0.01 || row.price > 9999.99)) {
      all.push({ severity: "warning", rule: "price-range", field: "Price", lines: [row.lineNum], message: `Line ${row.lineNum} (${row.sku}): price ${row.price} outside expected range` });
    }
  }

  // Quantity outliers
  const quantities = rows.map((r) => r.qty).filter((q) => q > 0).sort((a, b) => a - b);
  if (quantities.length >= 5) {
    const median = quantities[Math.floor(quantities.length / 2)];
    const threshold = median * 5;
    for (const row of rows) {
      if (row.qty > threshold) {
        all.push({ severity: "warning", rule: "qty-outlier", field: "Qty", lines: [row.lineNum], message: `Line ${row.lineNum} (${row.sku}): quantity ${row.qty} is unusually large (median: ${median})` });
      }
    }
  }

  // Line total consistency
  for (const row of rows) {
    if (row.lineTotal > 0 && row.price > 0 && row.qty > 0) {
      const expected = Math.round(row.price * row.qty * 100) / 100;
      if (Math.abs(row.lineTotal - expected) > 0.02) {
        all.push({ severity: "warning", rule: "line-total-mismatch", field: "Line Total", lines: [row.lineNum], message: `Line ${row.lineNum} (${row.sku}): total mismatch` });
      }
    }
  }

  // Missing names
  const missingNames = rows.filter((r) => !r.name || r.name.trim() === "");
  const hasNames = rows.some((r) => r.name && r.name.trim());
  if (missingNames.length > 0 && !hasNames) {
    all.push({ severity: "info", rule: "no-names", message: "No product names detected." });
  } else if (missingNames.length > 0) {
    all.push({ severity: "info", rule: "missing-names", lines: missingNames.map((r) => r.lineNum), message: `${missingNames.length} row(s) missing product name` });
  }

  // SKU format consistency
  const skus = rows.map((r) => r.sku);
  const hasLetters = skus.filter((s) => /[a-zA-Z]/.test(s));
  const pureNumeric = skus.filter((s) => /^\d+$/.test(s));
  if (hasLetters.length > 0 && pureNumeric.length > 0 && pureNumeric.length <= rows.length * 0.3) {
    all.push({ severity: "info", rule: "mixed-sku-format", message: `Mixed SKU formats: ${hasLetters.length} alphanumeric, ${pureNumeric.length} numeric-only.` });
  }

  const errors = all.filter((r) => r.severity === "error");
  const warnings = all.filter((r) => r.severity === "warning");
  const info = all.filter((r) => r.severity === "info");

  return { valid: errors.length === 0, errors, warnings, info, summary: { total: all.length, errors: errors.length, warnings: warnings.length, info: info.length } };
}

// 4a: Validate clean Tesco data
const validation = validate(extraction.stagingRows);
assert(validation.valid === true, `Tesco PO passes validation (no errors)`);
assertEq(validation.summary.errors, 0, `Zero validation errors`);
assert(validation.summary.warnings >= 0, `Warnings: ${validation.summary.warnings}`);

// Print validation summary
console.log(`  Summary: ${validation.summary.errors} errors, ${validation.summary.warnings} warnings, ${validation.summary.info} info`);
for (const w of validation.warnings) {
  console.log(`    ! ${w.message}`);
}
for (const i of validation.info) {
  console.log(`    ○ ${i.message}`);
}

// 4b: Validate edge cases — empty data
const emptyValidation = validate([]);
assert(emptyValidation.valid === false, `Empty data fails validation`);
assert(emptyValidation.errors.some((e) => e.rule === "no-data"), `Empty data triggers no-data error`);

// 4c: Validate edge cases — missing prices
const badPriceRows = [
  { lineNum: 1, sku: "TEST001", name: "Test", qty: 10, price: 0, uom: "EA", lineTotal: 0, status: "Pending" },
  { lineNum: 2, sku: "TEST002", name: "Test 2", qty: 5, price: 2.99, uom: "EA", lineTotal: 14.95, status: "Pending" },
];
const badPriceValidation = validate(badPriceRows);
assert(badPriceValidation.valid === false, `Missing price fails validation`);
assert(badPriceValidation.errors.some((e) => e.rule === "missing-price"), `Missing price triggers error`);

// 4d: Validate edge cases — duplicate SKUs
const dupeRows = [
  { lineNum: 1, sku: "SKU001", name: "Item A", qty: 10, price: 1.99, uom: "EA", lineTotal: 19.9, status: "Pending" },
  { lineNum: 2, sku: "SKU001", name: "Item A", qty: 5, price: 1.99, uom: "EA", lineTotal: 9.95, status: "Pending" },
  { lineNum: 3, sku: "SKU002", name: "Item B", qty: 8, price: 3.49, uom: "EA", lineTotal: 27.92, status: "Pending" },
];
const dupeValidation = validate(dupeRows);
assert(dupeValidation.warnings.some((w) => w.rule === "duplicate-sku"), `Duplicate SKU triggers warning`);

// 4e: Mixed SKU format detection on Tesco data (has both alphanumeric and numeric-only)
const mixedSkuInfo = validation.info.find((i) => i.rule === "mixed-sku-format");
assert(mixedSkuInfo !== undefined, `Mixed SKU format detected (45 alphanumeric + 5 numeric-only)`);

// ── Step 5: Reconcile ──

console.log("\nStep 5: Reconcile PO vs ERP");

const results = reconcile({
  poData,
  poColumns: { sku: "Item Number", price: "Unit Price", name: "Description", qty: "Order Qty" },
  erpData,
  erpColumns: { sku: "SKU", price: "List Price", name: "Product Description" },
  tolerance: 0.02,
});

console.log(`\n  Summary: ${results.summary.total} total, ${results.summary.matches} match, ${results.summary.tolerances} tolerance, ${results.summary.exceptions} exception, ${results.summary.warnings} warning`);
console.log(`  Exposure: £${results.summary.exposure.toFixed(2)}\n`);

// Expected scenarios:
// Exact matches: ~37 (most items + 4 prefix matches with same price)
// Tolerance (≤0.02): 3 → 1006 (+0.01), 1014 (+0.01), 1029 (+0.02)
// Exceptions (>0.02): 10 → 1002 (+0.20), 1008 (+0.20), 1009 (+0.30), 1016 (+0.16),
//   1022 (+0.16), 1025 (+0.16), 1037 (+0.30), 1041 (+0.20), 1046 (+0.10), 1047 (+0.50)
// Not in ERP: 1 → 1049 (Vitamin D)
// Not in PO: 3 → 1051, 1052, 1053

assert(results.summary.matches >= 35, `Matches ≥ 35 (got ${results.summary.matches})`);
assert(results.summary.tolerances >= 2, `Tolerances ≥ 2 (got ${results.summary.tolerances})`);
assert(results.summary.exceptions >= 8, `Exceptions ≥ 8 (got ${results.summary.exceptions})`);
assert(results.summary.exposure > 1, `Exposure > £1 (got £${results.summary.exposure.toFixed(2)})`);

// Check specific scenarios
const byStatus = {};
for (const row of results.rows) {
  if (!byStatus[row.status]) byStatus[row.status] = [];
  byStatus[row.status].push(row);
}

// Not in ERP — should be 1049 (Vitamin D)
const notInErp = byStatus["Not in ERP"] || [];
assert(notInErp.some((r) => r.sku === "1049"), `SKU 1049 (Vitamin D) flagged as Not in ERP`);

// Not in PO — should be 1051, 1052, 1053
const notInPo = byStatus["Not in PO"] || [];
assert(notInPo.length >= 3, `${notInPo.length} ERP-only SKUs detected (expected ≥3)`);

// Prefix match — 1046 (bare) should match 1046V001 in ERP
const row1046 = results.rows.find((r) => r.sku === "1046");
assert(row1046 !== undefined, `SKU 1046 (bare) found in results`);
if (row1046) {
  assert(row1046.erpSku === "1046V001" || row1046.status !== "Not in ERP", `SKU 1046 prefix-matched to ERP (status: ${row1046.status})`);
}

// Exception check — Chicken 1008V003: PO £4.29 vs ERP £4.49
const chicken = results.rows.find((r) => r.sku === "1008V003");
assert(chicken && chicken.status === "Exception", `Chicken (1008V003) is an Exception`);
if (chicken) {
  assertEq(chicken.poPrice, 4.29, `Chicken PO price = £4.29`);
  assertEq(chicken.erpPrice, 4.49, `Chicken ERP price = £4.49`);
  assertEq(chicken.diff, -0.20, `Chicken diff = -£0.20`);
}

// Tolerance check — Butter 1006V001: PO £1.79 vs ERP £1.80
const butter = results.rows.find((r) => r.sku === "1006V001");
assert(butter && butter.status === "Tolerance", `Butter (1006V001) is within Tolerance`);
if (butter) {
  assertEq(butter.diff, -0.01, `Butter diff = -£0.01`);
}

// Print exception details
console.log("  Exception details:");
for (const row of (byStatus["Exception"] || [])) {
  console.log(`    ${row.sku} ${(row.name || "").padEnd(30)} PO: £${row.poPrice?.toFixed(2)}  ERP: £${row.erpPrice?.toFixed(2)}  diff: £${row.diff?.toFixed(2)}`);
}

// ── Step 6: Staged Entry ──

console.log("\nStep 6: Generate ERP staging data");

// Inline staging entry generator (from src/entry/entry.js)
function generateStagingEntry(results, options = {}) {
  const entryRows = [];
  let totalValue = 0, readyCount = 0, reviewCount = 0, holdCount = 0, lineNum = 0;

  for (const row of results.rows) {
    if (row.status === "Not in PO") continue;
    if (row.poPrice == null && row.erpPrice == null) continue;
    lineNum++;

    let entryPrice, status, notes;
    switch (row.status) {
      case "Match":
        entryPrice = row.erpPrice ?? row.poPrice; status = "Ready"; notes = ""; break;
      case "Tolerance":
        entryPrice = row.erpPrice ?? row.poPrice; status = "Ready"; notes = `Within tolerance (diff: ${row.diff})`; break;
      case "Exception":
        entryPrice = row.erpPrice ?? row.poPrice; status = "Review"; notes = `Price exception: PO ${row.poPrice} vs ERP ${row.erpPrice}`; break;
      case "Not in ERP":
        entryPrice = row.poPrice; status = "Hold"; notes = "SKU not found in ERP"; break;
      case "Warning":
        entryPrice = row.poPrice ?? 0; status = "Hold"; notes = row.action || "Data quality issue"; break;
      default:
        entryPrice = row.erpPrice ?? row.poPrice ?? 0; status = "Review"; notes = "";
    }

    const qty = row.poQty || 1;
    const lineTotal = round(entryPrice * qty);
    totalValue = round(totalValue + lineTotal);

    if (status === "Ready") readyCount++;
    else if (status === "Review") reviewCount++;
    else holdCount++;

    entryRows.push({ lineNum, sku: row.sku, erpSku: row.erpSku || "", name: row.name || "", qty, uom: "EA", entryPrice, lineTotal, status, notes });
  }

  return {
    entryRows,
    totals: { lineCount: entryRows.length, totalValue, readyCount, reviewCount, holdCount },
    metadata: { poRef: options.poRef || "Unknown", customer: options.customer || "Unknown", generatedAt: new Date().toISOString() },
  };
}

const entryData = generateStagingEntry(results, { poRef: "PO-2025-0247", customer: "Tesco Stores Ltd" });

assertEq(entryData.entryRows.length, 50, `ERP staging has 50 lines (excludes ERP-only)`);
assert(entryData.totals.totalValue > 5000, `Total value: £${entryData.totals.totalValue.toFixed(2)} (>£5000)`);
assert(entryData.totals.readyCount >= 35, `Ready (green): ${entryData.totals.readyCount} (≥35)`);
assert(entryData.totals.reviewCount >= 8, `Review (yellow): ${entryData.totals.reviewCount} (≥8)`);
assert(entryData.totals.holdCount >= 1, `Hold (red): ${entryData.totals.holdCount} (≥1)`);

// Verify ready + review + hold = total
assertEq(entryData.totals.readyCount + entryData.totals.reviewCount + entryData.totals.holdCount, entryData.totals.lineCount, `Status counts sum to total lines`);

// Verify "Not in PO" rows are excluded
const erpOnlyInStaging = entryData.entryRows.filter((r) => r.sku === "1051V001" || r.sku === "1052V002" || r.sku === "1053V003");
assertEq(erpOnlyInStaging.length, 0, `ERP-only SKUs excluded from staging`);

// Verify "Not in ERP" row is on hold
const vitaminD = entryData.entryRows.find((r) => r.sku === "1049");
assert(vitaminD && vitaminD.status === "Hold", `Vitamin D (1049) is on Hold`);

// Verify exceptions use ERP price
const chickenEntry = entryData.entryRows.find((r) => r.sku === "1008V003");
assert(chickenEntry && chickenEntry.entryPrice === 4.49, `Chicken uses ERP price (£4.49) not PO price`);
assertEq(chickenEntry.status, "Review", `Chicken is flagged for Review`);

// Verify matches use ERP price
const teaEntry = entryData.entryRows.find((r) => r.sku === "1001V001");
assert(teaEntry && teaEntry.entryPrice === 2.49, `Tea uses correct price (£2.49)`);
assertEq(teaEntry.status, "Ready", `Tea is Ready`);

// Staging total should be higher than extraction total (ERP prices are generally higher in this test set)
assert(entryData.totals.totalValue > extraction.metadata.totalValue, `Staging total (£${entryData.totals.totalValue.toFixed(2)}) > PO total (£${extraction.metadata.totalValue.toFixed(2)}) — uses corrected ERP prices`);

console.log(`\n  Breakdown: ${entryData.totals.readyCount} ready / ${entryData.totals.reviewCount} review / ${entryData.totals.holdCount} hold`);
console.log(`  Total value: £${entryData.totals.totalValue.toFixed(2)}`);

// ── Step 7: Credit Note ──

console.log("\nStep 7: Generate credit note");

const creditNote = generateCreditNote(results);
// Count only Exception + Tolerance rows that have both PO and ERP prices (Not in ERP excluded)
const creditableCount = results.rows.filter((r) => (r.status === "Exception" || r.status === "Tolerance") && r.poPrice != null).length;
assertEq(creditNote.creditRows.length, creditableCount, `Credit note covers ${creditNote.creditRows.length} exception/tolerance lines`);
assert(creditNote.totals.totalCredit < 0, `Total credit is negative: £${creditNote.totals.totalCredit.toFixed(2)}`);

// Verify credit only covers exception + tolerance rows, not matched lines
assert(creditNote.creditRows.length < 50, `Credit note is selective (${creditNote.creditRows.length} < 50 total PO lines)`);

// Verify every credit row is an exception or tolerance
const allExOrTol = creditNote.creditRows.every((r) => {
  const match = results.rows.find((rr) => rr.sku === r.sku);
  return match && (match.status === "Exception" || match.status === "Tolerance");
});
assert(allExOrTol, `All credit rows are exception/tolerance lines`);

// ── Step 8: Re-Invoice ──

console.log("\nStep 8: Generate corrected re-invoice");

const reInvoice = generateCorrectedInvoice(results);
const reinvoiceableCount = results.rows.filter((r) => (r.status === "Exception" || r.status === "Tolerance") && r.erpPrice != null).length;
assertEq(reInvoice.invoiceRows.length, reinvoiceableCount, `Re-invoice covers ${reInvoice.invoiceRows.length} exception/tolerance lines`);
assert(reInvoice.totals.totalInvoice > 0, `Total re-invoice: £${reInvoice.totals.totalInvoice.toFixed(2)}`);

// All re-invoice lines should use ERP prices
const allUseErp = reInvoice.invoiceRows.every((r) => r.correctedPrice != null && r.correctedPrice > 0);
assert(allUseErp, `All re-invoice lines use ERP prices`);

// Re-invoice total should be higher than credit (ERP prices are generally higher)
assert(reInvoice.totals.totalInvoice > Math.abs(creditNote.totals.totalCredit), `Re-invoice (£${reInvoice.totals.totalInvoice.toFixed(2)}) > credit (£${Math.abs(creditNote.totals.totalCredit).toFixed(2)}) — ERP prices higher`);

// Net effect: credit + re-invoice = the price correction amount
const netEffect = round(creditNote.totals.totalCredit + reInvoice.totals.totalInvoice);
assert(netEffect > 0, `Net effect is positive: £${netEffect.toFixed(2)} (customer underpaid)`);
console.log(`  Net effect: £${creditNote.totals.totalCredit.toFixed(2)} (credit) + £${reInvoice.totals.totalInvoice.toFixed(2)} (re-invoice) = £${netEffect.toFixed(2)}`);

// ── Step 9: Email Draft ──

console.log("\nStep 9: Generate email draft");

const email = generateEmailDraft(results, "tesco-po-2025-0247.csv");
assert(email.subject.includes("PO-2025-0247"), `Email subject contains PO ref`);
assert(email.subject.includes("exception"), `Email subject mentions exceptions`);
assert(email.body.includes(String(results.summary.matches)), `Email body includes match count`);
assert(email.body.length > 100, `Email body is substantial (${email.body.length} chars)`);

console.log(`\n  Subject: ${email.subject}`);
console.log(`  Body preview: ${email.body.substring(0, 120)}...`);

// ── Step 10: Predict — Price Intelligence ──

console.log("\nStep 10: Price Intelligence (Predict module)");

// Inline toHistoryRecords from src/predict/predict.js
function toHistoryRecords(reconciliation, poRef, date) {
  return reconciliation.rows
    .filter((r) => r.poPrice != null && r.status !== "Not in PO")
    .map((r) => ({
      date,
      poRef,
      sku: r.sku,
      name: r.name || "",
      poPrice: r.poPrice,
      erpPrice: r.erpPrice,
      diff: r.diff,
      status: r.status,
      poQty: r.poQty || 0,
    }));
}

// Inline analyzeHistory core (simplified from src/predict/predict.js)
function avg(arr) { return arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : 0; }
function stdDev(arr) {
  if (arr.length < 2) return 0;
  const mean = avg(arr);
  const variance = arr.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (arr.length - 1);
  return Math.sqrt(variance);
}

function computeTrend(priceSeries) {
  const n = priceSeries.length;
  if (n < 2) return { direction: "insufficient", slope: 0, confidence: 0 };
  const xs = priceSeries.map((_, i) => i);
  const ys = priceSeries.map((p) => p.price);
  const xMean = avg(xs);
  const yMean = avg(ys);
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - xMean) * (ys[i] - yMean);
    den += (xs[i] - xMean) ** 2;
  }
  const slope = den !== 0 ? num / den : 0;
  const slopePct = yMean !== 0 ? (slope / yMean) * 100 : 0;
  let direction = Math.abs(slopePct) < 0.5 ? "stable" : slope > 0 ? "up" : "down";
  return { direction, slope: round(slope), slopePct: round(slopePct) };
}

function analyzeHistory(records) {
  if (!records || records.length === 0) return null;
  const bySku = new Map();
  for (const r of records) {
    const key = r.sku.trim().toUpperCase();
    if (!bySku.has(key)) bySku.set(key, []);
    bySku.get(key).push(r);
  }
  const skuAnalysis = new Map();
  let trendingUp = 0, trendingDown = 0, stable = 0, insufficientData = 0, anomalyCount = 0;
  for (const [sku, skuRecords] of bySku) {
    const sorted = [...skuRecords].sort((a, b) => a.date.localeCompare(b.date));
    const priceSeries = sorted.filter((r) => r.erpPrice != null).map((r) => ({ date: r.date, price: r.erpPrice }));
    const trend = priceSeries.length >= 2 ? computeTrend(priceSeries) : { direction: "insufficient", slope: 0 };
    const exceptionCount = sorted.filter((r) => r.status === "Exception").length;
    const prices = priceSeries.map((p) => p.price);
    let anomaly = false;
    if (prices.length >= 3) {
      const mean = avg(prices.slice(0, -1));
      const sd = stdDev(prices.slice(0, -1));
      if (sd > 0) anomaly = Math.abs((prices[prices.length - 1] - mean) / sd) > 2;
    }
    if (trend.direction === "up") trendingUp++;
    else if (trend.direction === "down") trendingDown++;
    else if (trend.direction === "stable") stable++;
    else insufficientData++;
    if (anomaly) anomalyCount++;
    skuAnalysis.set(sku, { trend, exceptionCount, anomaly, dataPoints: priceSeries.length, name: sorted[0].name });
  }
  const exceptionRecords = records.filter((r) => r.status === "Exception").length;
  const uniqueRuns = new Set(records.map((r) => r.poRef + "|" + r.date)).size;
  return {
    skuAnalysis,
    metrics: {
      totalSKUs: skuAnalysis.size,
      totalRecords: records.length,
      totalRuns: uniqueRuns,
      exceptionRate: round((exceptionRecords / records.length) * 100),
      anomalyCount,
      trendingUp,
      trendingDown,
      stable,
      insufficientData,
    },
  };
}

// Test with single reconciliation run
const run1Records = toHistoryRecords(results, "PO-2025-0247", "2025-01-15");
assert(run1Records.length >= 48, `History: ${run1Records.length} records from reconciliation (≥48)`);

// Verify records have correct shape
const firstRecord = run1Records[0];
assert(firstRecord.date === "2025-01-15", `Record has correct date`);
assert(firstRecord.poRef === "PO-2025-0247", `Record has correct PO ref`);
assert(typeof firstRecord.sku === "string" && firstRecord.sku.length > 0, `Record has SKU`);
assert(typeof firstRecord.poPrice === "number", `Record has numeric PO price`);

// Single-run analysis (limited — most trends need 2+ runs)
const analysis1 = analyzeHistory(run1Records);
assert(analysis1 !== null, `Analysis produced results`);
assert(analysis1.metrics.totalSKUs >= 40, `Tracked ${analysis1.metrics.totalSKUs} unique SKUs (≥40)`);
assertEq(analysis1.metrics.totalRuns, 1, `Single reconciliation run detected`);
assert(analysis1.metrics.insufficientData >= 30, `Most SKUs have insufficient data with 1 run (${analysis1.metrics.insufficientData})`);

// Simulate 3 reconciliation runs with price drift
const run2Records = run1Records.map((r) => ({
  ...r,
  date: "2025-02-15",
  poRef: "PO-2025-0312",
  erpPrice: r.erpPrice != null ? round(r.erpPrice * 1.02) : null, // 2% increase
}));

const run3Records = run1Records.map((r) => ({
  ...r,
  date: "2025-03-15",
  poRef: "PO-2025-0398",
  erpPrice: r.erpPrice != null ? round(r.erpPrice * 1.05) : null, // 5% increase from baseline
}));

const allRecords = [...run1Records, ...run2Records, ...run3Records];
assert(allRecords.length >= 144, `3 runs: ${allRecords.length} total records (≥144)`);

const analysis3 = analyzeHistory(allRecords);
assertEq(analysis3.metrics.totalRuns, 3, `Three reconciliation runs detected`);
assert(analysis3.metrics.trendingUp > 0, `Some SKUs trending up (got ${analysis3.metrics.trendingUp})`);
assert(analysis3.metrics.exceptionRate > 0, `Exception rate > 0% (got ${analysis3.metrics.exceptionRate}%)`);

// Check that a SKU with increasing prices has "up" trend
const teaAnalysis = analysis3.skuAnalysis.get("1001V001");
assert(teaAnalysis, `Tea (1001V001) has analysis data`);
assertEq(teaAnalysis.trend.direction, "up", `Tea trend is "up" with 2%→5% price increase`);
assert(teaAnalysis.dataPoints === 3, `Tea has 3 data points`);

console.log(`\n  Run 1: ${run1Records.length} records`);
console.log(`  Run 2: ${run2Records.length} records (+2% price drift)`);
console.log(`  Run 3: ${run3Records.length} records (+5% price drift)`);
console.log(`  Metrics: ${analysis3.metrics.trendingUp} up, ${analysis3.metrics.trendingDown} down, ${analysis3.metrics.stable} stable`);
console.log(`  Exception rate: ${analysis3.metrics.exceptionRate}%`);

// ── Step 11: Conversational Intelligence ──

console.log("\nStep 11: Conversational Intelligence (AI moat)");

// 11a: SKU Lookup — find records for a specific SKU
function lookupSKU(records, query) {
  const norm = query.trim().toUpperCase();
  return records.filter((r) => {
    const s = r.sku.trim().toUpperCase();
    return s === norm || s.startsWith(norm) || norm.startsWith(s);
  });
}

const teaRecords = lookupSKU(allRecords, "1001V001");
assertEq(teaRecords.length, 3, `SKU lookup: 1001V001 has 3 records across 3 runs`);

// Partial match: "1001" should find 1001V001
const partialRecords = lookupSKU(allRecords, "1001");
assert(partialRecords.length >= 3, `SKU lookup: partial "1001" finds ${partialRecords.length} records (≥3)`);

// Unknown SKU returns nothing
const unknownRecords = lookupSKU(allRecords, "ZZZZ999");
assertEq(unknownRecords.length, 0, `SKU lookup: unknown SKU returns empty`);

// 11b: PO Risk Assessment logic
function assessPORisk(results, historicalExceptionRate) {
  const s = results.summary;
  const exceptionPct = s.total > 0 ? (s.exceptions / s.total) * 100 : 0;
  const hasHighExposure = s.exposure > 100;
  const notInErp = results.rows.filter((r) => r.status === "Not in ERP");

  if (exceptionPct === 0 && notInErp.length === 0) return "ACCEPT";
  if (exceptionPct <= 10 && s.exposure < 50 && notInErp.length === 0) return "ACCEPT_WITH_REVIEW";
  if (exceptionPct <= 30 && !hasHighExposure) return "REVIEW";
  return "ESCALATE";
}

const riskLevel = assessPORisk(results, 0);
// Our test PO has ~22% exceptions (11/50) and moderate exposure — should be REVIEW
assert(riskLevel === "REVIEW" || riskLevel === "ACCEPT_WITH_REVIEW", `PO risk assessment: ${riskLevel} (expected REVIEW or ACCEPT_WITH_REVIEW)`);

// Simulate a clean PO (all matches)
const cleanResults = { summary: { total: 50, matches: 50, tolerances: 0, exceptions: 0, exposure: 0, warnings: 0 }, rows: [] };
assertEq(assessPORisk(cleanResults, 0), "ACCEPT", `Clean PO → ACCEPT`);

// Simulate a bad PO (high exceptions)
const badResults = {
  summary: { total: 50, matches: 10, tolerances: 0, exceptions: 35, exposure: 500, warnings: 5 },
  rows: [{ status: "Not in ERP", sku: "X" }, { status: "Not in ERP", sku: "Y" }],
};
assertEq(assessPORisk(badResults, 0), "ESCALATE", `High-exception PO → ESCALATE`);

// Simulate minor exceptions PO
const minorResults = {
  summary: { total: 100, matches: 95, tolerances: 2, exceptions: 3, exposure: 15, warnings: 0 },
  rows: [],
};
assertEq(assessPORisk(minorResults, 0), "ACCEPT_WITH_REVIEW", `Minor-exception PO → ACCEPT_WITH_REVIEW`);

// 11c: Verify SKU history across runs shows price drift
const teaSorted = teaRecords.sort((a, b) => a.date.localeCompare(b.date));
const teaFirstErp = teaSorted[0].erpPrice;
const teaLastErp = teaSorted[teaSorted.length - 1].erpPrice;
assert(teaLastErp > teaFirstErp, `Tea ERP price drifted up: £${teaFirstErp.toFixed(2)} → £${teaLastErp.toFixed(2)}`);
const teaDriftPct = round(((teaLastErp - teaFirstErp) / teaFirstErp) * 100);
assert(teaDriftPct >= 4 && teaDriftPct <= 6, `Tea drift ~5% (got ${teaDriftPct}%)`);

console.log(`  SKU lookup test: 1001V001 = ${teaRecords.length} records, drift = +${teaDriftPct}%`);
console.log(`  Risk levels: clean=ACCEPT, minor=ACCEPT_WITH_REVIEW, current=REVIEW, bad=ESCALATE`);

// ── Results ──

console.log("\n═══════════════════════════════════════════════════════");
if (failed === 0) {
  console.log(`  ALL ${passed} CHECKS PASSED ✓`);
} else {
  console.log(`  ${passed} passed, ${failed} FAILED ✗`);
}
console.log("═══════════════════════════════════════════════════════\n");

process.exit(failed > 0 ? 1 : 0);
