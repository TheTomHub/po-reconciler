import { detectColumns } from "../reconcile/detector";
import { parseNumber } from "../utils/format";

/**
 * Extended column detection for Capture module.
 * Adds delivery date, PO reference, customer, and UOM detection
 * on top of the base SKU/price/name/qty detection.
 */

const DATE_ALIASES = [
  "delivery date",
  "required date",
  "ship date",
  "due date",
  "date required",
  "req date",
  "del date",
  "arrival date",
  "need by",
  "need date",
  "eta",
];

const PO_REF_ALIASES = [
  "po number",
  "po no",
  "po ref",
  "po reference",
  "purchase order",
  "order number",
  "order no",
  "order ref",
  "customer po",
  "po#",
  "order #",
];

const CUSTOMER_ALIASES = [
  "customer",
  "customer name",
  "sold to",
  "bill to",
  "buyer",
  "account",
  "account name",
  "company",
  "ship to name",
];

const UOM_ALIASES = [
  "uom",
  "unit of measure",
  "unit",
  "pack size",
  "pack",
  "case size",
  "inner",
  "outer",
];

const LINE_TOTAL_ALIASES = [
  "line total",
  "total",
  "ext amount",
  "extended",
  "line amount",
  "net amount",
  "value",
  "ext price",
];

/**
 * Detect all available columns including extended fields.
 * Returns the base columns plus optional extended fields.
 */
export function detectAllColumns(headers) {
  const base = detectColumns(headers);

  return {
    ...base,
    deliveryDate: findColumn(headers, DATE_ALIASES),
    poRef: findColumn(headers, PO_REF_ALIASES),
    customer: findColumn(headers, CUSTOMER_ALIASES),
    uom: findColumn(headers, UOM_ALIASES),
    lineTotal: findColumn(headers, LINE_TOTAL_ALIASES),
  };
}

/**
 * Extract and normalize PO data into standardized staging rows.
 *
 * Input:  raw parsed data { headers, rows } from parser.js
 * Output: { stagingRows, metadata, warnings }
 */
export function extractPOData(parsedData, detectedColumns) {
  const cols = detectedColumns || detectAllColumns(parsedData.headers);

  if (!cols.sku) {
    throw new Error("Cannot find SKU column. Available headers: " + parsedData.headers.join(", "));
  }

  const stagingRows = [];
  const warnings = [];
  let poRef = null;
  let customer = null;

  for (let i = 0; i < parsedData.rows.length; i++) {
    const raw = parsedData.rows[i];
    const lineNum = i + 1;

    // Extract SKU
    const sku = normalizeValue(raw[cols.sku]);
    if (!sku) {
      warnings.push({ line: lineNum, field: "SKU", message: "Empty SKU — row skipped" });
      continue;
    }

    // Extract price
    const rawPrice = cols.price ? raw[cols.price] : null;
    const price = rawPrice != null ? parseNumber(rawPrice) : null;
    if (cols.price && price === null && rawPrice) {
      warnings.push({ line: lineNum, field: "Price", message: `Non-numeric price: "${rawPrice}"` });
    }

    // Extract quantity
    const rawQty = cols.qty ? raw[cols.qty] : null;
    const qty = rawQty != null ? parseNumber(rawQty) : null;
    if (cols.qty && qty === null && rawQty) {
      warnings.push({ line: lineNum, field: "Qty", message: `Non-numeric quantity: "${rawQty}"` });
    }
    if (qty !== null && qty <= 0) {
      warnings.push({ line: lineNum, field: "Qty", message: `Zero or negative quantity: ${qty}` });
    }

    // Extract name
    const name = cols.name ? normalizeValue(raw[cols.name]) : "";

    // Extract delivery date
    const deliveryDate = cols.deliveryDate ? normalizeDate(raw[cols.deliveryDate]) : null;

    // Extract UOM
    const uom = cols.uom ? normalizeValue(raw[cols.uom]) : "";

    // Extract line total
    const rawLineTotal = cols.lineTotal ? raw[cols.lineTotal] : null;
    let lineTotal = rawLineTotal != null ? parseNumber(rawLineTotal) : null;

    // Calculate line total if not provided
    if (lineTotal === null && price !== null && qty !== null) {
      lineTotal = Math.round(price * qty * 100) / 100;
    }

    // Cross-check line total
    if (lineTotal !== null && price !== null && qty !== null) {
      const expected = Math.round(price * qty * 100) / 100;
      if (Math.abs(lineTotal - expected) > 0.02) {
        warnings.push({
          line: lineNum,
          field: "Line Total",
          message: `Mismatch: ${lineTotal} vs expected ${expected} (price × qty)`,
        });
      }
    }

    // Extract PO ref from first row that has it
    if (!poRef && cols.poRef) {
      const val = normalizeValue(raw[cols.poRef]);
      if (val) poRef = val;
    }

    // Extract customer from first row that has it
    if (!customer && cols.customer) {
      const val = normalizeValue(raw[cols.customer]);
      if (val) customer = val;
    }

    stagingRows.push({
      lineNum,
      sku,
      name,
      qty: qty ?? 1,
      price: price ?? 0,
      uom,
      lineTotal: lineTotal ?? 0,
      deliveryDate: deliveryDate || "",
      status: "Pending", // default status for staging
    });
  }

  // Flag large quantity outliers (>3x median)
  const quantities = stagingRows.map((r) => r.qty).filter((q) => q > 0).sort((a, b) => a - b);
  if (quantities.length >= 5) {
    const median = quantities[Math.floor(quantities.length / 2)];
    const threshold = median * 3;
    for (const row of stagingRows) {
      if (row.qty > threshold) {
        warnings.push({
          line: row.lineNum,
          field: "Qty",
          message: `Unusually large quantity: ${row.qty} (median: ${median})`,
        });
      }
    }
  }

  // Summary metadata
  const totalValue = stagingRows.reduce((sum, r) => sum + r.lineTotal, 0);
  const metadata = {
    lineCount: stagingRows.length,
    totalValue: Math.round(totalValue * 100) / 100,
    poRef: poRef || "Unknown",
    customer: customer || "Unknown",
    detectedFields: Object.entries(cols)
      .filter(([, v]) => v !== null)
      .map(([k]) => k),
    warningCount: warnings.length,
    extractedAt: new Date().toISOString(),
  };

  return { stagingRows, metadata, warnings };
}

// ── Helpers ──

function normalizeValue(val) {
  if (val == null) return "";
  return String(val).trim();
}

function normalizeDate(val) {
  if (val == null || val === "") return null;
  const str = String(val).trim();

  // Try native Date parse
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return d.toISOString().split("T")[0]; // YYYY-MM-DD
  }

  // Try DD/MM/YYYY (UK format)
  const ukMatch = str.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/);
  if (ukMatch) {
    const day = ukMatch[1].padStart(2, "0");
    const month = ukMatch[2].padStart(2, "0");
    let year = ukMatch[3];
    if (year.length === 2) year = "20" + year;
    return `${year}-${month}-${day}`;
  }

  return str; // return as-is if unparseable
}

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
