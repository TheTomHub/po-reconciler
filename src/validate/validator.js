/**
 * Validate module — data quality gate between Capture and Reconcile.
 *
 * Runs a set of validation rules against extracted staging rows.
 * Returns { valid, errors, warnings, info, summary }.
 *
 * Pipeline behaviour:
 *   errors   → block reconciliation (must be fixed)
 *   warnings → flag for review (operator can proceed)
 *   info     → informational only
 */

import { parseNumber } from "../utils/format";

// ── Rule: No data ──

function checkNoData(rows) {
  if (rows.length === 0) {
    return [{ severity: "error", rule: "no-data", message: "No data rows found. Check the source file." }];
  }
  return [];
}

// ── Rule: Duplicate SKUs ──

function checkDuplicateSkus(rows) {
  const skuMap = new Map();
  for (const row of rows) {
    const norm = row.sku.toUpperCase().trim();
    if (!skuMap.has(norm)) skuMap.set(norm, []);
    skuMap.get(norm).push(row.lineNum);
  }

  const results = [];
  for (const [sku, lines] of skuMap) {
    if (lines.length > 1) {
      results.push({
        severity: "warning",
        rule: "duplicate-sku",
        field: "SKU",
        lines,
        message: `Duplicate SKU "${sku}" on lines ${lines.join(", ")}`,
      });
    }
  }
  return results;
}

// ── Rule: Missing prices ──

function checkMissingPrices(rows) {
  const missing = rows.filter((r) => r.price === 0 || r.price === null);
  if (missing.length === 0) return [];

  if (missing.length === rows.length) {
    return [{ severity: "error", rule: "no-prices", message: "No rows have a valid price. Check column detection." }];
  }

  return missing.map((r) => ({
    severity: "error",
    rule: "missing-price",
    field: "Price",
    lines: [r.lineNum],
    message: `Line ${r.lineNum} (${r.sku}): missing or zero price`,
  }));
}

// ── Rule: Zero / negative quantities ──

function checkQuantityIssues(rows) {
  const results = [];
  for (const row of rows) {
    if (row.qty <= 0) {
      results.push({
        severity: "error",
        rule: "invalid-qty",
        field: "Qty",
        lines: [row.lineNum],
        message: `Line ${row.lineNum} (${row.sku}): quantity is ${row.qty}`,
      });
    }
  }
  return results;
}

// ── Rule: Price range sanity ──

const MIN_PRICE = 0.01;
const MAX_PRICE = 9999.99;

function checkPriceRange(rows) {
  const results = [];
  for (const row of rows) {
    if (row.price > 0 && (row.price < MIN_PRICE || row.price > MAX_PRICE)) {
      results.push({
        severity: "warning",
        rule: "price-range",
        field: "Price",
        lines: [row.lineNum],
        message: `Line ${row.lineNum} (${row.sku}): price ${row.price} outside expected range (${MIN_PRICE}–${MAX_PRICE})`,
      });
    }
  }
  return results;
}

// ── Rule: Quantity outliers ──

function checkQuantityOutliers(rows) {
  const quantities = rows.map((r) => r.qty).filter((q) => q > 0).sort((a, b) => a - b);
  if (quantities.length < 5) return [];

  const median = quantities[Math.floor(quantities.length / 2)];
  const threshold = median * 5;

  const results = [];
  for (const row of rows) {
    if (row.qty > threshold) {
      results.push({
        severity: "warning",
        rule: "qty-outlier",
        field: "Qty",
        lines: [row.lineNum],
        message: `Line ${row.lineNum} (${row.sku}): quantity ${row.qty} is unusually large (median: ${median})`,
      });
    }
  }
  return results;
}

// ── Rule: Line total consistency ──

function checkLineTotals(rows) {
  const results = [];
  for (const row of rows) {
    if (row.lineTotal > 0 && row.price > 0 && row.qty > 0) {
      const expected = Math.round(row.price * row.qty * 100) / 100;
      const diff = Math.abs(row.lineTotal - expected);
      if (diff > 0.02) {
        results.push({
          severity: "warning",
          rule: "line-total-mismatch",
          field: "Line Total",
          lines: [row.lineNum],
          message: `Line ${row.lineNum} (${row.sku}): total ${row.lineTotal} ≠ price × qty (${expected})`,
        });
      }
    }
  }
  return results;
}

// ── Rule: Missing product names ──

function checkMissingNames(rows) {
  const missing = rows.filter((r) => !r.name || r.name.trim() === "");
  if (missing.length === 0) return [];

  // Only warn if some rows have names (partial coverage)
  const hasNames = rows.some((r) => r.name && r.name.trim());
  if (!hasNames) {
    return [{ severity: "info", rule: "no-names", message: "No product names detected — consider adding a Description column." }];
  }

  return [{
    severity: "info",
    rule: "missing-names",
    lines: missing.map((r) => r.lineNum),
    message: `${missing.length} row(s) missing product name`,
  }];
}

// ── Rule: SKU format consistency ──

function checkSkuFormat(rows) {
  // Check if SKUs follow a consistent pattern
  const skus = rows.map((r) => r.sku);
  const hasLetters = skus.filter((s) => /[a-zA-Z]/.test(s));
  const pureNumeric = skus.filter((s) => /^\d+$/.test(s));

  // Mixed formats: some have letters, some are pure numeric
  if (hasLetters.length > 0 && pureNumeric.length > 0 && pureNumeric.length <= rows.length * 0.3) {
    return [{
      severity: "info",
      rule: "mixed-sku-format",
      message: `Mixed SKU formats: ${hasLetters.length} alphanumeric, ${pureNumeric.length} numeric-only. This may affect ERP matching.`,
    }];
  }
  return [];
}

// ── Main validator ──

export function validate(stagingRows) {
  const all = [
    ...checkNoData(stagingRows),
    ...checkDuplicateSkus(stagingRows),
    ...checkMissingPrices(stagingRows),
    ...checkQuantityIssues(stagingRows),
    ...checkPriceRange(stagingRows),
    ...checkQuantityOutliers(stagingRows),
    ...checkLineTotals(stagingRows),
    ...checkMissingNames(stagingRows),
    ...checkSkuFormat(stagingRows),
  ];

  const errors = all.filter((r) => r.severity === "error");
  const warnings = all.filter((r) => r.severity === "warning");
  const info = all.filter((r) => r.severity === "info");

  return {
    valid: errors.length === 0,
    errors,
    warnings,
    info,
    summary: {
      total: all.length,
      errors: errors.length,
      warnings: warnings.length,
      info: info.length,
    },
  };
}

/**
 * Format validation results as a human-readable string.
 * Used by both the taskpane UI and the Copilot agent.
 */
export function formatValidationReport(validation) {
  const lines = [];

  if (validation.valid) {
    lines.push("Validation passed.");
  } else {
    lines.push(`Validation failed — ${validation.summary.errors} error(s) must be resolved.`);
  }

  if (validation.summary.warnings > 0) {
    lines.push(`${validation.summary.warnings} warning(s) to review.`);
  }

  if (validation.errors.length > 0) {
    lines.push("");
    lines.push("Errors:");
    for (const e of validation.errors) {
      lines.push(`  ✗ ${e.message}`);
    }
  }

  if (validation.warnings.length > 0) {
    lines.push("");
    lines.push("Warnings:");
    for (const w of validation.warnings) {
      lines.push(`  ! ${w.message}`);
    }
  }

  if (validation.info.length > 0) {
    lines.push("");
    lines.push("Info:");
    for (const i of validation.info) {
      lines.push(`  ○ ${i.message}`);
    }
  }

  return lines.join("\n");
}
