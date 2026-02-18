/**
 * Staged Entry module — generates ERP-ready data from reconciliation results.
 *
 * Option 3 architecture: Excel as universal staging layer.
 * Operator reviews the staging sheet, then copies/uploads into their ERP
 * (SAP, Oracle, D365, etc.) via DataLoad, LSMW, or manual entry.
 *
 * Input:  reconciliation results from reconcile()
 * Output: { entryRows, totals, metadata }
 */

/**
 * Status values for staging rows.
 * Ready  — price verified, safe to enter
 * Review — price mismatch or exception, needs operator decision
 * Hold   — missing from ERP or warning, cannot enter yet
 */
const STATUS = {
  READY: "Ready",
  REVIEW: "Review",
  HOLD: "Hold",
};

/**
 * Generate ERP staging data from reconciliation results.
 *
 * Uses the corrected (ERP) price where available.
 * Excludes ERP-only rows ("Not in PO") since they weren't ordered.
 *
 * @param {object} results - { summary, rows } from reconcile()
 * @param {object} options
 * @param {string} options.poRef - PO reference number
 * @param {string} options.customer - Customer name
 * @param {string} options.deliveryDate - Delivery date (optional)
 * @returns {{ entryRows, totals, metadata }}
 */
export function generateStagingEntry(results, options = {}) {
  const entryRows = [];
  let totalValue = 0;
  let readyCount = 0;
  let reviewCount = 0;
  let holdCount = 0;

  let lineNum = 0;

  for (const row of results.rows) {
    // Skip ERP-only rows — they weren't on the customer PO
    if (row.status === "Not in PO") continue;
    // Skip rows with no price at all
    if (row.poPrice == null && row.erpPrice == null) continue;

    lineNum++;

    // Determine the price to use and the entry status
    let entryPrice, status, notes;

    switch (row.status) {
      case "Match":
        entryPrice = row.erpPrice ?? row.poPrice;
        status = STATUS.READY;
        notes = "";
        break;

      case "Tolerance":
        entryPrice = row.erpPrice ?? row.poPrice;
        status = STATUS.READY;
        notes = `Within tolerance (diff: ${row.diff})`;
        break;

      case "Exception":
        // Use ERP price (the correct one) but flag for review
        entryPrice = row.erpPrice ?? row.poPrice;
        status = STATUS.REVIEW;
        notes = `Price exception: PO ${row.poPrice} vs ERP ${row.erpPrice} (diff: ${row.diff})`;
        break;

      case "Not in ERP":
        entryPrice = row.poPrice;
        status = STATUS.HOLD;
        notes = "SKU not found in ERP — verify item code";
        break;

      case "Warning":
        entryPrice = row.poPrice ?? 0;
        status = STATUS.HOLD;
        notes = row.action || "Data quality issue";
        break;

      default:
        entryPrice = row.erpPrice ?? row.poPrice ?? 0;
        status = STATUS.REVIEW;
        notes = "";
    }

    const qty = row.poQty || 1;
    const lineTotal = round(entryPrice * qty);
    totalValue = round(totalValue + lineTotal);

    if (status === STATUS.READY) readyCount++;
    else if (status === STATUS.REVIEW) reviewCount++;
    else holdCount++;

    entryRows.push({
      lineNum,
      sku: row.sku,
      erpSku: row.erpSku || "",
      name: row.name || "",
      qty,
      uom: "EA", // Default — future: carry forward from extraction
      entryPrice,
      lineTotal,
      status,
      notes,
    });
  }

  return {
    entryRows,
    totals: {
      lineCount: entryRows.length,
      totalValue,
      readyCount,
      reviewCount,
      holdCount,
    },
    metadata: {
      poRef: options.poRef || "Unknown",
      customer: options.customer || "Unknown",
      deliveryDate: options.deliveryDate || "",
      generatedAt: new Date().toISOString(),
      reconciliationSummary: {
        matches: results.summary.matches,
        tolerances: results.summary.tolerances,
        exceptions: results.summary.exceptions,
        exposure: results.summary.exposure,
      },
    },
  };
}

function round(n) {
  return Math.round(n * 100) / 100;
}
