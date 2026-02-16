/**
 * Credit note and corrected re-invoice generators.
 *
 * Company policy: when a PO has price exceptions, credit ALL lines
 * (not just exceptions), then re-invoice with corrected prices.
 */

/**
 * Generate credit note data from reconciliation results.
 * Credits ALL PO lines at their original PO prices (negative amounts).
 *
 * @param {object} results - { summary, rows } from reconcile()
 * @returns {{ creditRows: object[], totals: { lineCount: number, totalCredit: number } }}
 */
export function generateCreditNote(results) {
  const creditRows = [];
  let totalCredit = 0;

  for (const row of results.rows) {
    // Skip rows that aren't actual PO lines (ERP-only rows)
    if (row.status === "Not in PO") continue;
    // Skip rows with no usable price
    if (row.poPrice == null) continue;

    const qty = row.poQty || 1;
    const lineTotal = round(row.poPrice * qty);
    const creditAmount = round(-lineTotal);

    creditRows.push({
      sku: row.sku,
      name: row.name || "",
      qty,
      originalPrice: row.poPrice,
      lineTotal,
      creditAmount,
    });

    totalCredit = round(totalCredit + creditAmount);
  }

  return {
    creditRows,
    totals: {
      lineCount: creditRows.length,
      totalCredit,
    },
  };
}

/**
 * Generate corrected re-invoice data from reconciliation results.
 * Uses ERP (correct) prices for all lines. Falls back to PO price
 * when ERP price is unavailable.
 *
 * @param {object} results - { summary, rows } from reconcile()
 * @returns {{ invoiceRows: object[], totals: { lineCount: number, totalInvoice: number } }}
 */
export function generateCorrectedInvoice(results) {
  const invoiceRows = [];
  let totalInvoice = 0;

  for (const row of results.rows) {
    if (row.status === "Not in PO") continue;
    if (row.poPrice == null && row.erpPrice == null) continue;

    const qty = row.poQty || 1;
    const correctedPrice = row.erpPrice != null ? row.erpPrice : row.poPrice;
    const lineTotal = round(correctedPrice * qty);
    const priceChanged = row.erpPrice != null && row.poPrice != null && row.erpPrice !== row.poPrice;

    invoiceRows.push({
      sku: row.sku,
      name: row.name || "",
      qty,
      correctedPrice,
      lineTotal,
      priceChanged,
    });

    totalInvoice = round(totalInvoice + lineTotal);
  }

  return {
    invoiceRows,
    totals: {
      lineCount: invoiceRows.length,
      totalInvoice,
    },
  };
}

function round(n) {
  return Math.round(n * 100) / 100;
}
