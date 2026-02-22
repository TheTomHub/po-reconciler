/**
 * Credit note and corrected re-invoice generators.
 *
 * Credit notes and re-invoices only cover exception lines —
 * lines where the PO price doesn't match the ERP price.
 */

/**
 * Generate credit note data from reconciliation results.
 * Credits only EXCEPTION lines at their original PO prices (negative amounts).
 *
 * @param {object} results - { summary, rows } from reconcile()
 * @returns {{ creditRows: object[], totals: { lineCount: number, totalCredit: number } }}
 */
export function generateCreditNote(results) {
  const creditRows = [];
  let totalCredit = 0;

  for (const row of results.rows) {
    // Only credit lines with price exceptions
    if (row.status !== "Exception" && row.status !== "Tolerance") continue;
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
      erpPrice: row.erpPrice,
      diff: row.diff,
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
 * Only includes exception lines, re-invoiced at correct ERP prices.
 *
 * @param {object} results - { summary, rows } from reconcile()
 * @returns {{ invoiceRows: object[], totals: { lineCount: number, totalInvoice: number } }}
 */
export function generateCorrectedInvoice(results) {
  const invoiceRows = [];
  let totalInvoice = 0;

  for (const row of results.rows) {
    // Only re-invoice lines with price exceptions
    if (row.status !== "Exception" && row.status !== "Tolerance") continue;
    if (row.erpPrice == null) continue;

    const qty = row.poQty || 1;
    const correctedPrice = row.erpPrice;
    const lineTotal = round(correctedPrice * qty);

    invoiceRows.push({
      sku: row.sku,
      name: row.name || "",
      qty,
      originalPrice: row.poPrice,
      correctedPrice,
      lineTotal,
      diff: row.diff,
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
