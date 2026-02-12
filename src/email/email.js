import { formatCurrency } from "../utils/format";

/**
 * Generate email draft from reconciliation results.
 * Returns { subject, body }
 */
export function generateEmailDraft(results, filename) {
  const poNumber = extractPONumber(filename);
  const { summary, rows } = results;

  const exceptionRows = rows.filter(
    (r) => r.status === "Exception" || r.status === "Not in Oracle" || r.status === "Not in PO"
  );

  const topExceptions = exceptionRows.slice(0, 3);

  const subject = `PO ${poNumber} — Reconciliation: ${summary.exceptions} exception${summary.exceptions !== 1 ? "s" : ""} found`;

  const topLines = topExceptions.map((r) => {
    if (r.status === "Not in Oracle") {
      return `  - SKU ${r.sku}: Not found in Oracle`;
    }
    if (r.status === "Not in PO") {
      return `  - SKU ${r.sku}: In Oracle but not on PO`;
    }
    return `  - SKU ${r.sku}: PO ${formatCurrency(r.poPrice)} vs Oracle ${formatCurrency(r.oraclePrice)} (diff: ${formatCurrency(r.diff)})`;
  });

  const body = `Hi Team,

PO reconciliation for ${poNumber} has been completed. Please see the summary below:

  Total line items:    ${summary.total}
  Perfect matches:     ${summary.matches}
  Within tolerance:    ${summary.tolerances}
  Exceptions:          ${summary.exceptions}
  Total $ exposure:    ${formatCurrency(summary.exposure)}

${summary.exceptions > 0 ? `Top exceptions:\n${topLines.join("\n")}\n${exceptionRows.length > 3 ? `  ... and ${exceptionRows.length - 3} more\n` : ""}` : "All items matched within tolerance."}
Full reconciliation details are in the Recon sheet attached to this workbook.

Please review and advise on next steps.

Best regards`;

  return { subject, body };
}

/**
 * Build a mailto: link from draft.
 */
export function buildMailtoLink(draft) {
  const params = new URLSearchParams({
    subject: draft.subject,
    body: draft.body,
  });
  return `mailto:?${params.toString()}`;
}

/**
 * Extract PO number from filename.
 * Tries patterns like "PO-12345", "PO_12345", "PO 12345", or just the filename stem.
 */
function extractPONumber(filename) {
  if (!filename) return "Unknown";

  // Remove extension
  const stem = filename.replace(/\.[^.]+$/, "");

  // Try common PO patterns
  const poMatch = stem.match(/(?:PO[-_ ]?)(\d+[-\w]*)/i);
  if (poMatch) return `PO-${poMatch[1]}`;

  // Fallback: use filename stem
  return stem;
}
