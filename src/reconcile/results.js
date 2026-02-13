import { formatCurrency } from "../utils/format";

/* global Excel */

const HEADER_BG = "#1F4E79";
const HEADER_FG = "#FFFFFF";
const EXCEPTION_BG = "#FFC7CE";
const TOLERANCE_BG = "#FFEB9C";
const MATCH_BG = "#C6EFCE";
const WARNING_BG = "#F2F2F2";

const TABLE_HEADERS = [
  "Status",
  "SKU",
  "Product Name",
  "Oracle $",
  "PO $",
  "Difference",
  "% Diff",
  "Action",
];

/**
 * Write reconciliation results to a new Excel sheet.
 */
export async function writeResultsSheet(results, tolerance) {
  await Excel.run(async (context) => {
    const sheetName = getSheetName();

    // Delete existing sheet with same name if present
    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    // Create new sheet
    const sheet = context.workbook.worksheets.add(sheetName);

    // --- Summary Section (rows 1-7) ---
    const summaryData = [
      ["PO Reconciliation Summary", ""],
      ["Total Line Items", results.summary.total],
      ["Perfect Matches", results.summary.matches],
      ["Within Tolerance", results.summary.tolerances],
      ["Exceptions", results.summary.exceptions],
      ["Total $ Exposure", formatCurrency(results.summary.exposure)],
      ["Tolerance Used", formatCurrency(tolerance)],
      ["Timestamp", new Date().toLocaleString()],
    ];

    const summaryRange = sheet.getRange("A1:B8");
    summaryRange.values = summaryData;

    // Format summary header
    const summaryTitle = sheet.getRange("A1:B1");
    summaryTitle.merge();
    summaryTitle.format.font.bold = true;
    summaryTitle.format.font.size = 14;
    summaryTitle.format.font.color = HEADER_BG;

    // Format summary labels
    const summaryLabels = sheet.getRange("A2:A8");
    summaryLabels.format.font.bold = true;

    // Highlight exceptions row
    const exceptionsRow = sheet.getRange("A5:B5");
    exceptionsRow.format.font.color = "#A4262C";
    exceptionsRow.format.font.bold = true;

    // --- Table Section (row 10+) ---
    const tableStartRow = 10;

    // Header row
    const headerRange = sheet.getRange(`A${tableStartRow}:H${tableStartRow}`);
    headerRange.values = [TABLE_HEADERS];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = HEADER_FG;
    headerRange.format.fill.color = HEADER_BG;

    // Freeze panes — header row and above
    sheet.freezePanes.freezeRows(tableStartRow);

    // Data rows
    if (results.rows.length > 0) {
      const dataValues = results.rows.map((row) => [
        row.duplicate ? `${row.status} (DUP)` : row.status,
        row.oracleSku ? `${row.sku} → ${row.oracleSku}` : row.sku,
        row.name || "",
        row.oraclePrice != null ? row.oraclePrice : "",
        row.poPrice != null ? row.poPrice : "",
        row.diff != null ? row.diff : "",
        row.pctDiff != null ? `${row.pctDiff}%` : "",
        row.action,
      ]);

      const dataStartRow = tableStartRow + 1;
      const dataEndRow = dataStartRow + dataValues.length - 1;
      const dataRange = sheet.getRange(`A${dataStartRow}:H${dataEndRow}`);
      dataRange.values = dataValues;

      // Conditional formatting per row
      for (let i = 0; i < results.rows.length; i++) {
        const rowRange = sheet.getRange(`A${dataStartRow + i}:H${dataStartRow + i}`);
        const status = results.rows[i].status;

        if (status === "Exception" || status === "Not in Oracle" || status === "Not in PO") {
          rowRange.format.fill.color = EXCEPTION_BG;
        } else if (status === "Tolerance") {
          rowRange.format.fill.color = TOLERANCE_BG;
        } else if (status === "Match") {
          rowRange.format.fill.color = MATCH_BG;
        } else if (status === "Warning") {
          rowRange.format.fill.color = WARNING_BG;
        }
      }

      // Currency format for price columns (D, E, F)
      const priceColumns = ["D", "E", "F"];
      for (const col of priceColumns) {
        const priceRange = sheet.getRange(`${col}${dataStartRow}:${col}${dataEndRow}`);
        priceRange.numberFormat = [["$#,##0.00"]];
      }
    }

    // Auto-fit columns
    const fullRange = sheet.getRange(`A1:H${tableStartRow + results.rows.length}`);
    fullRange.format.autofitColumns();

    // Activate the new sheet
    sheet.activate();

    await context.sync();
  });
}

function getSheetName() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `Recon_${y}-${m}-${d}`;
}
