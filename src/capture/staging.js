import { formatCurrency, getCurrencyFormat } from "../utils/format";

/* global Excel */

const HEADER_BG = "#1F4E79";
const HEADER_FG = "#FFFFFF";
const WARNING_BG = "#FFEB9C";
const PENDING_BG = "#E2EFDA";

const STAGING_HEADERS = [
  "#",
  "SKU",
  "Product Name",
  "Qty",
  "Unit Price",
  "UOM",
  "Line Total",
  "Delivery Date",
  "Status",
];

/**
 * Write a PO staging sheet to the active workbook.
 *
 * This is the standardized output of the Capture module —
 * a clean, review-ready sheet the operator can validate
 * before pushing to ERP via DataLoad or manual entry.
 */
export async function writeStagingSheet(extractionResult) {
  const { stagingRows, metadata, warnings } = extractionResult;

  await Excel.run(async (context) => {
    const sheetName = getStagingSheetName(metadata.poRef);

    // Delete existing sheet with same name if present
    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const sheet = context.workbook.worksheets.add(sheetName);

    // ── Summary Section (rows 1-7) ──
    const summaryData = [
      ["PO Staging Sheet", ""],
      ["PO Reference", metadata.poRef],
      ["Customer", metadata.customer],
      ["Line Items", metadata.lineCount],
      ["Total Value", formatCurrency(metadata.totalValue)],
      ["Warnings", metadata.warningCount],
      ["Extracted", new Date().toLocaleString()],
    ];

    const summaryRange = sheet.getRange("A1:B7");
    summaryRange.values = summaryData;

    // Format summary
    const summaryTitle = sheet.getRange("A1:B1");
    summaryTitle.merge();
    summaryTitle.format.font.bold = true;
    summaryTitle.format.font.size = 14;
    summaryTitle.format.font.color = HEADER_BG;

    const summaryLabels = sheet.getRange("A2:A7");
    summaryLabels.format.font.bold = true;

    // Highlight warnings count if > 0
    if (metadata.warningCount > 0) {
      const warningCell = sheet.getRange("B6");
      warningCell.format.font.color = "#A4262C";
      warningCell.format.font.bold = true;
    }

    // ── Data Table (row 9+) ──
    const tableStartRow = 9;

    const headerRange = sheet.getRange(`A${tableStartRow}:I${tableStartRow}`);
    headerRange.values = [STAGING_HEADERS];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = HEADER_FG;
    headerRange.format.fill.color = HEADER_BG;

    sheet.freezePanes.freezeRows(tableStartRow);

    if (stagingRows.length > 0) {
      // Build warning lookup: line number -> messages
      const warningsByLine = new Map();
      for (const w of warnings) {
        if (!warningsByLine.has(w.line)) warningsByLine.set(w.line, []);
        warningsByLine.get(w.line).push(`${w.field}: ${w.message}`);
      }

      const dataValues = stagingRows.map((row) => [
        row.lineNum,
        row.sku,
        row.name,
        row.qty,
        row.price,
        row.uom,
        row.lineTotal,
        row.deliveryDate,
        warningsByLine.has(row.lineNum) ? "Review" : row.status,
      ]);

      const dataStartRow = tableStartRow + 1;
      const dataEndRow = dataStartRow + dataValues.length - 1;
      const dataRange = sheet.getRange(`A${dataStartRow}:I${dataEndRow}`);
      dataRange.values = dataValues;

      // Currency format for price columns (E, G)
      for (const col of ["E", "G"]) {
        const priceRange = sheet.getRange(`${col}${dataStartRow}:${col}${dataEndRow}`);
        priceRange.numberFormat = [[getCurrencyFormat()]];
      }

      // Date format for delivery date (H)
      const dateRange = sheet.getRange(`H${dataStartRow}:H${dataEndRow}`);
      dateRange.numberFormat = [["yyyy-mm-dd"]];

      // Color-code rows
      for (let i = 0; i < stagingRows.length; i++) {
        const rowRange = sheet.getRange(`A${dataStartRow + i}:I${dataStartRow + i}`);
        if (warningsByLine.has(stagingRows[i].lineNum)) {
          // Warning rows get yellow background
          rowRange.format.fill.color = WARNING_BG;
          // Add comment with warning details
          const statusCell = sheet.getRange(`I${dataStartRow + i}`);
          statusCell.note = warningsByLine.get(stagingRows[i].lineNum).join("\n");
        } else {
          rowRange.format.fill.color = PENDING_BG;
        }
      }

      // Footer row: totals
      const footerRow = dataEndRow + 1;
      const footerRange = sheet.getRange(`A${footerRow}:I${footerRow}`);
      footerRange.values = [["", "", "", stagingRows.reduce((s, r) => s + r.qty, 0), "", "", metadata.totalValue, "", `${stagingRows.length} lines`]];
      footerRange.format.font.bold = true;

      const footerTotalCell = sheet.getRange(`G${footerRow}`);
      footerTotalCell.numberFormat = [[getCurrencyFormat()]];

      // Auto-fit columns
      const fullRange = sheet.getRange(`A1:I${footerRow}`);
      fullRange.format.autofitColumns();
    } else {
      const fullRange = sheet.getRange(`A1:I${tableStartRow}`);
      fullRange.format.autofitColumns();
    }

    // ── Warnings Sheet (if warnings exist) ──
    if (warnings.length > 0) {
      const warningSheetName = sheetName + "_Warnings";
      const existingWarning = context.workbook.worksheets.getItemOrNullObject(warningSheetName);
      await context.sync();
      if (!existingWarning.isNullObject) {
        existingWarning.delete();
        await context.sync();
      }

      const wSheet = context.workbook.worksheets.add(warningSheetName);
      const wHeaders = ["Line", "Field", "Warning"];
      const wHeaderRange = wSheet.getRange("A1:C1");
      wHeaderRange.values = [wHeaders];
      wHeaderRange.format.font.bold = true;
      wHeaderRange.format.font.color = HEADER_FG;
      wHeaderRange.format.fill.color = "#A4262C";

      const wData = warnings.map((w) => [w.line, w.field, w.message]);
      const wDataRange = wSheet.getRange(`A2:C${1 + wData.length}`);
      wDataRange.values = wData;

      const wFullRange = wSheet.getRange(`A1:C${1 + wData.length}`);
      wFullRange.format.autofitColumns();
    }

    sheet.activate();
    await context.sync();
  });
}

function getStagingSheetName(poRef) {
  const clean = (poRef || "PO").replace(/[^a-zA-Z0-9-_]/g, "").slice(0, 20);
  return `Staging_${clean}`;
}
