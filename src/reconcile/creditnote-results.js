import { formatCurrency, getCurrencyFormat } from "../utils/format";

/* global Excel */

// Reuse color constants from results.js
const HEADER_BG = "#1F4E79";
const HEADER_FG = "#FFFFFF";
const CREDIT_FG = "#A4262C"; // dark red for negative amounts
const CHANGED_BG = "#FFEB9C"; // yellow highlight for changed prices

/**
 * Write a credit note sheet to the active workbook.
 */
export async function writeCreditNoteSheet(creditData, poFilename) {
  await Excel.run(async (context) => {
    const sheetName = getSheetName("CreditNote");

    // Delete existing sheet with same name if present
    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const sheet = context.workbook.worksheets.add(sheetName);
    const { creditRows, totals } = creditData;

    // --- Summary Section (rows 1-5) ---
    const summaryData = [
      ["Credit Note", ""],
      ["Date", new Date().toLocaleDateString()],
      ["PO Reference", poFilename || "—"],
      ["Total Lines", totals.lineCount],
      ["Total Credit", formatCurrency(totals.totalCredit)],
    ];

    const summaryRange = sheet.getRange("A1:B5");
    summaryRange.values = summaryData;

    // Format summary header
    const summaryTitle = sheet.getRange("A1:B1");
    summaryTitle.merge();
    summaryTitle.format.font.bold = true;
    summaryTitle.format.font.size = 14;
    summaryTitle.format.font.color = HEADER_BG;

    const summaryLabels = sheet.getRange("A2:A5");
    summaryLabels.format.font.bold = true;

    // Highlight total credit row in red
    const totalCreditRow = sheet.getRange("A5:B5");
    totalCreditRow.format.font.color = CREDIT_FG;
    totalCreditRow.format.font.bold = true;

    // --- Table Section (row 7+) ---
    const tableStartRow = 7;
    const headers = ["SKU", "Product Name", "Qty", "Original Price", "Line Total", "Credit Amount"];

    const headerRange = sheet.getRange(`A${tableStartRow}:F${tableStartRow}`);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = HEADER_FG;
    headerRange.format.fill.color = HEADER_BG;

    sheet.freezePanes.freezeRows(tableStartRow);

    // Data rows
    if (creditRows.length > 0) {
      const dataValues = creditRows.map((row) => [
        row.sku,
        row.name,
        row.qty,
        row.originalPrice,
        row.lineTotal,
        row.creditAmount,
      ]);

      const dataStartRow = tableStartRow + 1;
      const dataEndRow = dataStartRow + dataValues.length - 1;
      const dataRange = sheet.getRange(`A${dataStartRow}:F${dataEndRow}`);
      dataRange.values = dataValues;

      // Currency format for price columns (D, E, F)
      for (const col of ["D", "E", "F"]) {
        const priceRange = sheet.getRange(`${col}${dataStartRow}:${col}${dataEndRow}`);
        priceRange.numberFormat = [[getCurrencyFormat()]];
      }

      // Red font for credit amount column (F)
      const creditCol = sheet.getRange(`F${dataStartRow}:F${dataEndRow}`);
      creditCol.format.font.color = CREDIT_FG;

      // Footer row: total credit
      const footerRow = dataEndRow + 1;
      const footerRange = sheet.getRange(`A${footerRow}:F${footerRow}`);
      footerRange.values = [["", "", "", "", "Total Credit:", totals.totalCredit]];
      footerRange.format.font.bold = true;

      const footerPriceCell = sheet.getRange(`F${footerRow}`);
      footerPriceCell.numberFormat = [[getCurrencyFormat()]];
      footerPriceCell.format.font.color = CREDIT_FG;

      // Auto-fit columns
      const fullRange = sheet.getRange(`A1:F${footerRow}`);
      fullRange.format.autofitColumns();
    } else {
      const fullRange = sheet.getRange(`A1:F${tableStartRow}`);
      fullRange.format.autofitColumns();
    }

    sheet.activate();
    await context.sync();
  });
}

/**
 * Write a corrected re-invoice sheet to the active workbook.
 */
export async function writeReInvoiceSheet(invoiceData, poFilename, exceptionCount) {
  await Excel.run(async (context) => {
    const sheetName = getSheetName("ReInvoice");

    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const sheet = context.workbook.worksheets.add(sheetName);
    const { invoiceRows, totals } = invoiceData;

    // --- Summary Section (rows 1-5) ---
    const summaryData = [
      ["Corrected Re-Invoice", ""],
      ["Date", new Date().toLocaleDateString()],
      ["PO Reference", poFilename || "—"],
      ["Corrected Total", formatCurrency(totals.totalInvoice)],
      ["Price Corrections", exceptionCount],
    ];

    const summaryRange = sheet.getRange("A1:B5");
    summaryRange.values = summaryData;

    const summaryTitle = sheet.getRange("A1:B1");
    summaryTitle.merge();
    summaryTitle.format.font.bold = true;
    summaryTitle.format.font.size = 14;
    summaryTitle.format.font.color = HEADER_BG;

    const summaryLabels = sheet.getRange("A2:A5");
    summaryLabels.format.font.bold = true;

    // --- Table Section (row 7+) ---
    const tableStartRow = 7;
    const headers = ["SKU", "Product Name", "Qty", "Corrected Price", "Line Total"];

    const headerRange = sheet.getRange(`A${tableStartRow}:E${tableStartRow}`);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = HEADER_FG;
    headerRange.format.fill.color = HEADER_BG;

    sheet.freezePanes.freezeRows(tableStartRow);

    if (invoiceRows.length > 0) {
      const dataValues = invoiceRows.map((row) => [
        row.sku,
        row.name,
        row.qty,
        row.correctedPrice,
        row.lineTotal,
      ]);

      const dataStartRow = tableStartRow + 1;
      const dataEndRow = dataStartRow + dataValues.length - 1;
      const dataRange = sheet.getRange(`A${dataStartRow}:E${dataEndRow}`);
      dataRange.values = dataValues;

      // Currency format for price columns (D, E)
      for (const col of ["D", "E"]) {
        const priceRange = sheet.getRange(`${col}${dataStartRow}:${col}${dataEndRow}`);
        priceRange.numberFormat = [[getCurrencyFormat()]];
      }

      // Yellow highlight on rows where price changed
      for (let i = 0; i < invoiceRows.length; i++) {
        if (invoiceRows[i].priceChanged) {
          const rowRange = sheet.getRange(`A${dataStartRow + i}:E${dataStartRow + i}`);
          rowRange.format.fill.color = CHANGED_BG;
        }
      }

      // Footer row: total invoice
      const footerRow = dataEndRow + 1;
      const footerRange = sheet.getRange(`A${footerRow}:E${footerRow}`);
      footerRange.values = [["", "", "", "Total Invoice:", totals.totalInvoice]];
      footerRange.format.font.bold = true;

      const footerPriceCell = sheet.getRange(`E${footerRow}`);
      footerPriceCell.numberFormat = [[getCurrencyFormat()]];

      const fullRange = sheet.getRange(`A1:E${footerRow}`);
      fullRange.format.autofitColumns();
    } else {
      const fullRange = sheet.getRange(`A1:E${tableStartRow}`);
      fullRange.format.autofitColumns();
    }

    sheet.activate();
    await context.sync();
  });
}

function getSheetName(prefix) {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${prefix}_${y}-${m}-${d}`;
}
