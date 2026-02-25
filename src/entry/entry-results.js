import { formatCurrency, getCurrencyFormat } from "../utils/format";

/* global Excel */

const HEADER_BG = "#1F4E79";
const HEADER_FG = "#FFFFFF";
const READY_BG = "#E2EFDA";   // green
const REVIEW_BG = "#FFEB9C";  // yellow
const HOLD_BG = "#FFC7CE";    // red

/**
 * Write an ERP staging entry sheet to the active workbook.
 */
export async function writeEntrySheet(entryData) {
  await Excel.run(async (context) => {
    const sheetName = getSheetName(entryData.metadata.poRef);

    const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }

    const sheet = context.workbook.worksheets.add(sheetName);
    const { entryRows, totals, metadata } = entryData;

    // --- Summary Section (rows 1-8) ---
    const summaryData = [
      ["ERP Staging Sheet", ""],
      ["PO Reference", metadata.poRef],
      ["Customer", metadata.customer],
      ["Delivery Date", metadata.deliveryDate || "—"],
      ["Generated", new Date().toLocaleDateString()],
      ["Total Lines", totals.lineCount],
      ["Total Value", formatCurrency(totals.totalValue)],
      ["Status", `${totals.readyCount} ready, ${totals.reviewCount} review, ${totals.holdCount} hold`],
    ];

    const summaryRange = sheet.getRange("A1:B8");
    summaryRange.values = summaryData;

    const summaryTitle = sheet.getRange("A1:B1");
    summaryTitle.merge();
    summaryTitle.format.font.bold = true;
    summaryTitle.format.font.size = 14;
    summaryTitle.format.font.color = HEADER_BG;

    const summaryLabels = sheet.getRange("A2:A8");
    summaryLabels.format.font.bold = true;

    // --- Table Section (row 10+) ---
    const tableStartRow = 10;
    const headers = ["Line", "SKU", "ERP SKU", "Description", "Qty", "UOM", "Unit Price", "Line Total", "Status", "Notes"];
    const colCount = headers.length; // A through J
    const lastCol = String.fromCharCode(64 + colCount); // J

    const headerRange = sheet.getRange(`A${tableStartRow}:${lastCol}${tableStartRow}`);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = HEADER_FG;
    headerRange.format.fill.color = HEADER_BG;

    sheet.freezePanes.freezeRows(tableStartRow);

    if (entryRows.length > 0) {
      const dataValues = entryRows.map((row) => [
        row.lineNum,
        row.sku,
        row.erpSku,
        row.name,
        row.qty,
        row.uom,
        row.entryPrice,
        row.lineTotal,
        row.status,
        row.notes,
      ]);

      const dataStartRow = tableStartRow + 1;
      const dataEndRow = dataStartRow + dataValues.length - 1;
      const dataRange = sheet.getRange(`A${dataStartRow}:${lastCol}${dataEndRow}`);
      dataRange.values = dataValues;

      // Currency format for price columns (G = Unit Price, H = Line Total)
      for (const col of ["G", "H"]) {
        const priceRange = sheet.getRange(`${col}${dataStartRow}:${col}${dataEndRow}`);
        priceRange.numberFormat = [[getCurrencyFormat()]];
      }

      // Color-code rows by status
      for (let i = 0; i < entryRows.length; i++) {
        const rowRange = sheet.getRange(`A${dataStartRow + i}:${lastCol}${dataStartRow + i}`);
        switch (entryRows[i].status) {
          case "Ready":
            rowRange.format.fill.color = READY_BG;
            break;
          case "Review":
            rowRange.format.fill.color = REVIEW_BG;
            break;
          case "Hold":
            rowRange.format.fill.color = HOLD_BG;
            break;
        }
      }

      // Footer row: totals
      const footerRow = dataEndRow + 2;
      const footerRange = sheet.getRange(`A${footerRow}:${lastCol}${footerRow}`);
      footerRange.values = [["", "", "", "", "", "", "Total:", totals.totalValue, "", ""]];
      footerRange.format.font.bold = true;

      const footerPriceCell = sheet.getRange(`H${footerRow}`);
      footerPriceCell.numberFormat = [[getCurrencyFormat()]];

      // Legend row
      const legendRow = footerRow + 2;
      sheet.getRange(`A${legendRow}`).values = [["Legend:"]];
      sheet.getRange(`A${legendRow}`).format.font.bold = true;
      sheet.getRange(`A${legendRow + 1}:B${legendRow + 1}`).values = [["Ready", "Price verified — safe to enter into ERP"]];
      sheet.getRange(`A${legendRow + 1}:B${legendRow + 1}`).format.fill.color = READY_BG;
      sheet.getRange(`A${legendRow + 2}:B${legendRow + 2}`).values = [["Review", "Price mismatch — operator decision needed"]];
      sheet.getRange(`A${legendRow + 2}:B${legendRow + 2}`).format.fill.color = REVIEW_BG;
      sheet.getRange(`A${legendRow + 3}:B${legendRow + 3}`).values = [["Hold", "Missing from ERP or data issue — cannot enter"]];
      sheet.getRange(`A${legendRow + 3}:B${legendRow + 3}`).format.fill.color = HOLD_BG;

      // Auto-fit columns
      const fullRange = sheet.getRange(`A1:${lastCol}${legendRow + 3}`);
      fullRange.format.autofitColumns();
    } else {
      const fullRange = sheet.getRange(`A1:${lastCol}${tableStartRow}`);
      fullRange.format.autofitColumns();
    }

    sheet.activate();
    await context.sync();
  });
}

function getSheetName(poRef) {
  const clean = (poRef || "Entry").replace(/[\\/*?\[\]:]/g, "").trim().slice(0, 20);
  return `ERP Entry ${clean}`;
}
