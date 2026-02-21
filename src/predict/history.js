/**
 * PriceHistory sheet — Office.js reader/writer.
 *
 * Maintains a running log of reconciliation results across PO runs.
 * Each reconciliation appends rows; the sheet accumulates over time.
 */

const SHEET_NAME = "PriceHistory";
const HEADERS = ["Date", "PO Ref", "SKU", "Product Name", "PO Price", "ERP Price", "Diff", "Status", "Qty"];

const HEADER_BG = "#1f4e79";
const HEADER_FG = "#ffffff";

/**
 * Append history records to the PriceHistory sheet.
 * Creates the sheet if it doesn't exist.
 *
 * @param {object[]} records - From toHistoryRecords()
 */
export async function appendHistory(records) {
  if (!records || records.length === 0) return;

  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    let sheet;

    try {
      sheet = sheets.getItem(SHEET_NAME);
      sheet.load("name");
      await ctx.sync();
    } catch {
      // Sheet doesn't exist — create it with headers
      sheet = sheets.add(SHEET_NAME);

      const headerRange = sheet.getRangeByIndexes(0, 0, 1, HEADERS.length);
      headerRange.values = [HEADERS];
      headerRange.format.font.bold = true;
      headerRange.format.font.color = HEADER_FG;
      headerRange.format.fill.color = HEADER_BG;
      headerRange.format.horizontalAlignment = "Center";

      // Set column widths
      const widths = [100, 130, 110, 220, 85, 85, 75, 85, 60];
      for (let i = 0; i < widths.length; i++) {
        sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = widths[i];
      }

      await ctx.sync();
    }

    // Find next empty row
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("rowCount");
    await ctx.sync();

    const startRow = usedRange.isNullObject ? 1 : usedRange.rowCount;

    // Write records
    const values = records.map((r) => [
      r.date,
      r.poRef,
      r.sku,
      r.name,
      r.poPrice,
      r.erpPrice != null ? r.erpPrice : "",
      r.diff != null ? r.diff : "",
      r.status,
      r.poQty || "",
    ]);

    const dataRange = sheet.getRangeByIndexes(startRow, 0, values.length, HEADERS.length);
    dataRange.values = values;

    // Format number columns
    const priceFormat = "£#,##0.00";
    sheet.getRangeByIndexes(startRow, 4, values.length, 1).numberFormat = values.map(() => [priceFormat]); // PO Price
    sheet.getRangeByIndexes(startRow, 5, values.length, 1).numberFormat = values.map(() => [priceFormat]); // ERP Price
    sheet.getRangeByIndexes(startRow, 6, values.length, 1).numberFormat = values.map(() => [priceFormat]); // Diff

    await ctx.sync();
  });
}

/**
 * Read all historical records from the PriceHistory sheet.
 *
 * @returns {object[]} Array of record objects, or empty array if no history.
 */
export async function readHistory() {
  let records = [];

  await Excel.run(async (ctx) => {
    let sheet;
    try {
      sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      sheet.load("name");
      await ctx.sync();
    } catch {
      // No history sheet yet
      return;
    }

    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load(["values", "rowCount"]);
    await ctx.sync();

    if (usedRange.isNullObject || usedRange.rowCount < 2) return;

    const allValues = usedRange.values;
    // Skip header row
    for (let i = 1; i < allValues.length; i++) {
      const row = allValues[i];
      // Skip empty rows
      if (!row[2]) continue; // SKU is required

      records.push({
        date: String(row[0] || ""),
        poRef: String(row[1] || ""),
        sku: String(row[2] || ""),
        name: String(row[3] || ""),
        poPrice: parseNum(row[4]),
        erpPrice: parseNum(row[5]),
        diff: parseNum(row[6]),
        status: String(row[7] || ""),
        poQty: parseNum(row[8]) || 0,
      });
    }
  });

  return records;
}

/**
 * Get a summary of history (record count, date range, unique SKUs).
 * Lightweight check without full analysis.
 */
export async function getHistorySummary() {
  const records = await readHistory();
  if (records.length === 0) {
    return { hasHistory: false, recordCount: 0, skuCount: 0, runs: 0 };
  }

  const dates = records.map((r) => r.date).filter(Boolean).sort();
  const uniqueSkus = new Set(records.map((r) => r.sku.toUpperCase())).size;
  const uniqueRuns = new Set(records.map((r) => r.poRef + "|" + r.date)).size;

  return {
    hasHistory: true,
    recordCount: records.length,
    skuCount: uniqueSkus,
    runs: uniqueRuns,
    firstDate: dates[0] || "",
    lastDate: dates[dates.length - 1] || "",
  };
}

function parseNum(val) {
  if (val == null || val === "") return null;
  const n = Number(val);
  return isNaN(n) ? null : n;
}
