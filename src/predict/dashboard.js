/**
 * Dashboard sheet — Office.js writer.
 *
 * Creates a color-coded Dashboard sheet with:
 *  - Summary metrics (leading indicators)
 *  - Top movers table (biggest recent price changes)
 *  - Watch list (SKUs needing attention)
 *  - Exception rate and trend breakdown
 */

const SHEET_NAME = "Dashboard";

// Colors
const HEADER_BG = "#1f4e79";
const HEADER_FG = "#ffffff";
const SECTION_BG = "#d6e4f0";
const METRIC_VALUE_FG = "#1f4e79";
const UP_BG = "#FFC7CE";    // Red — price increasing
const DOWN_BG = "#E2EFDA";  // Green — price decreasing
const STABLE_BG = "#F2F2F2";
const ANOMALY_BG = "#FFC7CE";
const WATCH_BG = "#FFEB9C";

/**
 * Write the Dashboard sheet from analysis results.
 *
 * @param {object} analysis - From analyzeHistory()
 */
export async function writeDashboard(analysis) {
  const { metrics, topMovers, watchList } = analysis;

  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;

    // Delete existing dashboard
    try {
      const existing = sheets.getItem(SHEET_NAME);
      existing.delete();
      await ctx.sync();
    } catch {
      // Doesn't exist — fine
    }

    const sheet = sheets.add(SHEET_NAME);
    sheet.activate();
    let row = 0;

    // ── Title ──
    const titleCell = sheet.getRangeByIndexes(row, 0, 1, 7);
    titleCell.merge(true);
    titleCell.values = [["Price Intelligence Dashboard"]];
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 16;
    titleCell.format.font.color = HEADER_BG;
    row += 1;

    const subtitle = sheet.getRangeByIndexes(row, 0, 1, 7);
    subtitle.merge(true);
    subtitle.values = [["Generated: " + new Date().toLocaleDateString("en-GB")]];
    subtitle.format.font.size = 10;
    subtitle.format.font.color = "#797775";
    row += 2;

    // ── Summary Metrics ──
    const metricsTitle = sheet.getRangeByIndexes(row, 0, 1, 7);
    metricsTitle.merge(true);
    metricsTitle.values = [["Summary"]];
    metricsTitle.format.font.bold = true;
    metricsTitle.format.font.size = 13;
    metricsTitle.format.fill.color = SECTION_BG;
    row += 1;

    const metricPairs = [
      ["SKUs Tracked", metrics.totalSKUs],
      ["Reconciliation Runs", metrics.totalRuns],
      ["Total Data Points", metrics.totalRecords],
      ["Avg Price Drift", (metrics.avgDrift >= 0 ? "+" : "") + metrics.avgDrift.toFixed(2) + "%"],
      ["Exception Rate", metrics.exceptionRate.toFixed(1) + "%"],
      ["Anomalies Detected", metrics.anomalyCount],
    ];

    for (const [label, value] of metricPairs) {
      const labelCell = sheet.getRangeByIndexes(row, 0, 1, 2);
      labelCell.merge(true);
      labelCell.values = [[label]];
      labelCell.format.font.bold = true;

      const valueCell = sheet.getRangeByIndexes(row, 2, 1, 2);
      valueCell.merge(true);
      valueCell.values = [[value]];
      valueCell.format.font.color = METRIC_VALUE_FG;
      valueCell.format.font.size = 12;
      valueCell.format.font.bold = true;
      row++;
    }

    row += 1;

    // ── Trend Breakdown ──
    const trendTitle = sheet.getRangeByIndexes(row, 0, 1, 7);
    trendTitle.merge(true);
    trendTitle.values = [["Trend Breakdown"]];
    trendTitle.format.font.bold = true;
    trendTitle.format.font.size = 13;
    trendTitle.format.fill.color = SECTION_BG;
    row += 1;

    const trendRows = [
      ["Trending Up", metrics.trendingUp, UP_BG],
      ["Trending Down", metrics.trendingDown, DOWN_BG],
      ["Stable", metrics.stable, STABLE_BG],
      ["Insufficient Data", metrics.insufficientData, "#F2F2F2"],
    ];

    for (const [label, count, bg] of trendRows) {
      const labelCell = sheet.getRangeByIndexes(row, 0, 1, 2);
      labelCell.merge(true);
      labelCell.values = [[label]];

      const countCell = sheet.getRangeByIndexes(row, 2, 1, 1);
      countCell.values = [[count]];
      countCell.format.font.bold = true;

      // Color bar (visual indicator)
      if (count > 0) {
        const barWidth = Math.min(4, Math.max(1, Math.round((count / metrics.totalSKUs) * 4)));
        const barRange = sheet.getRangeByIndexes(row, 3, 1, barWidth);
        barRange.format.fill.color = bg;
      }
      row++;
    }

    row += 1;

    // ── Top Movers ──
    if (topMovers.length > 0) {
      const moversTitle = sheet.getRangeByIndexes(row, 0, 1, 7);
      moversTitle.merge(true);
      moversTitle.values = [["Top Movers — Biggest Recent Price Changes"]];
      moversTitle.format.font.bold = true;
      moversTitle.format.font.size = 13;
      moversTitle.format.fill.color = SECTION_BG;
      row += 1;

      // Header row
      const moversHeaders = ["SKU", "Product", "Change", "Change %", "Current Price", "Trend", "Data Points"];
      const mhRange = sheet.getRangeByIndexes(row, 0, 1, moversHeaders.length);
      mhRange.values = [moversHeaders];
      mhRange.format.font.bold = true;
      mhRange.format.font.color = HEADER_FG;
      mhRange.format.fill.color = HEADER_BG;
      row += 1;

      for (const m of topMovers) {
        const sign = m.lastChange >= 0 ? "+" : "";
        const rowData = [
          m.sku,
          m.name,
          sign + "£" + m.lastChange.toFixed(2),
          sign + m.lastChangePct.toFixed(1) + "%",
          "£" + m.currentPrice.toFixed(2),
          m.direction,
          m.dataPoints,
        ];
        const dataRange = sheet.getRangeByIndexes(row, 0, 1, rowData.length);
        dataRange.values = [rowData];

        // Color-code based on direction
        const bg = m.lastChange > 0 ? UP_BG : m.lastChange < 0 ? DOWN_BG : STABLE_BG;
        dataRange.format.fill.color = bg;
        row++;
      }

      row += 1;
    }

    // ── Watch List ──
    if (watchList.length > 0) {
      const watchTitle = sheet.getRangeByIndexes(row, 0, 1, 7);
      watchTitle.merge(true);
      watchTitle.values = [["Watch List — SKUs Needing Attention"]];
      watchTitle.format.font.bold = true;
      watchTitle.format.font.size = 13;
      watchTitle.format.fill.color = SECTION_BG;
      row += 1;

      const watchHeaders = ["SKU", "Product", "Flags", "Risk Score", "Current Price", "Trend"];
      const whRange = sheet.getRangeByIndexes(row, 0, 1, watchHeaders.length);
      whRange.values = [watchHeaders];
      whRange.format.font.bold = true;
      whRange.format.font.color = HEADER_FG;
      whRange.format.fill.color = HEADER_BG;
      row += 1;

      for (const w of watchList) {
        const rowData = [
          w.sku,
          w.name,
          w.flags.join("; "),
          w.riskScore,
          w.currentErpPrice != null ? "£" + w.currentErpPrice.toFixed(2) : "",
          w.trend,
        ];
        const dataRange = sheet.getRangeByIndexes(row, 0, 1, rowData.length);
        dataRange.values = [rowData];

        // Color-code by risk score
        if (w.riskScore >= 50) {
          dataRange.format.fill.color = ANOMALY_BG;
        } else {
          dataRange.format.fill.color = WATCH_BG;
        }
        row++;
      }

      row += 1;
    }

    // ── No data message ──
    if (topMovers.length === 0 && watchList.length === 0) {
      const noData = sheet.getRangeByIndexes(row, 0, 1, 7);
      noData.merge(true);
      noData.values = [["Not enough historical data yet. Run more reconciliations to see trends and predictions."]];
      noData.format.font.italic = true;
      noData.format.font.color = "#797775";
      row += 2;
    }

    // ── Legend ──
    const legendTitle = sheet.getRangeByIndexes(row, 0, 1, 7);
    legendTitle.merge(true);
    legendTitle.values = [["Legend"]];
    legendTitle.format.font.bold = true;
    legendTitle.format.font.size = 11;
    row += 1;

    const legends = [
      [UP_BG, "Price increasing — review supplier terms"],
      [DOWN_BG, "Price decreasing — favorable trend"],
      [WATCH_BG, "Watch — approaching risk threshold"],
      [ANOMALY_BG, "Anomaly — price outside expected range"],
    ];

    for (const [color, desc] of legends) {
      const swatch = sheet.getRangeByIndexes(row, 0, 1, 1);
      swatch.format.fill.color = color;

      const descCell = sheet.getRangeByIndexes(row, 1, 1, 6);
      descCell.merge(true);
      descCell.values = [[desc]];
      descCell.format.font.size = 10;
      descCell.format.font.color = "#605e5c";
      row++;
    }

    // Column widths
    const colWidths = [110, 200, 140, 90, 100, 80, 85];
    for (let i = 0; i < colWidths.length; i++) {
      sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = colWidths[i];
    }

    await ctx.sync();
  });
}
