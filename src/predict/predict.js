/**
 * Predict module — pure JS analytics engine.
 *
 * Analyzes historical price records across reconciliation runs to detect
 * trends, anomalies, and leading indicators of pricing risk.
 *
 * No Office.js dependency — testable from Node.js.
 */

/**
 * Convert reconciliation results into history records for storage.
 *
 * @param {object} reconciliation - { summary, rows } from reconcile()
 * @param {string} poRef - PO reference number
 * @param {string} date - ISO date string (YYYY-MM-DD)
 * @returns {object[]} Records to append to history
 */
export function toHistoryRecords(reconciliation, poRef, date) {
  return reconciliation.rows
    .filter((r) => r.poPrice != null && r.status !== "Not in PO")
    .map((r) => ({
      date,
      poRef,
      sku: r.sku,
      name: r.name || "",
      poPrice: r.poPrice,
      erpPrice: r.erpPrice,
      diff: r.diff,
      status: r.status,
      poQty: r.poQty || 0,
    }));
}

/**
 * Analyze historical price records and produce predictions/insights.
 *
 * @param {object[]} records - All historical records (from PriceHistory sheet)
 * @returns {object} { skuAnalysis, topMovers, watchList, metrics }
 */
export function analyzeHistory(records) {
  if (!records || records.length === 0) {
    return emptyAnalysis();
  }

  // Group records by SKU
  const bySku = new Map();
  for (const r of records) {
    const key = normSku(r.sku);
    if (!bySku.has(key)) bySku.set(key, []);
    bySku.get(key).push(r);
  }

  // Analyze each SKU
  const skuAnalysis = new Map();
  for (const [sku, skuRecords] of bySku) {
    // Sort by date ascending
    const sorted = [...skuRecords].sort((a, b) => a.date.localeCompare(b.date));
    skuAnalysis.set(sku, analyzeSku(sorted));
  }

  // Top movers: SKUs with the biggest absolute price change in the most recent record
  const topMovers = [...skuAnalysis.entries()]
    .filter(([, a]) => a.lastChange !== null && a.dataPoints >= 2)
    .sort((a, b) => Math.abs(b[1].lastChange) - Math.abs(a[1].lastChange))
    .slice(0, 10)
    .map(([sku, a]) => ({
      sku,
      name: a.name,
      lastChange: a.lastChange,
      lastChangePct: a.lastChangePct,
      direction: a.trend.direction,
      currentPrice: a.currentErpPrice,
      dataPoints: a.dataPoints,
    }));

  // Watch list: SKUs with concerning patterns
  const watchList = [...skuAnalysis.entries()]
    .filter(([, a]) => a.flags.length > 0)
    .sort((a, b) => b[1].flags.length - a[1].flags.length || b[1].riskScore - a[1].riskScore)
    .slice(0, 15)
    .map(([sku, a]) => ({
      sku,
      name: a.name,
      flags: a.flags,
      riskScore: a.riskScore,
      currentErpPrice: a.currentErpPrice,
      trend: a.trend.direction,
    }));

  // Aggregate metrics
  const metrics = computeMetrics(records, skuAnalysis);

  return { skuAnalysis, topMovers, watchList, metrics };
}

/**
 * Format analysis results as a human-readable report (for Copilot agent).
 */
export function formatPredictReport(analysis) {
  const { metrics, topMovers, watchList } = analysis;

  const lines = [];
  lines.push("PRICE INTELLIGENCE REPORT");
  lines.push("=".repeat(40));
  lines.push("");

  lines.push("SUMMARY");
  lines.push(`  SKUs tracked: ${metrics.totalSKUs}`);
  lines.push(`  Data points: ${metrics.totalRecords}`);
  lines.push(`  PO runs analyzed: ${metrics.totalRuns}`);
  lines.push(`  Avg ERP price drift: ${metrics.avgDrift >= 0 ? "+" : ""}${metrics.avgDrift.toFixed(2)}%`);
  lines.push(`  Exception rate: ${metrics.exceptionRate.toFixed(1)}%`);
  lines.push(`  Anomalies detected: ${metrics.anomalyCount}`);
  lines.push("");

  lines.push("TREND BREAKDOWN");
  lines.push(`  Trending up: ${metrics.trendingUp} SKUs`);
  lines.push(`  Trending down: ${metrics.trendingDown} SKUs`);
  lines.push(`  Stable: ${metrics.stable} SKUs`);
  lines.push(`  Insufficient data: ${metrics.insufficientData} SKUs`);
  lines.push("");

  if (topMovers.length > 0) {
    lines.push("TOP MOVERS (biggest recent price changes)");
    for (const m of topMovers) {
      const sign = m.lastChange >= 0 ? "+" : "";
      lines.push(`  ${m.sku} ${m.name}`);
      lines.push(`    Change: ${sign}£${m.lastChange.toFixed(2)} (${sign}${m.lastChangePct.toFixed(1)}%) → £${m.currentPrice.toFixed(2)}`);
    }
    lines.push("");
  }

  if (watchList.length > 0) {
    lines.push("WATCH LIST (SKUs needing attention)");
    for (const w of watchList) {
      lines.push(`  ${w.sku} ${w.name} — ${w.flags.join(", ")}`);
    }
    lines.push("");
  }

  if (topMovers.length === 0 && watchList.length === 0) {
    lines.push("No significant price movements or anomalies detected.");
    lines.push("");
  }

  return lines.join("\n");
}

// ── Internal helpers ──

function emptyAnalysis() {
  return {
    skuAnalysis: new Map(),
    topMovers: [],
    watchList: [],
    metrics: {
      totalSKUs: 0,
      totalRecords: 0,
      totalRuns: 0,
      avgDrift: 0,
      exceptionRate: 0,
      anomalyCount: 0,
      trendingUp: 0,
      trendingDown: 0,
      stable: 0,
      insufficientData: 0,
    },
  };
}

/**
 * Analyze a single SKU's price history.
 */
function analyzeSku(sortedRecords) {
  const latest = sortedRecords[sortedRecords.length - 1];
  const name = latest.name || sortedRecords.find((r) => r.name)?.name || "";

  // Extract ERP price series (use ERP price when available, fall back to PO price)
  const priceSeries = sortedRecords
    .filter((r) => r.erpPrice != null)
    .map((r) => ({ date: r.date, price: r.erpPrice }));

  const currentErpPrice = priceSeries.length > 0
    ? priceSeries[priceSeries.length - 1].price
    : latest.poPrice;

  // Trend analysis (needs at least 2 data points)
  const trend = priceSeries.length >= 2
    ? computeTrend(priceSeries)
    : { direction: "insufficient", slope: 0, confidence: 0 };

  // Last price change
  let lastChange = null;
  let lastChangePct = null;
  if (priceSeries.length >= 2) {
    const prev = priceSeries[priceSeries.length - 2].price;
    const curr = priceSeries[priceSeries.length - 1].price;
    lastChange = round(curr - prev);
    lastChangePct = prev !== 0 ? round(((curr - prev) / prev) * 100) : 0;
  }

  // Volatility (standard deviation of prices)
  const prices = priceSeries.map((p) => p.price);
  const volatility = prices.length >= 2 ? stdDev(prices) : 0;

  // Anomaly detection (z-score of latest price)
  let anomaly = false;
  if (prices.length >= 3) {
    const mean = avg(prices.slice(0, -1)); // mean of all except latest
    const sd = stdDev(prices.slice(0, -1));
    if (sd > 0) {
      const zScore = Math.abs((prices[prices.length - 1] - mean) / sd);
      anomaly = zScore > 2;
    }
  }

  // Exception frequency
  const exceptionCount = sortedRecords.filter((r) => r.status === "Exception").length;
  const exceptionRate = sortedRecords.length > 0
    ? round((exceptionCount / sortedRecords.length) * 100)
    : 0;

  // Flags
  const flags = [];
  if (anomaly) flags.push("Price anomaly");
  if (trend.direction === "up" && trend.confidence > 0.5) flags.push("Consistent price increase");
  if (exceptionRate >= 50 && sortedRecords.length >= 2) flags.push("Frequent exceptions");
  if (lastChangePct !== null && Math.abs(lastChangePct) > 10) flags.push("Large recent change (>" + Math.abs(lastChangePct).toFixed(0) + "%)");
  if (volatility > 0 && currentErpPrice > 0 && (volatility / currentErpPrice) > 0.1) flags.push("High volatility");

  // Risk score (0-100)
  let riskScore = 0;
  if (anomaly) riskScore += 30;
  if (trend.direction === "up") riskScore += 15;
  if (exceptionRate >= 50) riskScore += 20;
  if (lastChangePct !== null && Math.abs(lastChangePct) > 5) riskScore += 15;
  if (volatility > 0 && currentErpPrice > 0) riskScore += Math.min(20, (volatility / currentErpPrice) * 200);
  riskScore = Math.min(100, Math.round(riskScore));

  return {
    name,
    dataPoints: priceSeries.length,
    currentErpPrice,
    trend,
    lastChange,
    lastChangePct,
    volatility: round(volatility),
    anomaly,
    exceptionRate,
    exceptionCount,
    flags,
    riskScore,
  };
}

/**
 * Linear regression on price over time.
 * Returns { direction: "up"|"down"|"stable", slope, confidence (R²) }
 */
function computeTrend(priceSeries) {
  const n = priceSeries.length;
  if (n < 2) return { direction: "insufficient", slope: 0, confidence: 0 };

  // Convert dates to numeric indices (0, 1, 2, ...)
  const xs = priceSeries.map((_, i) => i);
  const ys = priceSeries.map((p) => p.price);

  const xMean = avg(xs);
  const yMean = avg(ys);

  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - xMean) * (ys[i] - yMean);
    den += (xs[i] - xMean) ** 2;
  }

  const slope = den !== 0 ? num / den : 0;

  // R² (coefficient of determination)
  const yPred = xs.map((x) => yMean + slope * (x - xMean));
  const ssRes = ys.reduce((sum, y, i) => sum + (y - yPred[i]) ** 2, 0);
  const ssTot = ys.reduce((sum, y) => sum + (y - yMean) ** 2, 0);
  const r2 = ssTot > 0 ? 1 - ssRes / ssTot : 0;

  // Classify direction (slope relative to mean price)
  const slopePct = yMean !== 0 ? (slope / yMean) * 100 : 0;
  let direction;
  if (Math.abs(slopePct) < 0.5) {
    direction = "stable";
  } else {
    direction = slope > 0 ? "up" : "down";
  }

  return {
    direction,
    slope: round(slope),
    slopePct: round(slopePct),
    confidence: round(r2),
  };
}

/**
 * Aggregate metrics across all SKUs.
 */
function computeMetrics(records, skuAnalysis) {
  const totalSKUs = skuAnalysis.size;
  const totalRecords = records.length;
  const uniqueRuns = new Set(records.map((r) => r.poRef + "|" + r.date)).size;

  // Drift: average of per-SKU trend slopes (as % of mean price)
  let driftSum = 0;
  let driftCount = 0;
  let trendingUp = 0;
  let trendingDown = 0;
  let stable = 0;
  let insufficientData = 0;
  let anomalyCount = 0;

  for (const [, a] of skuAnalysis) {
    if (a.trend.direction === "up") {
      trendingUp++;
      driftSum += a.trend.slopePct || 0;
      driftCount++;
    } else if (a.trend.direction === "down") {
      trendingDown++;
      driftSum += a.trend.slopePct || 0;
      driftCount++;
    } else if (a.trend.direction === "stable") {
      stable++;
      driftCount++;
    } else {
      insufficientData++;
    }
    if (a.anomaly) anomalyCount++;
  }

  const avgDrift = driftCount > 0 ? round(driftSum / driftCount) : 0;

  // Exception rate across all records
  const exceptionRecords = records.filter((r) => r.status === "Exception").length;
  const exceptionRate = totalRecords > 0 ? round((exceptionRecords / totalRecords) * 100) : 0;

  return {
    totalSKUs,
    totalRecords,
    totalRuns: uniqueRuns,
    avgDrift,
    exceptionRate,
    anomalyCount,
    trendingUp,
    trendingDown,
    stable,
    insufficientData,
  };
}

// ── Utility functions ──

function normSku(sku) {
  return String(sku).trim().toUpperCase();
}

function avg(arr) {
  return arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
}

function stdDev(arr) {
  if (arr.length < 2) return 0;
  const mean = avg(arr);
  const variance = arr.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (arr.length - 1);
  return Math.sqrt(variance);
}

function round(n) {
  return Math.round(n * 100) / 100;
}
