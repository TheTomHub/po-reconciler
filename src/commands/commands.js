/* global Office, Excel */

import { reconcile } from "../reconcile/reconcile";
import { detectColumns } from "../reconcile/detector";
import { writeResultsSheet } from "../reconcile/results";
import { generateCreditNote, generateCorrectedInvoice } from "../reconcile/creditnote";
import { writeCreditNoteSheet, writeReInvoiceSheet } from "../reconcile/creditnote-results";
import { generateEmailDraft } from "../email/email";
import { formatCurrency, setCurrency, parseNumber } from "../utils/format";
import { detectAllColumns, extractPOData } from "../capture/extractor";
import { writeStagingSheet } from "../capture/staging";
import { validate, formatValidationReport } from "../validate/validator";
import { generateStagingEntry } from "../entry/entry";
import { writeEntrySheet } from "../entry/entry-results";
import { toHistoryRecords, analyzeHistory, formatPredictReport } from "../predict/predict";
import { appendHistory, readHistory } from "../predict/history";
import { writeDashboard } from "../predict/dashboard";
import { checkLicense, hasFeature, getLineLimit, getUpgradeMessage } from "../license/license";

/**
 * Shared state for agent actions within a session.
 * Persists across multiple agent invocations in the same runtime.
 */
const agentState = {
  results: null,
  poFilename: "",
};

Office.onReady(() => {
  // Warm the license cache non-blocking so first Copilot action doesn't wait.
  checkLicense().catch(() => {});
  // Agent actions are registered below via Office.actions.associate
});

// ── ReconcilePO ──
// Reads PO data from a sheet named "PO" (or first sheet) and ERP data from:
//   1. A sheet named "ERPPrices" (written by the Copilot agent from SharePoint), or
//   2. The user's current selection (manual fallback).

async function handleReconcilePO(message) {
  const params = message ? JSON.parse(message) : {};
  const tolerance = params.tolerance ?? 0.02;
  const currency = params.currency ?? "GBP";

  setCurrency(currency);

  await checkLicense();

  // Read ERP data — prefer ERPPrices sheet (from SharePoint), fall back to selected range
  let erpData, erpColumns, erpSource;
  await Excel.run(async (context) => {
    let values;

    // Check for ERPPrices sheet written by the agent from SharePoint
    const erpSheet = context.workbook.worksheets.getItemOrNullObject("ERPPrices");
    await context.sync();

    if (!erpSheet.isNullObject) {
      const usedRange = erpSheet.getUsedRange();
      usedRange.load("values");
      await context.sync();
      values = usedRange.values;
      erpSource = "SharePoint price list (ERPPrices sheet)";
    } else {
      // Fall back to user's selected range
      const selection = context.workbook.getSelectedRange();
      selection.load("values");
      await context.sync();
      values = selection.values;
      erpSource = "selected range";
    }

    if (!values || values.length < 2) {
      throw new Error(
        erpSource === "selected range"
          ? "No ERP data found. Either ask the agent to load the price list from SharePoint, or select your ERP data range in Excel first."
          : "ERPPrices sheet has no data. Please ensure the price list was written correctly from SharePoint."
      );
    }

    const headers = values[0].map((h) => String(h).trim()).filter(Boolean);
    const rows = values.slice(1)
      .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""))
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] != null ? String(row[i]) : ""; });
        return obj;
      });

    erpColumns = detectColumns(headers);
    if (!erpColumns.sku || !erpColumns.price) {
      throw new Error(`Could not detect SKU and Price columns in ERP data. Found headers: ${headers.join(", ")}`);
    }

    erpData = { headers, rows };
  });

  // Read PO data from sheet named "PO" or first sheet
  let poData, poColumns;
  await Excel.run(async (context) => {
    let sheet;
    const poSheet = context.workbook.worksheets.getItemOrNullObject("PO");
    await context.sync();

    if (!poSheet.isNullObject) {
      sheet = poSheet;
    } else {
      sheet = context.workbook.worksheets.getFirst();
    }

    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const values = usedRange.values;
    if (!values || values.length < 2) {
      throw new Error("PO sheet has no data. Upload a PO file first.");
    }

    const headers = values[0].map((h) => String(h).trim()).filter(Boolean);
    const rows = values.slice(1)
      .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""))
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] != null ? String(row[i]) : ""; });
        return obj;
      });

    poColumns = detectColumns(headers);
    if (!poColumns.sku || !poColumns.price) {
      throw new Error(`Could not detect SKU and Price columns in PO data. Found headers: ${headers.join(", ")}`);
    }

    poData = { headers, rows };
    agentState.poFilename = sheet.name;
  });

  // Enforce line limit for free tier
  const lineLimit = getLineLimit();
  if (lineLimit > 0 && poData.rows.length > lineLimit) {
    return `Free plan limit: ReconcilePO supports up to ${lineLimit} lines. This PO has ${poData.rows.length} lines.\n\n${getUpgradeMessage("ReconcilePO")}`;
  }

  // Run reconciliation
  const results = reconcile({ poData, poColumns, erpData, erpColumns, tolerance });
  agentState.results = results;

  // Write results sheet
  await writeResultsSheet(results, tolerance, agentState.poFilename);

  // Append to price history
  try {
    const today = new Date().toISOString().slice(0, 10);
    const historyRecords = toHistoryRecords(results, agentState.poFilename, today);
    await appendHistory(historyRecords);
  } catch {
    // Non-critical — don't fail the reconciliation
  }

  // Format summary
  const s = results.summary;
  return `Reconciliation complete.\n\nERP data: ${erpSource}\nTotal items: ${s.total}\nPerfect matches: ${s.matches}\nWithin tolerance: ${s.tolerances}\nExceptions: ${s.exceptions}\nWarnings: ${s.warnings}\nTotal exposure: ${formatCurrency(s.exposure)}\n\nResults sheet created with color-coded status rows.`;
}

// ── GenerateCreditNote ──

async function handleGenerateCreditNote() {
  if (!agentState.results) {
    return "No reconciliation results available. Please run ReconcilePO first.";
  }

  const creditData = generateCreditNote(agentState.results);
  await writeCreditNoteSheet(creditData, agentState.poFilename);

  const t = creditData.totals;
  return `Credit note created.\n\nLines: ${t.lineCount}\nTotal credit: ${formatCurrency(t.totalCredit)}\n\nThe CreditNote sheet has been activated.`;
}

// ── GenerateReInvoice ──

async function handleGenerateReInvoice() {
  if (!agentState.results) {
    return "No reconciliation results available. Please run ReconcilePO first.";
  }

  const invoiceData = generateCorrectedInvoice(agentState.results);
  const exceptionCount = agentState.results.summary.exceptions;
  await writeReInvoiceSheet(invoiceData, agentState.poFilename, exceptionCount);

  const t = invoiceData.totals;
  return `Re-invoice created.\n\nLines: ${t.lineCount}\nCorrected total: ${formatCurrency(t.totalInvoice)}\nPrice corrections: ${exceptionCount}\n\nThe ReInvoice sheet has been activated. Yellow-highlighted rows indicate price corrections.`;
}

// ── DraftExceptionEmail ──

async function handleDraftExceptionEmail() {
  if (!agentState.results) {
    return "No reconciliation results available. Please run ReconcilePO first.";
  }

  const draft = generateEmailDraft(agentState.results, agentState.poFilename);
  return `Subject: ${draft.subject}\n\n${draft.body}`;
}

// ── ExtractPOData ──
// Reads PO data from the active sheet (or a named sheet), extracts structured
// data using extended column detection, and writes a staging sheet.

async function handleExtractPOData(message) {
  const params = message ? JSON.parse(message) : {};
  const sheetNameParam = params.sheetName || null;
  const currency = params.currency ?? "GBP";

  setCurrency(currency);

  await checkLicense();

  // Read PO data from specified sheet or active sheet
  let parsedData;
  await Excel.run(async (context) => {
    let sheet;
    if (sheetNameParam) {
      sheet = context.workbook.worksheets.getItemOrNullObject(sheetNameParam);
      await context.sync();
      if (sheet.isNullObject) {
        throw new Error(`Sheet "${sheetNameParam}" not found. Available sheets can be listed.`);
      }
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }

    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    sheet.load("name");
    await context.sync();

    const values = usedRange.values;
    if (!values || values.length < 2) {
      throw new Error("Sheet has no data. Open or paste PO data first.");
    }

    // Find the best header row (row with most recognized column keywords)
    let bestRow = 0;
    let bestScore = 0;
    const keywords = ["sku", "item", "product", "price", "cost", "qty", "quantity", "description", "name", "unit", "total", "amount", "date", "order"];

    for (let r = 0; r < Math.min(values.length, 15); r++) {
      let score = 0;
      for (const cell of values[r]) {
        const val = String(cell || "").toLowerCase().trim();
        if (val.length > 40) continue;
        for (const kw of keywords) {
          if (val.includes(kw)) { score++; break; }
        }
      }
      if (score > bestScore) { bestScore = score; bestRow = r; }
    }

    const headers = values[bestRow].map((h) => String(h ?? "").trim()).filter(Boolean);
    const rows = values.slice(bestRow + 1)
      .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""))
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] != null ? String(row[i]) : ""; });
        return obj;
      });

    parsedData = { headers, rows };
    agentState.poFilename = sheet.name;
  });

  // Enforce line limit for free tier
  const lineLimit = getLineLimit();
  if (lineLimit > 0 && parsedData.rows.length > lineLimit) {
    return `Free plan limit: ExtractPOData supports up to ${lineLimit} lines. This PO has ${parsedData.rows.length} lines.\n\n${getUpgradeMessage("ExtractPOData")}`;
  }

  // Detect columns and extract
  const columns = detectAllColumns(parsedData.headers);
  const extraction = extractPOData(parsedData, columns);

  // Use sheet name as fallback when no PO ref column was detected
  if (extraction.metadata.poRef === "Unknown") {
    extraction.metadata.poRef = agentState.poFilename || "PO";
  }
  if (extraction.metadata.customer === "Unknown") {
    extraction.metadata.customer = "";
  }

  // Write staging sheet
  await writeStagingSheet(extraction);

  // Run validation
  const validation = validate(extraction.stagingRows);

  // Store for potential follow-up reconciliation
  agentState.lastExtraction = extraction;
  agentState.lastValidation = validation;

  // Format response
  const m = extraction.metadata;
  const detectedStr = m.detectedFields.join(", ");
  let response = `PO data extracted and staging sheet created.\n\n`;
  response += `PO Reference: ${m.poRef}\n`;
  response += `Customer: ${m.customer}\n`;
  response += `Line items: ${m.lineCount}\n`;
  response += `Total value: ${formatCurrency(m.totalValue)}\n`;
  response += `Detected fields: ${detectedStr}\n`;

  // Validation results
  response += `\n--- Validation ---\n`;
  response += formatValidationReport(validation);

  if (m.warningCount > 0) {
    response += `\n\nExtraction warnings: ${m.warningCount}\n`;
    const topWarnings = extraction.warnings.slice(0, 5);
    for (const w of topWarnings) {
      response += `  Line ${w.line}: ${w.field} — ${w.message}\n`;
    }
    if (extraction.warnings.length > 5) {
      response += `  ... and ${extraction.warnings.length - 5} more (see Warnings sheet)\n`;
    }
    response += `\nYellow-highlighted rows need review. Check cell notes for details.`;
  }

  if (!validation.valid) {
    response += `\n\nValidation errors must be resolved before reconciliation.`;
  } else {
    response += `\n\nData is ready for reconciliation.`;
  }

  return response;
}

// ── GenerateERPStaging ──

async function handleGenerateERPStaging(message) {
  await checkLicense();
  if (!hasFeature("GenerateERPStaging")) {
    return getUpgradeMessage("GenerateERPStaging");
  }

  if (!agentState.results) {
    return "No reconciliation results available. Please run ReconcilePO first.";
  }

  const params = message ? JSON.parse(message) : {};

  const entryData = generateStagingEntry(agentState.results, {
    poRef: agentState.poFilename,
    customer: params.customer || "",
    deliveryDate: params.deliveryDate || "",
  });

  await writeEntrySheet(entryData);

  const t = entryData.totals;
  const m = entryData.metadata;
  let response = `ERP staging sheet created.\n\n`;
  response += `PO Reference: ${m.poRef}\n`;
  response += `Total lines: ${t.lineCount}\n`;
  response += `Total value: ${formatCurrency(t.totalValue)}\n\n`;
  response += `Status breakdown:\n`;
  response += `  Ready (green):  ${t.readyCount} — safe to enter into ERP\n`;
  response += `  Review (yellow): ${t.reviewCount} — price exceptions, operator decision needed\n`;
  response += `  Hold (red):     ${t.holdCount} — missing from ERP or data issue\n\n`;
  response += `The ERPStaging sheet has been activated. Green rows can be entered directly. Yellow rows need price confirmation. Red rows need item verification.`;

  return response;
}

// ── GenerateDashboard ──

async function handleGenerateDashboard() {
  await checkLicense();
  if (!hasFeature("GenerateDashboard")) {
    return getUpgradeMessage("GenerateDashboard");
  }

  const records = await readHistory();

  if (records.length === 0) {
    return "No price history available yet. Run at least one reconciliation to start building history. Each reconciliation automatically saves price data for trend analysis.";
  }

  const analysis = analyzeHistory(records);
  await writeDashboard(analysis);

  const m = analysis.metrics;
  let response = `Price Intelligence Dashboard created.\n\n`;
  response += `Data: ${m.totalRecords} records across ${m.totalRuns} reconciliation runs, ${m.totalSKUs} unique SKUs.\n\n`;
  response += `Key metrics:\n`;
  response += `  Average price drift: ${m.avgDrift >= 0 ? "+" : ""}${m.avgDrift.toFixed(2)}%\n`;
  response += `  Exception rate: ${m.exceptionRate.toFixed(1)}%\n`;
  response += `  Anomalies detected: ${m.anomalyCount}\n\n`;
  response += `Trend breakdown:\n`;
  response += `  Trending up: ${m.trendingUp} SKUs\n`;
  response += `  Trending down: ${m.trendingDown} SKUs\n`;
  response += `  Stable: ${m.stable} SKUs\n`;

  if (analysis.topMovers.length > 0) {
    response += `\nTop movers (biggest recent price changes):\n`;
    for (const mv of analysis.topMovers.slice(0, 5)) {
      const sign = mv.lastChange >= 0 ? "+" : "";
      response += `  ${mv.sku} ${mv.name}: ${sign}£${mv.lastChange.toFixed(2)} (${sign}${mv.lastChangePct.toFixed(1)}%)\n`;
    }
  }

  if (analysis.watchList.length > 0) {
    response += `\nWatch list (SKUs needing attention):\n`;
    for (const w of analysis.watchList.slice(0, 5)) {
      response += `  ${w.sku} ${w.name} — ${w.flags.join(", ")}\n`;
    }
  }

  response += `\nThe Dashboard sheet has been activated with full details.`;
  return response;
}

// ── GetPriceIntelligence ──
// Returns a text-based intelligence report for conversational analysis.
// The agent can discuss trends, anomalies, and risks without writing a sheet.

async function handleGetPriceIntelligence() {
  await checkLicense();
  if (!hasFeature("GetPriceIntelligence")) {
    return getUpgradeMessage("GetPriceIntelligence");
  }

  const records = await readHistory();

  if (records.length === 0) {
    return "No price history available yet. Run at least one reconciliation to start building history.";
  }

  const analysis = analyzeHistory(records);
  return formatPredictReport(analysis);
}

// ── LookupSKU ──
// Looks up a specific SKU's complete price history, trend, and risk flags.

async function handleLookupSKU(message) {
  const params = message ? JSON.parse(message) : {};
  const query = (params.sku || "").trim().toUpperCase();

  if (!query) {
    return "Please provide a SKU to look up.";
  }

  const records = await readHistory();
  if (records.length === 0) {
    return "No price history available yet. Run at least one reconciliation first.";
  }

  // Find matching records (exact or partial match)
  const matches = records.filter((r) => {
    const norm = r.sku.trim().toUpperCase();
    return norm === query || norm.startsWith(query) || query.startsWith(norm);
  });

  if (matches.length === 0) {
    return `SKU "${query}" not found in price history. Available SKUs can be seen in the Dashboard.`;
  }

  // Sort by date ascending
  const sorted = [...matches].sort((a, b) => a.date.localeCompare(b.date));
  const latest = sorted[sorted.length - 1];
  const name = latest.name || sorted.find((r) => r.name)?.name || "";

  // Build price timeline
  const lines = [];
  lines.push(`SKU LOOKUP: ${query}`);
  if (name) lines.push(`Product: ${name}`);
  lines.push(`Data points: ${sorted.length}`);
  lines.push(`Date range: ${sorted[0].date} to ${latest.date}`);
  lines.push("");

  // Price history table
  lines.push("PRICE HISTORY");
  for (const r of sorted) {
    const po = r.poPrice != null ? `PO £${r.poPrice.toFixed(2)}` : "";
    const erp = r.erpPrice != null ? `ERP £${r.erpPrice.toFixed(2)}` : "";
    const diff = r.diff != null ? `Diff £${r.diff.toFixed(2)}` : "";
    lines.push(`  ${r.date}  ${r.poRef}  ${po}  ${erp}  ${diff}  [${r.status}]`);
  }
  lines.push("");

  // Trend analysis
  const erpPrices = sorted.filter((r) => r.erpPrice != null).map((r) => r.erpPrice);
  if (erpPrices.length >= 2) {
    const first = erpPrices[0];
    const last = erpPrices[erpPrices.length - 1];
    const change = last - first;
    const changePct = first !== 0 ? ((change / first) * 100).toFixed(1) : "N/A";
    const direction = change > 0.01 ? "INCREASING" : change < -0.01 ? "DECREASING" : "STABLE";

    lines.push("TREND");
    lines.push(`  Direction: ${direction}`);
    lines.push(`  First ERP price: £${first.toFixed(2)}`);
    lines.push(`  Latest ERP price: £${last.toFixed(2)}`);
    lines.push(`  Total change: ${change >= 0 ? "+" : ""}£${change.toFixed(2)} (${change >= 0 ? "+" : ""}${changePct}%)`);

    // Volatility
    const mean = erpPrices.reduce((a, b) => a + b, 0) / erpPrices.length;
    const variance = erpPrices.reduce((sum, p) => sum + (p - mean) ** 2, 0) / (erpPrices.length - 1);
    const sd = Math.sqrt(variance);
    lines.push(`  Volatility (std dev): £${sd.toFixed(2)}`);
    lines.push("");
  }

  // Exception stats
  const exceptionCount = sorted.filter((r) => r.status === "Exception").length;
  const exceptionRate = ((exceptionCount / sorted.length) * 100).toFixed(0);
  lines.push("RISK FLAGS");
  lines.push(`  Exception rate: ${exceptionRate}% (${exceptionCount}/${sorted.length} runs)`);
  if (exceptionRate >= 50 && sorted.length >= 2) lines.push(`  ⚠ Frequent exceptions — review supplier terms`);
  if (erpPrices.length >= 3) {
    const mean = erpPrices.slice(0, -1).reduce((a, b) => a + b, 0) / (erpPrices.length - 1);
    const sd2 = Math.sqrt(erpPrices.slice(0, -1).reduce((sum, p) => sum + (p - mean) ** 2, 0) / (erpPrices.length - 2));
    if (sd2 > 0) {
      const zScore = Math.abs((erpPrices[erpPrices.length - 1] - mean) / sd2);
      if (zScore > 2) lines.push(`  ⚠ Price anomaly — latest price is ${zScore.toFixed(1)} std devs from historical mean`);
    }
  }

  return lines.join("\n");
}

// ── AssessPORisk ──
// Provides a risk assessment of the current PO based on reconciliation results
// and historical price data. Returns accept/review/escalate recommendation.

async function handleAssessPORisk() {
  await checkLicense();
  if (!hasFeature("AssessPORisk")) {
    return getUpgradeMessage("AssessPORisk");
  }

  if (!agentState.results) {
    return "No reconciliation results available. Please run ReconcilePO first.";
  }

  const results = agentState.results;
  const s = results.summary;

  // Load historical data for context
  let historicalContext = "";
  try {
    const records = await readHistory();
    if (records.length > 0) {
      const analysis = analyzeHistory(records);
      const m = analysis.metrics;

      // Check how current PO compares to history
      const currentExceptionRate = s.total > 0 ? ((s.exceptions / s.total) * 100) : 0;
      const historicalExceptionRate = m.exceptionRate;
      const rateDiff = currentExceptionRate - historicalExceptionRate;

      historicalContext += "\nHISTORICAL CONTEXT\n";
      historicalContext += `  Your historical exception rate: ${historicalExceptionRate.toFixed(1)}%\n`;
      historicalContext += `  This PO's exception rate: ${currentExceptionRate.toFixed(1)}%\n`;
      if (rateDiff > 5) {
        historicalContext += `  ⚠ This PO has ${rateDiff.toFixed(1)}% MORE exceptions than your average\n`;
      } else if (rateDiff < -5) {
        historicalContext += `  ✓ This PO has ${Math.abs(rateDiff).toFixed(1)}% FEWER exceptions than your average\n`;
      } else {
        historicalContext += `  This PO is within normal range\n`;
      }

      // Check which exception SKUs have concerning history
      const exceptionRows = results.rows.filter((r) => r.status === "Exception");
      const concerningSKUs = [];

      for (const row of exceptionRows) {
        const norm = row.sku.trim().toUpperCase();
        const watchItem = analysis.watchList.find((w) => w.sku === norm);
        if (watchItem) {
          concerningSKUs.push({ sku: row.sku, flags: watchItem.flags, risk: watchItem.riskScore });
        }
      }

      if (concerningSKUs.length > 0) {
        historicalContext += "\n  SKUs with historical concerns:\n";
        for (const c of concerningSKUs.slice(0, 5)) {
          historicalContext += `    ${c.sku} (risk: ${c.risk}/100) — ${c.flags.join(", ")}\n`;
        }
      }

      // Check for any anomalous prices in this PO
      const anomalousInPO = [];
      for (const row of exceptionRows) {
        const norm = row.sku.trim().toUpperCase();
        const skuData = analysis.skuAnalysis.get(norm);
        if (skuData && skuData.anomaly) {
          anomalousInPO.push(row.sku);
        }
      }

      if (anomalousInPO.length > 0) {
        historicalContext += `\n  Price anomalies detected in this PO: ${anomalousInPO.join(", ")}\n`;
      }
    }
  } catch {
    // History unavailable — proceed without
  }

  // Build risk assessment
  const lines = [];
  lines.push("PO RISK ASSESSMENT");
  lines.push("=".repeat(40));
  lines.push("");

  // Overview
  lines.push("OVERVIEW");
  lines.push(`  PO Reference: ${agentState.poFilename}`);
  lines.push(`  Total items: ${s.total}`);
  lines.push(`  Matches: ${s.matches} (${s.total > 0 ? ((s.matches / s.total) * 100).toFixed(0) : 0}%)`);
  lines.push(`  Exceptions: ${s.exceptions} (${s.total > 0 ? ((s.exceptions / s.total) * 100).toFixed(0) : 0}%)`);
  lines.push(`  Tolerances: ${s.tolerances}`);
  lines.push(`  Warnings: ${s.warnings}`);
  lines.push(`  Total exposure: ${formatCurrency(s.exposure)}`);
  lines.push("");

  // Biggest risks
  const exceptionRows = results.rows
    .filter((r) => r.status === "Exception" && r.diff != null)
    .sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));

  if (exceptionRows.length > 0) {
    lines.push("TOP PRICE DISCREPANCIES");
    for (const row of exceptionRows.slice(0, 8)) {
      const sign = row.diff >= 0 ? "+" : "";
      const qty = row.poQty || 1;
      const lineImpact = Math.abs(row.diff * qty);
      lines.push(`  ${row.sku} ${row.name || ""}`);
      lines.push(`    PO: ${formatCurrency(row.poPrice)} → ERP: ${formatCurrency(row.erpPrice)}  (${sign}${formatCurrency(row.diff)} per unit, ${formatCurrency(lineImpact)} total)`);
    }
    lines.push("");
  }

  // Items not in ERP
  const notInErp = results.rows.filter((r) => r.status === "Not in ERP");
  if (notInErp.length > 0) {
    lines.push(`ITEMS NOT FOUND IN ERP: ${notInErp.length}`);
    for (const row of notInErp.slice(0, 5)) {
      lines.push(`  ${row.sku} ${row.name || ""} — ${formatCurrency(row.poPrice)}`);
    }
    if (notInErp.length > 5) lines.push(`  ... and ${notInErp.length - 5} more`);
    lines.push("");
  }

  // Historical context
  if (historicalContext) {
    lines.push(historicalContext);
  }

  // Recommendation
  lines.push("");
  lines.push("RECOMMENDATION");

  const exceptionPct = s.total > 0 ? (s.exceptions / s.total) * 100 : 0;
  const hasHighExposure = s.exposure > 100;
  const hasNotInErp = notInErp.length > 0;

  if (exceptionPct === 0 && !hasNotInErp) {
    lines.push("  ✓ ACCEPT — All prices match or are within tolerance.");
    lines.push("  Action: Generate ERP staging sheet and process the order.");
  } else if (exceptionPct <= 10 && s.exposure < 50 && notInErp.length === 0) {
    lines.push("  ✓ ACCEPT WITH REVIEW — Minor price discrepancies.");
    lines.push(`  Action: Review ${s.exceptions} exception(s), generate credit note if needed, then process.`);
  } else if (exceptionPct <= 30 && !hasHighExposure) {
    lines.push("  ⚠ REVIEW — Moderate price discrepancies found.");
    lines.push(`  Action: Generate credit note and re-invoice for ${s.exceptions} exception(s). Contact supplier if patterns persist.`);
  } else {
    lines.push("  🛑 ESCALATE — Significant pricing issues.");
    lines.push(`  Action: ${s.exceptions} exceptions with ${formatCurrency(s.exposure)} exposure. Escalate to pricing team before processing.`);
    if (hasNotInErp) lines.push(`  ${notInErp.length} item(s) not found in ERP — may be new products or incorrect SKUs.`);
  }

  return lines.join("\n");
}

// ── Register agent actions ──

Office.actions.associate("ExtractPOData", async (message) => {
  try {
    return await handleExtractPOData(message);
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("ReconcilePO", async (message) => {
  try {
    return await handleReconcilePO(message);
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("GenerateCreditNote", async () => {
  try {
    return await handleGenerateCreditNote();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("GenerateReInvoice", async () => {
  try {
    return await handleGenerateReInvoice();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("DraftExceptionEmail", async () => {
  try {
    return await handleDraftExceptionEmail();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("GenerateERPStaging", async (message) => {
  try {
    return await handleGenerateERPStaging(message);
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("GenerateDashboard", async () => {
  try {
    return await handleGenerateDashboard();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("GetPriceIntelligence", async () => {
  try {
    return await handleGetPriceIntelligence();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("LookupSKU", async (message) => {
  try {
    return await handleLookupSKU(message);
  } catch (err) {
    return `Error: ${err.message}`;
  }
});

Office.actions.associate("AssessPORisk", async () => {
  try {
    return await handleAssessPORisk();
  } catch (err) {
    return `Error: ${err.message}`;
  }
});
