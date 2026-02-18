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

/**
 * Shared state for agent actions within a session.
 * Persists across multiple agent invocations in the same runtime.
 */
const agentState = {
  results: null,
  poFilename: "",
};

Office.onReady(() => {
  // Agent actions are registered below via Office.actions.associate
});

// ── ReconcilePO ──
// Reads PO data from a sheet named "PO" (or first sheet) and ERP data from the
// user's current selection, then runs reconciliation and writes results sheet.

async function handleReconcilePO(message) {
  const params = message ? JSON.parse(message) : {};
  const tolerance = params.tolerance ?? 0.02;
  const currency = params.currency ?? "GBP";

  setCurrency(currency);

  // Read ERP data from current selection
  let erpData, erpColumns;
  await Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange();
    selection.load("values");
    await context.sync();

    const values = selection.values;
    if (!values || values.length < 2) {
      throw new Error("Please select ERP data in Excel first (header row + data rows).");
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

  // Run reconciliation
  const results = reconcile({ poData, poColumns, erpData, erpColumns, tolerance });
  agentState.results = results;

  // Write results sheet
  await writeResultsSheet(results, tolerance);

  // Format summary
  const s = results.summary;
  return `Reconciliation complete.\n\nTotal items: ${s.total}\nPerfect matches: ${s.matches}\nWithin tolerance: ${s.tolerances}\nExceptions: ${s.exceptions}\nWarnings: ${s.warnings}\nTotal exposure: ${formatCurrency(s.exposure)}\n\nResults sheet created with color-coded status rows.`;
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

  // Detect columns and extract
  const columns = detectAllColumns(parsedData.headers);
  const extraction = extractPOData(parsedData, columns);

  // Write staging sheet
  await writeStagingSheet(extraction);

  // Store for potential follow-up reconciliation
  agentState.lastExtraction = extraction;

  // Format response
  const m = extraction.metadata;
  const detectedStr = m.detectedFields.join(", ");
  let response = `PO data extracted and staging sheet created.\n\n`;
  response += `PO Reference: ${m.poRef}\n`;
  response += `Customer: ${m.customer}\n`;
  response += `Line items: ${m.lineCount}\n`;
  response += `Total value: ${formatCurrency(m.totalValue)}\n`;
  response += `Detected fields: ${detectedStr}\n`;

  if (m.warningCount > 0) {
    response += `\nWarnings: ${m.warningCount}\n`;
    const topWarnings = extraction.warnings.slice(0, 5);
    for (const w of topWarnings) {
      response += `  Line ${w.line}: ${w.field} — ${w.message}\n`;
    }
    if (extraction.warnings.length > 5) {
      response += `  ... and ${extraction.warnings.length - 5} more (see Warnings sheet)\n`;
    }
    response += `\nYellow-highlighted rows need review. Check cell notes for details.`;
  } else {
    response += `\nNo warnings — all data looks clean.`;
  }

  return response;
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
