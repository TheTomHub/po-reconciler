import "./taskpane.css";
import { parseFile } from "../reconcile/parser";
import { detectColumns } from "../reconcile/detector";
import { reconcile } from "../reconcile/reconcile";
import { writeResultsSheet } from "../reconcile/results";
import { generateEmailDraft, buildMailtoLink } from "../email/email";
import { generateCreditNote, generateCorrectedInvoice } from "../reconcile/creditnote";
import { writeCreditNoteSheet, writeReInvoiceSheet } from "../reconcile/creditnote-results";
import { formatCurrency, setCurrency } from "../utils/format";
import { detectAllColumns, extractPOData } from "../capture/extractor";
import { writeStagingSheet } from "../capture/staging";

/* global Office, Excel */

// App state
const state = {
  poData: null,       // { headers, rows }
  poColumns: null,    // { sku, price, name? }
  erpData: null,   // { headers, rows }
  erpColumns: null, // { sku, price, name? }
  tolerance: 0.02,
  results: null,
  poFilename: "",
  browserMode: false,
};

// DOM references (set after Office.onReady)
let els = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initUI(false);
  } else {
    // Browser preview mode — file upload for ERP instead of range selection
    initUI(true);
    console.log("Running in browser preview mode (no Excel APIs)");
  }
});

function initUI(browserMode) {
  state.browserMode = browserMode;

  els = {
    poFileInput: document.getElementById("po-file-input"),
    poStatus: document.getElementById("po-status"),
    selectRangeBtn: document.getElementById("select-range-btn"),
    erpStatus: document.getElementById("erp-status"),
    toleranceInput: document.getElementById("tolerance-input"),
    currencySelect: document.getElementById("currency-select"),
    reconcileBtn: document.getElementById("reconcile-btn"),
    progressSection: document.getElementById("progress-section"),
    progressFill: document.getElementById("progress-fill"),
    progressText: document.getElementById("progress-text"),
    resultsSection: document.getElementById("results-section"),
    resultTotal: document.getElementById("result-total"),
    resultMatches: document.getElementById("result-matches"),
    resultTolerances: document.getElementById("result-tolerances"),
    resultExceptions: document.getElementById("result-exceptions"),
    resultWarnings: document.getElementById("result-warnings"),
    resultExposure: document.getElementById("result-exposure"),
    emailBtn: document.getElementById("email-btn"),
    creditNoteBtn: document.getElementById("credit-note-btn"),
    reinvoiceBtn: document.getElementById("reinvoice-btn"),
    emailSection: document.getElementById("email-section"),
    emailDraft: document.getElementById("email-draft"),
    copyEmailBtn: document.getElementById("copy-email-btn"),
    mailtoLink: document.getElementById("mailto-link"),
    emailStatus: document.getElementById("email-status"),
    actionStatus: document.getElementById("action-status"),
    errorSection: document.getElementById("error-section"),
    errorMessage: document.getElementById("error-message"),
    // Manual column selectors
    manualColumnsPo: document.getElementById("manual-columns-po"),
    poSkuCol: document.getElementById("po-sku-col"),
    poPriceCol: document.getElementById("po-price-col"),
    poNameCol: document.getElementById("po-name-col"),
    poQtyCol: document.getElementById("po-qty-col"),
    applyPoColumns: document.getElementById("apply-po-columns"),
    manualColumnsErp: document.getElementById("manual-columns-erp"),
    erpSkuCol: document.getElementById("erp-sku-col"),
    erpPriceCol: document.getElementById("erp-price-col"),
    erpNameCol: document.getElementById("erp-name-col"),
    erpQtyCol: document.getElementById("erp-qty-col"),
    applyErpColumns: document.getElementById("apply-erp-columns"),
    // Browser mode elements
    browserResultsTable: document.getElementById("browser-results-table"),
    resultsTable: document.getElementById("results-table"),
    // Tab elements
    tabBtns: document.querySelectorAll(".tab"),
    panelReconcile: document.getElementById("panel-reconcile"),
    panelActions: document.getElementById("panel-actions"),
    actionsHint: document.getElementById("actions-hint"),
    actionsContent: document.getElementById("actions-content"),
    tabActions: document.querySelector('[data-tab="actions"]'),
    // Capture module
    extractBtn: document.getElementById("extract-btn"),
    extractStatus: document.getElementById("extract-status"),
  };

  // Event listeners
  els.poFileInput.addEventListener("change", handleFileUpload);
  els.selectRangeBtn.addEventListener("click", handleSelectRange);
  els.toleranceInput.addEventListener("change", () => {
    state.tolerance = parseFloat(els.toleranceInput.value) || 0.02;
  });
  els.currencySelect.addEventListener("change", () => {
    setCurrency(els.currencySelect.value);
  });
  els.reconcileBtn.addEventListener("click", handleReconcile);
  els.emailBtn.addEventListener("click", handleShowEmail);
  els.copyEmailBtn.addEventListener("click", handleCopyEmail);
  els.creditNoteBtn.addEventListener("click", handleCreditNote);
  els.reinvoiceBtn.addEventListener("click", handleReInvoice);
  els.extractBtn.addEventListener("click", handleExtract);
  els.applyPoColumns.addEventListener("click", () => applyManualColumns("po"));
  els.applyErpColumns.addEventListener("click", () => applyManualColumns("erp"));

  // Tab switching
  els.tabBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.disabled) return;
      switchTab(btn.dataset.tab);
    });
  });
}

// --- Tab Switching ---

function switchTab(tabName) {
  els.tabBtns.forEach((btn) => btn.classList.remove("active"));
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));

  const btn = document.querySelector(`.tab[data-tab="${tabName}"]`);
  const panel = document.getElementById(`panel-${tabName}`);
  if (btn) btn.classList.add("active");
  if (panel) panel.classList.add("active");

  // Clear notification dot when visiting the tab
  if (btn) btn.classList.remove("tab-notify");
}

// --- File Upload ---

async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  hideError();
  setStatus(els.poStatus, "Parsing...", "");
  state.poFilename = file.name;

  try {
    state.poData = await parseFile(file);
    const columns = detectColumns(state.poData.headers);

    if (!columns.sku || !columns.price) {
      showManualColumnSelection("po", state.poData.headers);
      setStatus(els.poStatus, `${file.name} — ${state.poData.headers.length} cols, ${state.poData.rows.length} rows — select columns below`, "");
    } else {
      state.poColumns = columns;
      els.manualColumnsPo.hidden = true;
      setStatus(els.poStatus, `${file.name} (${state.poData.rows.length} rows)`, "success");
    }
    updateReconcileButton();
  } catch (err) {
    state.poData = null;
    state.poColumns = null;
    setStatus(els.poStatus, "No file selected", "");
    showError(err.message);
  }
}

// --- ERP Range Selection ---

async function handleSelectRange() {
  hideError();
  setStatus(els.erpStatus, "Reading selection...", "");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      if (!values || values.length < 2) {
        throw new Error("Please select a range containing ERP data with at least a header row and one data row.");
      }

      const headers = values[0].map((h) => String(h).trim());
      const rows = values.slice(1).map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i] != null ? String(row[i]) : "";
        });
        return obj;
      });

      state.erpData = { headers, rows };
      const columns = detectColumns(headers);

      if (!columns.sku || !columns.price) {
        showManualColumnSelection("erp", headers);
        setStatus(els.erpStatus, `${rows.length} rows — select columns below`, "");
      } else {
        state.erpColumns = columns;
        els.manualColumnsErp.hidden = true;
        setStatus(els.erpStatus, `${rows.length} rows loaded`, "success");
      }
      updateReconcileButton();
    });
  } catch (err) {
    state.erpData = null;
    state.erpColumns = null;
    setStatus(els.erpStatus, "Select ERP data in Excel, then click Load", "");
    showError(err.message);
  }
}

// --- Manual Column Selection ---

function showManualColumnSelection(source, headers) {
  const isErp = source === "erp";
  const skuSelect = isErp ? els.erpSkuCol : els.poSkuCol;
  const priceSelect = isErp ? els.erpPriceCol : els.poPriceCol;
  const nameSelect = isErp ? els.erpNameCol : els.poNameCol;
  const qtySelect = isErp ? els.erpQtyCol : els.poQtyCol;
  const container = isErp ? els.manualColumnsErp : els.manualColumnsPo;

  // Populate dropdowns
  [skuSelect, priceSelect].forEach((select) => {
    select.innerHTML = '<option value="">— Select —</option>';
    headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.textContent = h;
      select.appendChild(opt);
    });
  });

  // Name and Qty are optional — have "None" default
  [nameSelect, qtySelect].forEach((select) => {
    select.innerHTML = '<option value="">— None —</option>';
    headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.textContent = h;
      select.appendChild(opt);
    });
  });

  container.hidden = false;
}

function applyManualColumns(source) {
  const isErp = source === "erp";
  const skuSelect = isErp ? els.erpSkuCol : els.poSkuCol;
  const priceSelect = isErp ? els.erpPriceCol : els.poPriceCol;
  const nameSelect = isErp ? els.erpNameCol : els.poNameCol;
  const qtySelect = isErp ? els.erpQtyCol : els.poQtyCol;
  const statusEl = isErp ? els.erpStatus : els.poStatus;
  const container = isErp ? els.manualColumnsErp : els.manualColumnsPo;

  const sku = skuSelect.value;
  const price = priceSelect.value;
  const name = nameSelect.value || null;
  const qty = qtySelect.value || null;

  if (!sku || !price) {
    showError("Please select both SKU and Price columns.");
    return;
  }

  const columns = { sku, price, name, qty };

  if (isErp) {
    state.erpColumns = columns;
    setStatus(statusEl, `${state.erpData.rows.length} rows loaded`, "success");
  } else {
    state.poColumns = columns;
    setStatus(statusEl, `${state.poFilename} (${state.poData.rows.length} rows)`, "success");
  }

  container.hidden = true;
  hideError();
  updateReconcileButton();
}

// --- Reconciliation ---

async function handleReconcile() {
  hideError();
  els.resultsSection.hidden = true;
  els.emailSection.hidden = true;
  els.progressSection.hidden = false;
  els.reconcileBtn.disabled = true;
  setProgress(10, "Preparing data...");

  try {
    state.tolerance = parseFloat(els.toleranceInput.value) || 0.02;

    setProgress(30, "Running reconciliation...");

    const results = reconcile({
      poData: state.poData,
      poColumns: state.poColumns,
      erpData: state.erpData,
      erpColumns: state.erpColumns,
      tolerance: state.tolerance,
    });

    state.results = results;

    setProgress(60, "Writing results...");

    if (state.browserMode) {
      renderBrowserResultsTable(results);
    } else {
      await writeResultsSheet(results, state.tolerance);
    }

    setProgress(90, "Finalizing...");

    // Display summary
    els.resultTotal.textContent = results.summary.total;
    els.resultMatches.textContent = results.summary.matches;
    els.resultTolerances.textContent = results.summary.tolerances;
    els.resultExceptions.textContent = results.summary.exceptions;
    els.resultWarnings.textContent = results.summary.warnings || 0;
    els.resultExposure.textContent = formatCurrency(results.summary.exposure);

    setProgress(100, "Complete!");
    els.resultsSection.hidden = false;

    // Enable Actions tab and show action buttons
    els.tabActions.disabled = false;
    els.tabActions.removeAttribute("title");
    els.actionsHint.hidden = true;
    els.actionsContent.hidden = false;
    els.tabActions.classList.add("tab-notify");

    // Show credit note buttons only when exceptions exist
    const hasExceptions = results.summary.exceptions > 0;
    els.creditNoteBtn.hidden = !hasExceptions;
    els.reinvoiceBtn.hidden = !hasExceptions;

    // Clear any previous action status
    setStatus(els.actionStatus, "", "");

    setTimeout(() => {
      els.progressSection.hidden = true;
    }, 1000);
  } catch (err) {
    showError(err.message);
    els.progressSection.hidden = true;
  } finally {
    els.reconcileBtn.disabled = false;
  }
}

// --- Extract to Staging ---

async function handleExtract() {
  if (!state.poData || !state.poColumns) return;

  hideError();
  els.extractBtn.disabled = true;
  setStatus(els.extractStatus, "Extracting...", "");

  try {
    const columns = detectAllColumns(state.poData.headers);
    const extraction = extractPOData(state.poData, columns);

    if (!state.browserMode) {
      await writeStagingSheet(extraction);
    }

    const m = extraction.metadata;
    let statusText = `Staging sheet created: ${m.lineCount} lines, ${formatCurrency(m.totalValue)}`;
    if (m.warningCount > 0) {
      statusText += ` (${m.warningCount} warnings)`;
    }
    setStatus(els.extractStatus, statusText, "success");
  } catch (err) {
    setStatus(els.extractStatus, "", "");
    showError(err.message);
  } finally {
    els.extractBtn.disabled = false;
  }
}

// --- Email ---

function handleShowEmail() {
  if (!state.results) return;

  const draft = generateEmailDraft(state.results, state.poFilename);
  els.emailDraft.textContent = draft.body;
  els.mailtoLink.href = buildMailtoLink(draft);
  els.emailSection.hidden = false;
}

async function handleCopyEmail() {
  const text = els.emailDraft.textContent;
  try {
    await navigator.clipboard.writeText(text);
    setStatus(els.emailStatus, "Copied!", "success");
    setTimeout(() => setStatus(els.emailStatus, "", ""), 2000);
  } catch {
    // Fallback for older browsers
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    setStatus(els.emailStatus, "Copied!", "success");
    setTimeout(() => setStatus(els.emailStatus, "", ""), 2000);
  }
}

// --- Credit Note & Re-Invoice ---

async function handleCreditNote() {
  if (!state.results) return;
  try {
    els.creditNoteBtn.disabled = true;
    setStatus(els.actionStatus, "Generating...", "");
    const creditData = generateCreditNote(state.results);
    await writeCreditNoteSheet(creditData, state.poFilename);
    setStatus(els.actionStatus, "Credit Note sheet created", "success");
    setTimeout(() => setStatus(els.actionStatus, "", ""), 4000);
  } catch (err) {
    showError(err.message);
  } finally {
    els.creditNoteBtn.disabled = false;
  }
}

async function handleReInvoice() {
  if (!state.results) return;
  try {
    els.reinvoiceBtn.disabled = true;
    setStatus(els.actionStatus, "Generating...", "");
    const invoiceData = generateCorrectedInvoice(state.results);
    await writeReInvoiceSheet(invoiceData, state.poFilename, state.results.summary.exceptions);
    setStatus(els.actionStatus, "Re-Invoice sheet created", "success");
    setTimeout(() => setStatus(els.actionStatus, "", ""), 4000);
  } catch (err) {
    showError(err.message);
  } finally {
    els.reinvoiceBtn.disabled = false;
  }
}

// --- Browser Results Table ---

function renderBrowserResultsTable(results) {
  const table = els.resultsTable;
  const headers = ["Status", "SKU", "Name", "ERP $", "PO $", "Diff", "% Diff", "Action"];

  let html = "<thead><tr>";
  headers.forEach((h) => { html += `<th>${h}</th>`; });
  html += "</tr></thead><tbody>";

  results.rows.forEach((row) => {
    const statusClass =
      (row.status === "Exception" || row.status === "Not in ERP" || row.status === "Not in PO")
        ? "row-exception"
        : row.status === "Tolerance" ? "row-tolerance"
        : row.status === "Match" ? "row-match"
        : "row-warning";

    const label = row.duplicate ? `${row.status} (DUP)` : row.status;

    html += `<tr class="${statusClass}">`;
    html += `<td>${label}</td>`;
    html += `<td>${row.erpSku ? `${row.sku} → ${row.erpSku}` : row.sku}</td>`;
    html += `<td>${row.name || ""}</td>`;
    html += `<td>${row.erpPrice != null ? formatCurrency(row.erpPrice) : ""}</td>`;
    html += `<td>${row.poPrice != null ? formatCurrency(row.poPrice) : ""}</td>`;
    html += `<td>${row.diff != null ? formatCurrency(row.diff) : ""}</td>`;
    html += `<td>${row.pctDiff != null ? row.pctDiff + "%" : ""}</td>`;
    html += `<td>${row.action}</td>`;
    html += "</tr>";
  });

  html += "</tbody>";
  table.innerHTML = html;
  els.browserResultsTable.hidden = false;
}

// --- UI Helpers ---

function updateReconcileButton() {
  const ready =
    state.poData && state.poColumns &&
    state.erpData && state.erpColumns;
  els.reconcileBtn.disabled = !ready;
  // Enable extract button when PO data is loaded (doesn't need ERP data)
  els.extractBtn.disabled = !(state.poData && state.poColumns);
}

function setStatus(el, text, type) {
  el.textContent = text;
  el.className = "status" + (type ? " " + type : "");
}

function setProgress(pct, text) {
  els.progressFill.style.width = pct + "%";
  els.progressText.textContent = text;
}

function showError(msg) {
  els.errorMessage.textContent = msg;
  els.errorSection.hidden = false;
}

function hideError() {
  els.errorSection.hidden = true;
}
