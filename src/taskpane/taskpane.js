import "./taskpane.css";
import { parseFile } from "../reconcile/parser";
import { detectColumns } from "../reconcile/detector";
import { reconcile } from "../reconcile/reconcile";
import { writeResultsSheet } from "../reconcile/results";
import { generateEmailDraft, buildMailtoLink } from "../email/email";
import { formatCurrency } from "../utils/format";

/* global Office, Excel */

// App state
const state = {
  poData: null,       // { headers, rows }
  poColumns: null,    // { sku, price, name? }
  oracleData: null,   // { headers, rows }
  oracleColumns: null, // { sku, price, name? }
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
    // Browser preview mode — file upload for Oracle instead of range selection
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
    oracleStatus: document.getElementById("oracle-status"),
    toleranceInput: document.getElementById("tolerance-input"),
    reconcileBtn: document.getElementById("reconcile-btn"),
    progressSection: document.getElementById("progress-section"),
    progressFill: document.getElementById("progress-fill"),
    progressText: document.getElementById("progress-text"),
    resultsSection: document.getElementById("results-section"),
    resultTotal: document.getElementById("result-total"),
    resultMatches: document.getElementById("result-matches"),
    resultTolerances: document.getElementById("result-tolerances"),
    resultExceptions: document.getElementById("result-exceptions"),
    resultExposure: document.getElementById("result-exposure"),
    emailBtn: document.getElementById("email-btn"),
    emailSection: document.getElementById("email-section"),
    emailDraft: document.getElementById("email-draft"),
    copyEmailBtn: document.getElementById("copy-email-btn"),
    mailtoLink: document.getElementById("mailto-link"),
    emailStatus: document.getElementById("email-status"),
    errorSection: document.getElementById("error-section"),
    errorMessage: document.getElementById("error-message"),
    // Manual column selectors
    manualColumnsPo: document.getElementById("manual-columns-po"),
    poSkuCol: document.getElementById("po-sku-col"),
    poPriceCol: document.getElementById("po-price-col"),
    poNameCol: document.getElementById("po-name-col"),
    applyPoColumns: document.getElementById("apply-po-columns"),
    manualColumnsOracle: document.getElementById("manual-columns-oracle"),
    oracleSkuCol: document.getElementById("oracle-sku-col"),
    oraclePriceCol: document.getElementById("oracle-price-col"),
    oracleNameCol: document.getElementById("oracle-name-col"),
    applyOracleColumns: document.getElementById("apply-oracle-columns"),
    // Browser mode elements
    oracleRangeBody: document.getElementById("oracle-range-body"),
    oracleFileBody: document.getElementById("oracle-file-body"),
    oracleFileInput: document.getElementById("oracle-file-input"),
    oracleFileStatus: document.getElementById("oracle-file-status"),
    browserResultsTable: document.getElementById("browser-results-table"),
    resultsTable: document.getElementById("results-table"),
  };

  // In browser mode, swap "Select Range" for file upload
  if (browserMode) {
    els.oracleRangeBody.hidden = true;
    els.oracleFileBody.hidden = false;
    els.oracleFileInput.addEventListener("change", handleOracleFileUpload);
  }

  // Event listeners
  els.poFileInput.addEventListener("change", handleFileUpload);
  els.selectRangeBtn.addEventListener("click", handleSelectRange);
  els.toleranceInput.addEventListener("change", () => {
    state.tolerance = parseFloat(els.toleranceInput.value) || 0.02;
  });
  els.reconcileBtn.addEventListener("click", handleReconcile);
  els.emailBtn.addEventListener("click", handleShowEmail);
  els.copyEmailBtn.addEventListener("click", handleCopyEmail);
  els.applyPoColumns.addEventListener("click", () => applyManualColumns("po"));
  els.applyOracleColumns.addEventListener("click", () => applyManualColumns("oracle"));
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
      setStatus(els.poStatus, `${file.name} — select columns below`, "");
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

// --- Oracle File Upload (browser mode) ---

async function handleOracleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  hideError();
  setStatus(els.oracleFileStatus, "Parsing...", "");

  try {
    state.oracleData = await parseFile(file);
    const columns = detectColumns(state.oracleData.headers);

    if (!columns.sku || !columns.price) {
      showManualColumnSelection("oracle", state.oracleData.headers);
      setStatus(els.oracleFileStatus, `${file.name} — select columns below`, "");
      setStatus(els.oracleStatus, `${file.name} — select columns below`, "");
    } else {
      state.oracleColumns = columns;
      els.manualColumnsOracle.hidden = true;
      setStatus(els.oracleFileStatus, `${file.name} (${state.oracleData.rows.length} rows)`, "success");
      setStatus(els.oracleStatus, `${state.oracleData.rows.length} rows loaded`, "success");
    }
    updateReconcileButton();
  } catch (err) {
    state.oracleData = null;
    state.oracleColumns = null;
    setStatus(els.oracleFileStatus, "No file selected", "");
    showError(err.message);
  }
}

// --- Oracle Range Selection ---

async function handleSelectRange() {
  hideError();
  setStatus(els.oracleStatus, "Reading selection...", "");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      if (!values || values.length < 2) {
        throw new Error("Please select a range containing Oracle data with at least a header row and one data row.");
      }

      const headers = values[0].map((h) => String(h).trim());
      const rows = values.slice(1).map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i] != null ? String(row[i]) : "";
        });
        return obj;
      });

      state.oracleData = { headers, rows };
      const columns = detectColumns(headers);

      if (!columns.sku || !columns.price) {
        showManualColumnSelection("oracle", headers);
        setStatus(els.oracleStatus, `${rows.length} rows — select columns below`, "");
      } else {
        state.oracleColumns = columns;
        els.manualColumnsOracle.hidden = true;
        setStatus(els.oracleStatus, `${rows.length} rows loaded`, "success");
      }
      updateReconcileButton();
    });
  } catch (err) {
    state.oracleData = null;
    state.oracleColumns = null;
    setStatus(els.oracleStatus, "No range selected", "");
    showError(err.message);
  }
}

// --- Manual Column Selection ---

function showManualColumnSelection(source, headers) {
  const isOracle = source === "oracle";
  const skuSelect = isOracle ? els.oracleSkuCol : els.poSkuCol;
  const priceSelect = isOracle ? els.oraclePriceCol : els.poPriceCol;
  const nameSelect = isOracle ? els.oracleNameCol : els.poNameCol;
  const container = isOracle ? els.manualColumnsOracle : els.manualColumnsPo;

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

  // Name is optional — already has "None" default
  nameSelect.innerHTML = '<option value="">— None —</option>';
  headers.forEach((h) => {
    const opt = document.createElement("option");
    opt.value = h;
    opt.textContent = h;
    nameSelect.appendChild(opt);
  });

  container.hidden = false;
}

function applyManualColumns(source) {
  const isOracle = source === "oracle";
  const skuSelect = isOracle ? els.oracleSkuCol : els.poSkuCol;
  const priceSelect = isOracle ? els.oraclePriceCol : els.poPriceCol;
  const nameSelect = isOracle ? els.oracleNameCol : els.poNameCol;
  const statusEl = isOracle ? els.oracleStatus : els.poStatus;
  const container = isOracle ? els.manualColumnsOracle : els.manualColumnsPo;

  const sku = skuSelect.value;
  const price = priceSelect.value;
  const name = nameSelect.value || null;

  if (!sku || !price) {
    showError("Please select both SKU and Price columns.");
    return;
  }

  const columns = { sku, price, name };

  if (isOracle) {
    state.oracleColumns = columns;
    setStatus(statusEl, `${state.oracleData.rows.length} rows loaded`, "success");
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
      oracleData: state.oracleData,
      oracleColumns: state.oracleColumns,
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
    els.resultExposure.textContent = formatCurrency(results.summary.exposure);

    setProgress(100, "Complete!");
    els.resultsSection.hidden = false;

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

// --- Browser Results Table ---

function renderBrowserResultsTable(results) {
  const table = els.resultsTable;
  const headers = ["Status", "SKU", "Name", "Oracle $", "PO $", "Diff", "% Diff", "Action"];

  let html = "<thead><tr>";
  headers.forEach((h) => { html += `<th>${h}</th>`; });
  html += "</tr></thead><tbody>";

  results.rows.forEach((row) => {
    const statusClass =
      (row.status === "Exception" || row.status === "Not in Oracle" || row.status === "Not in PO")
        ? "row-exception"
        : row.status === "Tolerance" ? "row-tolerance"
        : row.status === "Match" ? "row-match"
        : "row-warning";

    const label = row.duplicate ? `${row.status} (DUP)` : row.status;

    html += `<tr class="${statusClass}">`;
    html += `<td>${label}</td>`;
    html += `<td>${row.oracleSku ? `${row.sku} → ${row.oracleSku}` : row.sku}</td>`;
    html += `<td>${row.name || ""}</td>`;
    html += `<td>${row.oraclePrice != null ? formatCurrency(row.oraclePrice) : ""}</td>`;
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
    state.oracleData && state.oracleColumns;
  els.reconcileBtn.disabled = !ready;
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
