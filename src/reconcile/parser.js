import Papa from "papaparse";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { detectColumns } from "./detector";

// Disable worker — runs synchronously, fine for small PO documents
pdfjsLib.GlobalWorkerOptions.workerSrc = "";

// Known column name patterns to identify the real header row
const HEADER_KEYWORDS = [
  "sku", "item", "product", "part", "material", "article", "upc", "ordered",
  "price", "cost", "amount", "rate", "each", "total",
  "qty", "quantity", "description", "name",
  "unit",
];

/**
 * Parse an uploaded file into { headers: string[], rows: Record<string, string>[] }
 * Tries multiple candidate header rows until one yields recognizable SKU+Price columns.
 */
export async function parseFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();

  switch (ext) {
    case "csv":
    case "tsv": {
      const rawRows = await parseCSVRaw(file);
      return findBestTable(rawRows);
    }
    case "xlsx":
    case "xls":
      return parseExcelAllSheets(file);
    case "pdf":
      return parsePDF(file);
    default:
      throw new Error(`Unsupported file type: .${ext}. Please use .xlsx, .csv, or .pdf.`);
  }
}

/**
 * Try every sheet in the workbook — pick the first one where we auto-detect SKU+Price
 * with data rows. Falls back to the sheet with the most data-like content.
 */
async function parseExcelAllSheets(file) {
  const buffer = await file.arrayBuffer();
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array" });
  } catch {
    throw new Error("Could not read Excel file. Check it's not open in another program or corrupted.");
  }

  if (workbook.SheetNames.length === 0) {
    throw new Error("Excel file has no sheets.");
  }

  let bestFallback = null; // best manual-selection result across all sheets

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rawRows.length < 2) continue;

    const result = findBestTable(rawRows, false);
    if (result.autoDetected) {
      return result;
    }

    // Track the best fallback (most columns = most likely real header)
    if (!bestFallback || result.headers.length > bestFallback.headers.length) {
      bestFallback = result;
    }
  }

  if (bestFallback) {
    return bestFallback;
  }

  throw new Error("Could not find a valid data table in any sheet.");
}

/**
 * Given raw rows from a single sheet/CSV, find the best header row.
 * If throwOnFail=true, throws when no rows found. Otherwise returns best-effort result.
 */
function findBestTable(rawRows, throwOnFail = true) {
  if (rawRows.length < 2) {
    if (throwOnFail) throw new Error("File needs at least a header row and one data row.");
    return { headers: [], rows: [], autoDetected: false };
  }

  const candidates = rankHeaderCandidates(rawRows);

  // Try each candidate — use the first one where column detection finds SKU + Price
  // AND has at least 1 data row below it
  for (const idx of candidates) {
    const { headers, colMap } = extractHeaders(rawRows[idx]);
    if (headers.length < 2) continue;

    const detected = detectColumns(headers);
    if (detected.sku && detected.price) {
      const rows = buildRows(rawRows, idx, headers, colMap);
      if (rows.length > 0) {
        return { headers, rows, autoDetected: true };
      }
    }
  }

  // No auto-detect success — pick the candidate with the MOST columns for manual selection
  // (a real data header typically has 5+ columns)
  const fallbackCandidates = [...candidates].sort((a, b) => {
    const colsA = extractHeaders(rawRows[a]).headers.length;
    const colsB = extractHeaders(rawRows[b]).headers.length;
    return colsB - colsA;
  });

  for (const idx of fallbackCandidates) {
    const { headers, colMap } = extractHeaders(rawRows[idx]);
    if (headers.length < 2) continue;
    const rows = buildRows(rawRows, idx, headers, colMap);
    if (rows.length > 0) {
      return { headers, rows, autoDetected: false };
    }
  }

  if (throwOnFail) {
    throw new Error("Could not find a valid header row in the file.");
  }
  return { headers: [], rows: [], autoDetected: false };
}

/**
 * Extract non-empty headers and their original column indices.
 * Preserves the mapping so data rows are read from the correct columns.
 */
function extractHeaders(rawRow) {
  const headers = [];
  const colMap = []; // colMap[i] = original column index for headers[i]
  for (let i = 0; i < rawRow.length; i++) {
    const val = String(rawRow[i] ?? "").trim();
    if (val) {
      headers.push(val);
      colMap.push(i);
    }
  }
  return { headers, colMap };
}

function buildRows(rawRows, headerIdx, headers, colMap) {
  return rawRows.slice(headerIdx + 1)
    .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        const col = colMap[i];
        obj[h] = row[col] != null ? String(row[col]) : "";
      });
      return obj;
    });
}

/**
 * Rank all rows by how likely they are to be a data table header.
 * Returns array of row indices, best first.
 */
function rankHeaderCandidates(rawRows) {
  const scored = [];

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row) continue;

    const nonEmpty = row.filter((c) => c != null && String(c).trim() !== "");
    if (nonEmpty.length < 2) continue;

    let keywordHits = 0;
    let shortCells = 0;

    for (const cell of nonEmpty) {
      const val = String(cell).toLowerCase().trim();
      if (val.length > 40) continue; // skip long text (instructions/sentences)
      shortCells++;
      for (const keyword of HEADER_KEYWORDS) {
        if (val.includes(keyword)) {
          keywordHits++;
          break;
        }
      }
    }

    // Score heavily weights: keyword matches + number of short columns
    // A real data header has many columns (5+) with several keyword matches
    const score = keywordHits * 10 + shortCells;
    scored.push({ idx: i, score, shortCells, keywordHits });
  }

  // Sort by score descending, then by number of short cells descending
  scored.sort((a, b) => b.score - a.score || b.shortCells - a.shortCells);

  return scored.map((s) => s.idx);
}

function parseCSVRaw(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: false,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete(results) {
        if (results.errors.length > 0 && results.data.length === 0) {
          reject(new Error("Could not parse CSV file. Check the file format."));
          return;
        }
        resolve(results.data);
      },
      error(err) {
        reject(new Error(`Could not read file. ${err.message}`));
      },
    });
  });
}


// ── PDF positional parsing ──
// Uses text item coordinates from pdf.js to reconstruct table rows and columns,
// rather than naive text joining which loses all layout information.

const PDF_Y_TOLERANCE = 3;    // Points — items within this Y delta are on the same line
const PDF_ITEM_GAP = 8;       // Points — X gap to consider items as separate cells
const PDF_COL_TOLERANCE = 20; // Points — X tolerance for clustering cells into columns

async function parsePDF(file) {
  let pdf;
  try {
    const buffer = await file.arrayBuffer();
    pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  } catch {
    throw new Error("Cannot parse this PDF. Please export to Excel/CSV first.");
  }

  // Collect positioned text items across all pages
  const allItems = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();

    for (const item of textContent.items) {
      const text = item.str;
      if (!text || !text.trim()) continue;
      allItems.push({
        x: item.transform[4],
        y: viewport.height - item.transform[5], // flip Y to top-down
        text: text.trim(),
        width: item.width || 0,
      });
    }
  }

  if (allItems.length === 0) {
    throw new Error("Cannot parse this PDF — no text found. Please export to Excel/CSV first.");
  }

  // Group items into rows by Y proximity, detect columns, build 2D array
  const posRows = groupPdfRows(allItems);
  const rawRows = alignPdfColumns(posRows);

  if (rawRows.length < 2) {
    throw new Error("Cannot parse this PDF. Please export to Excel/CSV first.");
  }

  return findBestTable(rawRows);
}

/**
 * Group text items into rows by Y coordinate proximity.
 * Returns array of { y, items[] }, sorted top-to-bottom.
 */
function groupPdfRows(items) {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);

  const rows = [];
  for (const item of sorted) {
    const lastRow = rows[rows.length - 1];
    if (lastRow && Math.abs(lastRow.y - item.y) <= PDF_Y_TOLERANCE) {
      lastRow.items.push(item);
    } else {
      rows.push({ y: item.y, items: [item] });
    }
  }

  return rows;
}

/**
 * Convert positioned rows into a column-aligned 2D string array.
 * 1. Within each row, merge adjacent items into cells (gap < PDF_ITEM_GAP)
 * 2. Cluster cell X positions across all rows into columns
 * 3. Assign cells to columns → 2D string array for findBestTable()
 */
function alignPdfColumns(rows) {
  // Step 1: Merge items within each row into cells
  const rowCells = rows.map((row) => {
    const sorted = [...row.items].sort((a, b) => a.x - b.x);
    if (sorted.length === 0) return [];

    const cells = [];
    let text = sorted[0].text;
    let x = sorted[0].x;
    let right = sorted[0].x + sorted[0].width;

    for (let i = 1; i < sorted.length; i++) {
      const gap = sorted[i].x - right;
      if (gap > PDF_ITEM_GAP) {
        cells.push({ text, x });
        text = sorted[i].text;
        x = sorted[i].x;
      } else {
        text += " " + sorted[i].text;
      }
      right = Math.max(right, sorted[i].x + sorted[i].width);
    }
    cells.push({ text, x });
    return cells;
  });

  // Step 2: Cluster all cell X positions into column anchors
  const columns = [];
  for (const cells of rowCells) {
    for (const cell of cells) {
      const match = columns.find((c) => Math.abs(c.center - cell.x) <= PDF_COL_TOLERANCE);
      if (match) {
        match.count++;
        match.center += (cell.x - match.center) / match.count; // running average
      } else {
        columns.push({ center: cell.x, count: 1 });
      }
    }
  }
  columns.sort((a, b) => a.center - b.center);

  if (columns.length === 0) return [];

  // Column boundaries = midpoints between adjacent column centers
  const bounds = [-Infinity];
  for (let i = 0; i < columns.length - 1; i++) {
    bounds.push((columns[i].center + columns[i + 1].center) / 2);
  }
  bounds.push(Infinity);

  // Step 3: Assign cells to columns → 2D array
  return rowCells.map((cells) => {
    const row = new Array(columns.length).fill("");
    for (const cell of cells) {
      for (let c = 0; c < columns.length; c++) {
        if (cell.x >= bounds[c] && cell.x < bounds[c + 1]) {
          row[c] = row[c] ? row[c] + " " + cell.text : cell.text;
          break;
        }
      }
    }
    return row;
  });
}
