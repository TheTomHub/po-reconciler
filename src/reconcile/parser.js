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


async function parsePDF(file) {
  let pdf;
  try {
    const buffer = await file.arrayBuffer();
    pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  } catch {
    throw new Error("Cannot parse this PDF. Please export to Excel/CSV first.");
  }

  const lines = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ");
    lines.push(...pageText.split("\n").map((l) => l.trim()).filter(Boolean));
  }

  if (lines.length < 2) {
    throw new Error("Cannot parse this PDF. Please export to Excel/CSV first.");
  }

  const delimiter = detectPdfDelimiter(lines);
  if (!delimiter) {
    throw new Error("Cannot parse this PDF. No tabular data detected. Please export to Excel/CSV first.");
  }

  const headerLine = lines[0];
  const headers = headerLine.split(delimiter).map((h) => h.trim()).filter(Boolean);

  if (headers.length < 2) {
    throw new Error("Cannot parse this PDF. Please export to Excel/CSV first.");
  }

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split(delimiter).map((c) => c.trim());
    if (cells.length >= headers.length - 1) {
      const obj = {};
      headers.forEach((h, j) => {
        obj[h] = cells[j] || "";
      });
      rows.push(obj);
    }
  }

  if (rows.length === 0) {
    throw new Error("Cannot parse this PDF. No data rows found. Please export to Excel/CSV first.");
  }

  return { headers, rows };
}

function detectPdfDelimiter(lines) {
  const delimiters = ["\t", /\s{2,}/, "|"];
  for (const d of delimiters) {
    const headerParts = lines[0].split(d).filter(Boolean);
    if (headerParts.length >= 2) {
      for (let i = 1; i < Math.min(lines.length, 5); i++) {
        const parts = lines[i].split(d).filter(Boolean);
        if (parts.length >= headerParts.length - 1) {
          return d;
        }
      }
    }
  }
  return null;
}
