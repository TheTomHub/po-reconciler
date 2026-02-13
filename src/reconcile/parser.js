import Papa from "papaparse";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

// Disable worker — runs synchronously, fine for small PO documents
pdfjsLib.GlobalWorkerOptions.workerSrc = "";

// Known column name patterns to identify the real header row
const HEADER_KEYWORDS = [
  "sku", "item", "product", "part", "material", "article", "upc", "ordered",
  "price", "cost", "amount", "rate", "each", "total",
  "qty", "quantity", "description", "name",
  "unit", "status", "notes", "rule", "due",
];

/**
 * Parse an uploaded file into { headers: string[], rows: Record<string, string>[] }
 * Automatically detects the real header row, skipping instruction/metadata rows at the top.
 */
export async function parseFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();

  let rawRows;
  switch (ext) {
    case "csv":
    case "tsv":
      rawRows = await parseCSVRaw(file);
      break;
    case "xlsx":
    case "xls":
      rawRows = await parseExcelRaw(file);
      break;
    case "pdf":
      return parsePDF(file);
    default:
      throw new Error(`Unsupported file type: .${ext}. Please use .xlsx, .csv, or .pdf.`);
  }

  if (rawRows.length < 2) {
    throw new Error("File needs at least a header row and one data row.");
  }

  // Find the real header row by scanning for recognizable column names
  const headerIdx = findHeaderRow(rawRows);
  const headers = rawRows[headerIdx].map((h) => String(h).trim()).filter(Boolean);

  if (headers.length < 2) {
    throw new Error("Could not find a valid header row in the file.");
  }

  // Everything after the header row is data
  const rows = rawRows.slice(headerIdx + 1)
    .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] != null ? String(row[i]) : "";
      });
      return obj;
    });

  return { headers, rows };
}

/**
 * Scan rows to find which one is the actual header row.
 * Header rows have: multiple non-empty short cells, with recognized keywords.
 * Instruction rows have: 1-2 cells with long text (sentences).
 */
function findHeaderRow(rawRows) {
  let bestIdx = 0;
  let bestScore = -1;

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row) continue;

    const nonEmpty = row.filter((c) => c != null && String(c).trim() !== "");
    if (nonEmpty.length < 2) continue;

    let keywordHits = 0;
    let shortCells = 0;

    for (const cell of nonEmpty) {
      const val = String(cell).toLowerCase().trim();

      // Column headers are short (under 40 chars typically)
      if (val.length <= 40) shortCells++;

      // Only count keyword matches in short cells (skip sentences)
      if (val.length <= 40) {
        for (const keyword of HEADER_KEYWORDS) {
          if (val.includes(keyword)) {
            keywordHits++;
            break;
          }
        }
      }
    }

    // Score: keyword matches + bonus for having many short non-empty cells
    // A real header row has 3+ short cells with 2+ keyword hits
    const score = keywordHits * 3 + shortCells;

    // Require at least 2 keyword hits to be considered a header
    if (keywordHits >= 2 && score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  // If no row got 2+ keyword hits, fall back to row with the MOST non-empty short cells
  if (bestScore === -1) {
    let fallbackIdx = 0;
    let fallbackMax = 0;
    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row) continue;
      const nonEmpty = row.filter((c) => c != null && String(c).trim() !== "");
      const shortCells = nonEmpty.filter((c) => String(c).trim().length <= 40);
      if (shortCells.length > fallbackMax) {
        fallbackMax = shortCells.length;
        fallbackIdx = i;
      }
    }
    return fallbackIdx;
  }

  return bestIdx;
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

async function parseExcelRaw(file) {
  const buffer = await file.arrayBuffer();
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "array" });
  } catch {
    throw new Error("Could not read Excel file. Check it's not open in another program or corrupted.");
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Excel file has no sheets.");
  }

  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
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
