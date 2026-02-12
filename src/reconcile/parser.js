import Papa from "papaparse";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

// Disable worker — runs synchronously, fine for small PO documents
pdfjsLib.GlobalWorkerOptions.workerSrc = "";

/**
 * Parse an uploaded file into { headers: string[], rows: Record<string, string>[] }
 */
export async function parseFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();

  switch (ext) {
    case "csv":
    case "tsv":
      return parseCSV(file);
    case "xlsx":
    case "xls":
      return parseExcel(file);
    case "pdf":
      return parsePDF(file);
    default:
      throw new Error(`Unsupported file type: .${ext}. Please use .xlsx, .csv, or .pdf.`);
  }
}

function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete(results) {
        if (results.errors.length > 0 && results.data.length === 0) {
          reject(new Error("Could not parse CSV file. Check the file format."));
          return;
        }
        const headers = results.meta.fields || [];
        if (headers.length === 0) {
          reject(new Error("CSV file has no headers. First row should contain column names."));
          return;
        }
        resolve({ headers, rows: results.data });
      },
      error(err) {
        reject(new Error(`Could not read file. ${err.message}`));
      },
    });
  });
}

async function parseExcel(file) {
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
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if (data.length < 2) {
    throw new Error("Excel file needs at least a header row and one data row.");
  }

  const headers = data[0].map((h) => String(h).trim());
  const rows = data.slice(1)
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] != null ? String(row[i]) : "";
      });
      return obj;
    });

  return { headers, rows };
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

  // Try to detect tabular data: find a line with multiple tab/multi-space separators
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
  // Try tab first, then multi-space, then pipe
  const delimiters = ["\t", /\s{2,}/, "|"];
  for (const d of delimiters) {
    const headerParts = lines[0].split(d).filter(Boolean);
    if (headerParts.length >= 2) {
      // Check at least one data line also splits similarly
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
