/**
 * Generate test .xlsx files for PO Reconciler testing.
 * Run: node test-data/generate.js
 */
const XLSX = require("xlsx");
const path = require("path");

const oracleRows = [
  { SKU: "TEST001", "Product Name": "Widget Alpha", "Unit Price": 10.00 },
  { SKU: "TEST002", "Product Name": "Widget Beta", "Unit Price": 20.00 },
  { SKU: "TEST003", "Product Name": "Widget Gamma", "Unit Price": 15.00 },
  { SKU: "TEST004", "Product Name": "Widget Delta", "Unit Price": 8.50 },
  { SKU: "TEST005", "Product Name": "Widget Epsilon", "Unit Price": 12.75 },
  { SKU: "TEST006", "Product Name": "Widget Zeta", "Unit Price": 22.00 },
  { SKU: "TEST007", "Product Name": "Widget Eta", "Unit Price": 5.99 },
  { SKU: "TEST008", "Product Name": "Widget Theta", "Unit Price": 18.50 },
  { SKU: "TEST009", "Product Name": "Widget Iota", "Unit Price": 33.00 },
  { SKU: "TEST010", "Product Name": "Widget Kappa", "Unit Price": 7.25 },
  { SKU: "TEST011", "Product Name": "Widget Lambda", "Unit Price": 45.00 },
  { SKU: "TEST012", "Product Name": "Widget Mu", "Unit Price": 9.99 },
  { SKU: "TEST013", "Product Name": "Widget Nu", "Unit Price": 14.50 },
  { SKU: "TEST014", "Product Name": "Widget Xi", "Unit Price": 27.75 },
  { SKU: "TEST015", "Product Name": "Widget Omicron", "Unit Price": 11.00 },
  { SKU: "TEST016", "Product Name": "Widget Pi", "Unit Price": 6.50 },
  { SKU: "TEST017", "Product Name": "Widget Rho", "Unit Price": 19.99 },
  { SKU: "TEST018", "Product Name": "Widget Sigma", "Unit Price": 31.25 },
  { SKU: "TEST019", "Product Name": "Widget Tau", "Unit Price": 4.75 },
  { SKU: "TEST020", "Product Name": "Widget Upsilon", "Unit Price": 16.00 },
];

// PO has 3 intentional price differences:
// TEST001: Oracle $10.00 vs PO $10.25 (exception - $0.25 diff)
// TEST002: Oracle $20.00 vs PO $20.01 (tolerance - $0.01 diff)
// TEST003: Oracle $15.00 vs PO $15.00 (exact match)
const poRows = oracleRows.map((row) => {
  const poRow = { ...row };
  if (row.SKU === "TEST001") poRow["Unit Price"] = 10.25;
  if (row.SKU === "TEST002") poRow["Unit Price"] = 20.01;
  // TEST003 stays at 15.00 — exact match
  return poRow;
});

// Write Oracle export
const oracleSheet = XLSX.utils.json_to_sheet(oracleRows);
const oracleWB = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(oracleWB, oracleSheet, "Oracle Export");
XLSX.writeFile(oracleWB, path.join(__dirname, "oracle_export.xlsx"));
console.log("Created oracle_export.xlsx");

// Write Customer PO
const poSheet = XLSX.utils.json_to_sheet(poRows);
const poWB = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(poWB, poSheet, "Customer PO");
XLSX.writeFile(poWB, path.join(__dirname, "customer_po.xlsx"));
console.log("Created customer_po.xlsx");
