/**
 * Generate test files mimicking a real-world order form layout:
 * - Instructions / company header at top
 * - Customer info section
 * - Status legend table
 * - Section header "Ordered Items"
 * - Data table starting well below row 1
 * - Core SKUs in PO (e.g. "11080") vs full variant SKUs in ERP (e.g. "11080V012")
 *
 * Run: node test-data/generate-oxo.js
 */
const XLSX = require("xlsx");
const path = require("path");

// --- ERP Export (simple flat table with "Ordered Item" column) ---
const erpData = [
  { "Ordered Item": "11080V012", "Description": "Premium 3-Piece Peeler Set", "Unit Qty": 24, "Unit Price": 12.99 },
  { "Ordered Item": "11081V003", "Description": "Classic Swivel Peeler", "Unit Qty": 36, "Unit Price": 9.99 },
  { "Ordered Item": "11082V008", "Description": "Ergonomic Can Opener", "Unit Qty": 12, "Unit Price": 21.99 },
  { "Ordered Item": "11083V001", "Description": "Stainless Steel Garlic Press", "Unit Qty": 18, "Unit Price": 15.99 },
  { "Ordered Item": "11084V005", "Description": "Locking Tongs 12-Inch", "Unit Qty": 24, "Unit Price": 14.49 },
  { "Ordered Item": "11085V002", "Description": "Silicone Spatula", "Unit Qty": 30, "Unit Price": 11.49 },
  { "Ordered Item": "11086V010", "Description": "Balloon Whisk 11-Inch", "Unit Qty": 12, "Unit Price": 10.99 },
  { "Ordered Item": "11087V004", "Description": "3-Piece Mixing Bowl Set", "Unit Qty": 6, "Unit Price": 29.99 },
  { "Ordered Item": "11088V006", "Description": "Salad Spinner Large", "Unit Qty": 12, "Unit Price": 32.99 },
  { "Ordered Item": "11089V001", "Description": "Measuring Cups Set of 4", "Unit Qty": 36, "Unit Price": 9.49 },
  { "Ordered Item": "11090V003", "Description": "Bamboo Cutting Board", "Unit Qty": 24, "Unit Price": 18.99 },
  { "Ordered Item": "11091V007", "Description": "Stainless Steel Colander", "Unit Qty": 18, "Unit Price": 16.49 },
  { "Ordered Item": "11092V002", "Description": "Box Grater 4-Sided", "Unit Qty": 12, "Unit Price": 13.99 },
  { "Ordered Item": "11093V009", "Description": "Kitchen Shears Heavy Duty", "Unit Qty": 6, "Unit Price": 19.99 },
  { "Ordered Item": "11094V001", "Description": "Flexible Turner", "Unit Qty": 24, "Unit Price": 12.49 },
  { "Ordered Item": "11095V004", "Description": "Soup Ladle", "Unit Qty": 30, "Unit Price": 11.99 },
  { "Ordered Item": "11096V006", "Description": "Potato Masher", "Unit Qty": 18, "Unit Price": 13.49 },
  { "Ordered Item": "11097V002", "Description": "Ice Cream Scoop", "Unit Qty": 12, "Unit Price": 10.99 },
  { "Ordered Item": "11098V008", "Description": "Pizza Wheel Cutter", "Unit Qty": 24, "Unit Price": 14.99 },
  { "Ordered Item": "11099V003", "Description": "Bottle Opener", "Unit Qty": 36, "Unit Price": 7.99 },
];

const erpSheet = XLSX.utils.json_to_sheet(erpData);
const erpWB = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(erpWB, erpSheet, "Order Lines");
XLSX.writeFile(erpWB, path.join(__dirname, "erp_export.xlsx"));
console.log("Created erp_export.xlsx (20 items, 'Ordered Item' column)");

// --- PO Order Form (messy layout with instructions at top) ---
// Build raw rows to mimic a real supplier order form structure
const raw = [];

// Row 0-1: Company header
raw.push(["EU Warehouse Order Form", "", "", "", "", "", "", "", ""]);
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 2-3: Instructions
raw.push(["Please complete this form and return to your sales representative.", "", "", "", "", "", "", "", ""]);
raw.push(["All prices are in EUR. Minimum order quantity applies per item.", "", "", "", "", "", "", "", ""]);

// Row 4: blank
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 5-7: Customer info
raw.push(["Customer:", "Acme Distribution BV", "", "Date:", "2025-01-15", "", "", "", ""]);
raw.push(["Account #:", "EU-4820", "", "PO Number:", "PO-2025-0042", "", "", "", ""]);
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 8: Status legend header
raw.push(["Status", "Explanation", "Notes", "", "", "", "", "", ""]);

// Row 9-11: Status legend data
raw.push(["LIVE", "Currently available for ordering", "", "", "", "", "", "", ""]);
raw.push(["Discontinuing", "Available while stock lasts", "", "", "", "", "", "", ""]);
raw.push(["Discontinued", "No longer available", "", "", "", "", "", "", ""]);

// Row 12: blank
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 13: Section header
raw.push(["Ordered Items", "", "", "", "", "", "", "", ""]);

// Row 14: blank
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 15: DATA HEADER — this is the real table header
raw.push(["Item #", "Unit Qty", "Description", "Price", "Order Rule Check", "Order Rule", "Total Due", "Status", "Notes"]);

// Row 16+: Data rows — using CORE SKUs (no variant suffix)
// Intentional price differences:
//   11080: ERP 12.99 vs PO 13.50 (exception)
//   11081: ERP 9.99  vs PO 10.00 (tolerance at $0.02 threshold)
//   11082: ERP 21.99 vs PO 21.99 (exact match)
//   11090: ERP 18.99 vs PO 17.49 (exception — PO lower)
//   11099: ERP 7.99  vs PO 8.99  (exception)
const poItems = [
  ["11080", "24", "Premium 3-Piece Peeler Set", "13.50", "OK", "MOQ 12", "324.00", "LIVE", ""],
  ["11081", "36", "Classic Swivel Peeler", "10.00", "OK", "MOQ 12", "360.00", "LIVE", ""],
  ["11082", "12", "Ergonomic Can Opener", "21.99", "OK", "MOQ 6", "263.88", "LIVE", ""],
  ["11083", "18", "Stainless Steel Garlic Press", "15.99", "OK", "MOQ 6", "287.82", "LIVE", ""],
  ["11084", "24", "Locking Tongs 12-Inch", "14.49", "OK", "MOQ 12", "347.76", "LIVE", ""],
  ["11085", "30", "Silicone Spatula", "11.49", "OK", "MOQ 12", "344.70", "LIVE", ""],
  ["11086", "12", "Balloon Whisk 11-Inch", "10.99", "OK", "MOQ 6", "131.88", "LIVE", ""],
  ["11087", "6", "3-Piece Mixing Bowl Set", "29.99", "OK", "MOQ 3", "179.94", "LIVE", ""],
  ["11088", "12", "Salad Spinner Large", "32.99", "OK", "MOQ 6", "395.88", "LIVE", ""],
  ["11089", "36", "Measuring Cups Set of 4", "9.49", "OK", "MOQ 12", "341.64", "LIVE", ""],
  ["11090", "24", "Bamboo Cutting Board", "17.49", "OK", "MOQ 12", "419.76", "Discontinuing", "While stock lasts"],
  ["11091", "18", "Stainless Steel Colander", "16.49", "OK", "MOQ 6", "296.82", "LIVE", ""],
  ["11092", "12", "Box Grater 4-Sided", "13.99", "OK", "MOQ 6", "167.88", "LIVE", ""],
  ["11093", "6", "Kitchen Shears Heavy Duty", "19.99", "OK", "MOQ 3", "119.94", "LIVE", ""],
  ["11094", "24", "Flexible Turner", "12.49", "OK", "MOQ 12", "299.76", "LIVE", ""],
  ["11095", "30", "Soup Ladle", "11.99", "OK", "MOQ 12", "359.70", "LIVE", ""],
  ["11096", "18", "Potato Masher", "13.49", "OK", "MOQ 6", "242.82", "LIVE", ""],
  ["11097", "12", "Ice Cream Scoop", "10.99", "OK", "MOQ 6", "131.88", "LIVE", ""],
  ["11098", "24", "Pizza Wheel Cutter", "14.99", "OK", "MOQ 12", "359.76", "LIVE", ""],
  ["11099", "36", "Bottle Opener", "8.99", "OK", "MOQ 12", "323.64", "LIVE", ""],
];

poItems.forEach((item) => raw.push(item));

// Row 36: blank
raw.push(["", "", "", "", "", "", "", "", ""]);

// Row 37-38: Summary footer
raw.push(["", "", "", "", "", "Total Items:", "20", "", ""]);
raw.push(["", "", "", "", "", "Total Order Value:", "5699.24", "", ""]);

// Build the workbook with TWO sheets (cover + order form)
const poWB = XLSX.utils.book_new();

// Sheet 1: Cover / Instructions
const coverRows = [
  ["Supplier Co."],
  ["EU Warehouse Order Form"],
  [""],
  ["Instructions:"],
  ["1. Fill in the order quantities on the 'Order Form' sheet"],
  ["2. Ensure minimum order quantities are met"],
  ["3. Return completed form to your sales representative"],
  [""],
  ["For questions, contact orders@supplier.example.com"],
];
const coverSheet = XLSX.utils.aoa_to_sheet(coverRows);
XLSX.utils.book_append_sheet(poWB, coverSheet, "Cover");

// Sheet 2: The actual order form with messy layout
const orderSheet = XLSX.utils.aoa_to_sheet(raw);
XLSX.utils.book_append_sheet(poWB, orderSheet, "Order Form");

XLSX.writeFile(poWB, path.join(__dirname, "sample_po_order.xlsx"));
console.log("Created sample_po_order.xlsx (2 sheets, data header at row 16, core SKUs)");

console.log("\nExpected reconciliation results:");
console.log("  11080: ERP 12.99 vs PO 13.50 → Exception (+0.51)");
console.log("  11081: ERP 9.99  vs PO 10.00 → Tolerance (+0.01)");
console.log("  11082: ERP 21.99 vs PO 21.99 → Match");
console.log("  11090: ERP 18.99 vs PO 17.49 → Exception (-1.50)");
console.log("  11099: ERP 7.99  vs PO 8.99  → Exception (+1.00)");
console.log("  Remaining 15 items: exact match (via prefix SKU matching)");
console.log("  Total exposure: $3.02");
