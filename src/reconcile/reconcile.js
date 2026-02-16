import { parseNumber } from "../utils/format";

/**
 * Core reconciliation engine.
 *
 * Input: { poData, poColumns, erpData, erpColumns, tolerance }
 * Output: { summary, rows }
 */
export function reconcile({ poData, poColumns, erpData, erpColumns, tolerance }) {
  // Build ERP lookup map: normalizedSKU -> { price, name, originalRow, matched }
  const erpMap = new Map();
  const erpDuplicates = new Set();

  for (const row of erpData.rows) {
    const rawSku = row[erpColumns.sku];
    if (!rawSku) continue;
    const normSku = normalizeSku(rawSku);

    if (erpMap.has(normSku)) {
      erpDuplicates.add(normSku);
    }

    erpMap.set(normSku, {
      price: parseNumber(row[erpColumns.price]),
      name: erpColumns.name ? row[erpColumns.name] || "" : "",
      qty: erpColumns.qty ? parseNumber(row[erpColumns.qty]) || 1 : 1,
      originalRow: row,
      matched: false,
    });
  }

  // Track PO duplicates
  const poSkuCounts = new Map();
  for (const row of poData.rows) {
    const rawSku = row[poColumns.sku];
    if (!rawSku) continue;
    const normSku = normalizeSku(rawSku);
    poSkuCounts.set(normSku, (poSkuCounts.get(normSku) || 0) + 1);
  }
  const poDuplicates = new Set();
  for (const [sku, count] of poSkuCounts) {
    if (count > 1) poDuplicates.add(sku);
  }

  const resultRows = [];
  let matches = 0;
  let tolerances = 0;
  let exceptions = 0;
  let exposure = 0;
  let warnings = 0;

  // Process each PO row
  for (const row of poData.rows) {
    const rawSku = row[poColumns.sku];
    if (!rawSku) continue;

    const normSku = normalizeSku(rawSku);
    const poPrice = parseNumber(row[poColumns.price]);
    const poName = poColumns.name ? row[poColumns.name] || "" : "";
    const poQty = poColumns.qty ? parseNumber(row[poColumns.qty]) || 1 : 1;
    const isDuplicate = poDuplicates.has(normSku) || erpDuplicates.has(normSku);

    if (poPrice === null) {
      warnings++;
      resultRows.push({
        status: "Warning",
        sku: rawSku,
        name: poName,
        erpPrice: null,
        poPrice: null,
        diff: null,
        pctDiff: null,
        action: "Non-numeric price — skipped",
        duplicate: isDuplicate,
        poQty,
        erpQty: null,
        lineTotal: null,
      });
      continue;
    }

    // Try exact match first, then prefix match (customer core number → ERP full SKU)
    let oracle = erpMap.get(normSku);
    let matchType = "exact";

    if (!oracle) {
      const prefixMatches = findPrefixMatches(normSku, erpMap);

      if (prefixMatches.length === 1) {
        // Single prefix match — use it
        oracle = prefixMatches[0].entry;
        matchType = "prefix";
      } else if (prefixMatches.length > 1) {
        // Multiple ERP SKUs match this core number — flag for manual review
        const matchedSkus = prefixMatches.map((m) => m.sku).join(", ");
        warnings++;
        resultRows.push({
          status: "Warning",
          sku: rawSku,
          name: poName,
          erpPrice: null,
          poPrice,
          diff: null,
          pctDiff: null,
          action: `Multiple ERP matches: ${matchedSkus}`,
          duplicate: isDuplicate,
          poQty,
          erpQty: null,
          lineTotal: round(poPrice * poQty),
        });
        continue;
      }
    }

    if (!oracle) {
      exceptions++;
      resultRows.push({
        status: "Not in ERP",
        sku: rawSku,
        name: poName,
        erpPrice: null,
        poPrice,
        diff: null,
        pctDiff: null,
        action: "Review — SKU not found in ERP",
        duplicate: isDuplicate,
        poQty,
        erpQty: null,
        lineTotal: round(poPrice * poQty),
      });
      continue;
    }

    oracle.matched = true;

    if (oracle.price === null) {
      warnings++;
      resultRows.push({
        status: "Warning",
        sku: rawSku,
        name: oracle.name || poName,
        erpPrice: null,
        poPrice,
        diff: null,
        pctDiff: null,
        action: "Non-numeric ERP price",
        duplicate: isDuplicate,
        poQty,
        erpQty: oracle.qty,
        lineTotal: round(poPrice * poQty),
      });
      continue;
    }

    const diff = round(poPrice - oracle.price);
    const absDiff = Math.abs(diff);
    const pctDiff = oracle.price !== 0 ? round((diff / oracle.price) * 100) : (diff !== 0 ? 100 : 0);

    let status, action;

    if (absDiff === 0) {
      status = "Match";
      action = "OK";
      matches++;
    } else if (absDiff <= tolerance) {
      status = "Tolerance";
      action = "OK — within tolerance";
      tolerances++;
    } else {
      status = "Exception";
      action = "Review pricing";
      exceptions++;
      exposure += absDiff;
    }

    resultRows.push({
      status,
      sku: rawSku,
      erpSku: matchType === "prefix" ? denormalizeSku(findPrefixMatches(normSku, erpMap)[0].sku, erpData.rows, erpColumns.sku) : null,
      name: oracle.name || poName,
      erpPrice: oracle.price,
      poPrice,
      diff,
      pctDiff,
      action: matchType === "prefix" && action === "OK" ? "OK (prefix match)" : action,
      duplicate: isDuplicate,
      poQty,
      erpQty: oracle.qty,
      lineTotal: round(poPrice * poQty),
    });
  }

  // Unmatched ERP rows
  for (const [normSku, oracle] of erpMap) {
    if (oracle.matched) continue;

    resultRows.push({
      status: "Not in PO",
      sku: denormalizeSku(normSku, erpData.rows, erpColumns.sku),
      name: oracle.name,
      erpPrice: oracle.price,
      poPrice: null,
      diff: null,
      pctDiff: null,
      action: "Review — SKU not in PO",
      duplicate: erpDuplicates.has(normSku),
      poQty: null,
      erpQty: oracle.qty,
      lineTotal: null,
    });
  }

  // Sort: exceptions first, then tolerance, then matches; within each group by |diff| desc
  const statusOrder = {
    Exception: 0,
    "Not in ERP": 1,
    "Not in PO": 2,
    Warning: 3,
    Tolerance: 4,
    Match: 5,
  };

  resultRows.sort((a, b) => {
    const sa = statusOrder[a.status] ?? 99;
    const sb = statusOrder[b.status] ?? 99;
    if (sa !== sb) return sa - sb;
    const da = a.diff != null ? Math.abs(a.diff) : 0;
    const db = b.diff != null ? Math.abs(b.diff) : 0;
    return db - da;
  });

  return {
    summary: {
      total: resultRows.length,
      matches,
      tolerances,
      exceptions,
      exposure: round(exposure),
      warnings,
      timestamp: new Date().toISOString(),
    },
    rows: resultRows,
  };
}

/**
 * Find ERP SKUs that start with the given PO core number.
 * e.g. PO "1234" matches ERP "1234V012", "1234V013"
 */
function findPrefixMatches(normPoSku, erpMap) {
  const results = [];
  for (const [erpSku, entry] of erpMap) {
    if (erpSku.startsWith(normPoSku) && erpSku !== normPoSku) {
      results.push({ sku: erpSku, entry });
    }
  }
  return results;
}

function normalizeSku(sku) {
  return String(sku).trim().toUpperCase();
}

function denormalizeSku(normSku, rows, skuColumn) {
  for (const row of rows) {
    if (normalizeSku(row[skuColumn]) === normSku) {
      return row[skuColumn];
    }
  }
  return normSku;
}

function round(n) {
  return Math.round(n * 100) / 100;
}
