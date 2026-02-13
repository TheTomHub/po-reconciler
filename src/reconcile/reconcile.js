import { parseNumber } from "../utils/format";

/**
 * Core reconciliation engine.
 *
 * Input: { poData, poColumns, oracleData, oracleColumns, tolerance }
 * Output: { summary, rows }
 */
export function reconcile({ poData, poColumns, oracleData, oracleColumns, tolerance }) {
  // Build Oracle lookup map: normalizedSKU -> { price, name, originalRow, matched }
  const oracleMap = new Map();
  const oracleDuplicates = new Set();

  for (const row of oracleData.rows) {
    const rawSku = row[oracleColumns.sku];
    if (!rawSku) continue;
    const normSku = normalizeSku(rawSku);

    if (oracleMap.has(normSku)) {
      oracleDuplicates.add(normSku);
    }

    oracleMap.set(normSku, {
      price: parseNumber(row[oracleColumns.price]),
      name: oracleColumns.name ? row[oracleColumns.name] || "" : "",
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
    const isDuplicate = poDuplicates.has(normSku) || oracleDuplicates.has(normSku);

    if (poPrice === null) {
      warnings++;
      resultRows.push({
        status: "Warning",
        sku: rawSku,
        name: poName,
        oraclePrice: null,
        poPrice: null,
        diff: null,
        pctDiff: null,
        action: "Non-numeric price — skipped",
        duplicate: isDuplicate,
      });
      continue;
    }

    // Try exact match first, then prefix match (customer core number → Oracle full SKU)
    let oracle = oracleMap.get(normSku);
    let matchType = "exact";

    if (!oracle) {
      const prefixMatches = findPrefixMatches(normSku, oracleMap);

      if (prefixMatches.length === 1) {
        // Single prefix match — use it
        oracle = prefixMatches[0].entry;
        matchType = "prefix";
      } else if (prefixMatches.length > 1) {
        // Multiple Oracle SKUs match this core number — flag for manual review
        const matchedSkus = prefixMatches.map((m) => m.sku).join(", ");
        warnings++;
        resultRows.push({
          status: "Warning",
          sku: rawSku,
          name: poName,
          oraclePrice: null,
          poPrice,
          diff: null,
          pctDiff: null,
          action: `Multiple Oracle matches: ${matchedSkus}`,
          duplicate: isDuplicate,
        });
        continue;
      }
    }

    if (!oracle) {
      exceptions++;
      resultRows.push({
        status: "Not in Oracle",
        sku: rawSku,
        name: poName,
        oraclePrice: null,
        poPrice,
        diff: null,
        pctDiff: null,
        action: "Review — SKU not found in Oracle",
        duplicate: isDuplicate,
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
        oraclePrice: null,
        poPrice,
        diff: null,
        pctDiff: null,
        action: "Non-numeric Oracle price",
        duplicate: isDuplicate,
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
      oracleSku: matchType === "prefix" ? denormalizeSku(findPrefixMatches(normSku, oracleMap)[0].sku, oracleData.rows, oracleColumns.sku) : null,
      name: oracle.name || poName,
      oraclePrice: oracle.price,
      poPrice,
      diff,
      pctDiff,
      action: matchType === "prefix" && action === "OK" ? "OK (prefix match)" : action,
      duplicate: isDuplicate,
    });
  }

  // Unmatched Oracle rows
  for (const [normSku, oracle] of oracleMap) {
    if (oracle.matched) continue;

    resultRows.push({
      status: "Not in PO",
      sku: denormalizeSku(normSku, oracleData.rows, oracleColumns.sku),
      name: oracle.name,
      oraclePrice: oracle.price,
      poPrice: null,
      diff: null,
      pctDiff: null,
      action: "Review — SKU not in PO",
      duplicate: oracleDuplicates.has(normSku),
    });
  }

  // Sort: exceptions first, then tolerance, then matches; within each group by |diff| desc
  const statusOrder = {
    Exception: 0,
    "Not in Oracle": 1,
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
 * Find Oracle SKUs that start with the given PO core number.
 * e.g. PO "1234" matches Oracle "1234V012", "1234V013"
 */
function findPrefixMatches(normPoSku, oracleMap) {
  const results = [];
  for (const [oracleSku, entry] of oracleMap) {
    if (oracleSku.startsWith(normPoSku) && oracleSku !== normPoSku) {
      results.push({ sku: oracleSku, entry });
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
