const SKU_ALIASES = [
  "sku",
  "item number",
  "product code",
  "part number",
  "item no",
  "item #",
  "material",
  "product id",
  "article",
  "upc",
];

const PRICE_ALIASES = [
  "price",
  "unit price",
  "cost",
  "amount",
  "unit cost",
  "net price",
  "each",
  "rate",
  "ext price",
];

const NAME_ALIASES = [
  "product name",
  "description",
  "item description",
  "product description",
  "name",
  "item name",
  "product",
];

/**
 * Auto-detect SKU, Price, and optional Name columns from headers.
 * Returns { sku: string|null, price: string|null, name: string|null }
 */
export function detectColumns(headers) {
  return {
    sku: findColumn(headers, SKU_ALIASES),
    price: findColumn(headers, PRICE_ALIASES),
    name: findColumn(headers, NAME_ALIASES),
  };
}

function findColumn(headers, aliases) {
  const normalized = headers.map((h) => h.toLowerCase().trim());

  // Exact match first
  for (const alias of aliases) {
    const idx = normalized.indexOf(alias);
    if (idx !== -1) return headers[idx];
  }

  // Partial match (header contains alias)
  for (const alias of aliases) {
    const idx = normalized.findIndex((h) => h.includes(alias));
    if (idx !== -1) return headers[idx];
  }

  // Partial match (alias contains header) — for short headers like "sku"
  for (const alias of aliases) {
    const idx = normalized.findIndex((h) => alias.includes(h) && h.length >= 3);
    if (idx !== -1) return headers[idx];
  }

  return null;
}
