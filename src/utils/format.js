/**
 * Format a number as USD currency string.
 */
export function formatCurrency(value) {
  if (value == null) return "—";
  const num = typeof value === "number" ? value : parseFloat(value);
  if (isNaN(num)) return "—";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(num);
}

/**
 * Parse a string/number into a float, stripping currency symbols and commas.
 * Returns null if non-numeric.
 */
export function parseNumber(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") return isNaN(value) ? null : value;

  // Strip $, commas, spaces
  const cleaned = String(value).replace(/[$,\s]/g, "").trim();
  if (cleaned === "") return null;

  // Handle parentheses as negative: (123.45) -> -123.45
  const parenMatch = cleaned.match(/^\((.+)\)$/);
  if (parenMatch) {
    const num = parseFloat(parenMatch[1]);
    return isNaN(num) ? null : -num;
  }

  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}
