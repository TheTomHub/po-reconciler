/**
 * Currency configuration — configurable via Settings dropdown.
 */
const CURRENCY_CONFIG = {
  GBP: { locale: "en-GB", currency: "GBP", format: "£#,##0.00" },
  USD: { locale: "en-US", currency: "USD", format: "$#,##0.00" },
  EUR: { locale: "de-DE", currency: "EUR", format: "€#,##0.00" },
};

let activeCurrency = "GBP";

export function setCurrency(code) {
  if (CURRENCY_CONFIG[code]) activeCurrency = code;
}

export function getCurrency() {
  return activeCurrency;
}

/**
 * Return the Excel number format string for the active currency.
 */
export function getCurrencyFormat() {
  return CURRENCY_CONFIG[activeCurrency].format;
}

/**
 * Format a number as a currency string using the active currency.
 */
export function formatCurrency(value) {
  if (value == null) return "—";
  const num = typeof value === "number" ? value : parseFloat(value);
  if (isNaN(num)) return "—";
  const cfg = CURRENCY_CONFIG[activeCurrency];
  return new Intl.NumberFormat(cfg.locale, {
    style: "currency",
    currency: cfg.currency,
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

  // Strip currency symbols ($, £, €), commas, spaces
  const cleaned = String(value).replace(/[$£€,\s]/g, "").trim();
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
