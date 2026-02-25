// Chandlr license module
// Checks the user's license tier and gates premium features.
// Fail-open: if the API is unreachable, free tier is assumed.

// ── Update this after your first `npx vercel` deployment ──
const API_BASE = "https://chandlr-api.vercel.app"; // deployed

export const UPGRADE_URL = "https://thetomhub.github.io/po-reconciler/#pricing";

const STORAGE_KEY = "chandlr_license_key";

// Free tier returned when the API is unreachable or key is unknown.
const FREE_LICENSE = {
  valid: false,
  tier: "free",
  features: [
    "ExtractPOData",
    "ReconcilePO",
    "GenerateCreditNote",
    "GenerateReInvoice",
    "DraftExceptionEmail",
    "LookupSKU",
  ],
  lineLimit: 100,
};

// In-memory cache — cleared by saveLicenseKey()
let _license = null;

/**
 * Reads the stored license key, calls the API, caches and returns the result.
 * Safe to call multiple times — subsequent calls return the cached value.
 * Always resolves (never rejects) — network errors return FREE_LICENSE.
 */
export async function checkLicense() {
  if (_license) return _license;

  const key = getLicenseKey();

  try {
    const url = `${API_BASE}/api/license${key ? `?key=${encodeURIComponent(key)}` : ""}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    _license = await res.json();
  } catch {
    // Fail-open: network error → free tier
    _license = { ...FREE_LICENSE };
  }

  return _license;
}

/**
 * Returns true if the current license includes the named feature.
 * Reads the cache — call checkLicense() first to populate it.
 * If cache is empty (e.g. called before checkLicense resolves), falls back to FREE_LICENSE.
 */
export function hasFeature(featureName) {
  const lic = _license || FREE_LICENSE;
  return Array.isArray(lic.features) && lic.features.includes(featureName);
}

/** Returns the current tier string ("free" | "pro" | "enterprise"). */
export function getTier() {
  return (_license || FREE_LICENSE).tier;
}

/** Returns the line limit (100 for free, 0 = unlimited for paid). */
export function getLineLimit() {
  return (_license || FREE_LICENSE).lineLimit;
}

/** Saves a license key to localStorage and clears the in-memory cache. */
export function saveLicenseKey(key) {
  if (typeof localStorage !== "undefined") {
    if (key) {
      localStorage.setItem(STORAGE_KEY, key.trim());
    } else {
      localStorage.removeItem(STORAGE_KEY);
    }
  }
  _license = null; // force re-check on next checkLicense()
}

/** Reads the stored license key from localStorage. */
export function getLicenseKey() {
  if (typeof localStorage !== "undefined") {
    return localStorage.getItem(STORAGE_KEY) || "";
  }
  return "";
}

/**
 * Returns a user-facing upgrade message for a gated feature.
 * Used in Copilot responses when a premium function is called without a Pro key.
 */
export function getUpgradeMessage(featureName) {
  return (
    `${featureName} is a Pro feature.\n\n` +
    `Upgrade to Chandlr Pro to unlock ERP staging, price intelligence, risk scoring, and the dashboard.\n\n` +
    `Get a license key at: ${UPGRADE_URL}\n\n` +
    `Once you have a key, enter it in the task pane under Settings → License.`
  );
}
