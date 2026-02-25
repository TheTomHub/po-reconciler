// Vercel serverless function — GET /api/license?key=xxx
// Returns tier, features, lineLimit, valid for a given license key.
// Reads LICENSES env var: JSON string of { "key": "tier", ... }
// Fail-open: unknown/missing key → free tier.

const TIER_CONFIG = {
  free: {
    features: [
      "ExtractPOData",
      "ReconcilePO",
      "GenerateCreditNote",
      "GenerateReInvoice",
      "DraftExceptionEmail",
      "LookupSKU",
    ],
    lineLimit: 100,
  },
  pro: {
    features: [
      "ExtractPOData",
      "ReconcilePO",
      "GenerateCreditNote",
      "GenerateReInvoice",
      "DraftExceptionEmail",
      "LookupSKU",
      "GenerateERPStaging",
      "GetPriceIntelligence",
      "GenerateDashboard",
      "AssessPORisk",
    ],
    lineLimit: 0,
  },
  enterprise: {
    features: [
      "ExtractPOData",
      "ReconcilePO",
      "GenerateCreditNote",
      "GenerateReInvoice",
      "DraftExceptionEmail",
      "LookupSKU",
      "GenerateERPStaging",
      "GetPriceIntelligence",
      "GenerateDashboard",
      "AssessPORisk",
      "SharePointAutoFetch",
    ],
    lineLimit: 0,
  },
};

const ALLOWED_ORIGIN = "https://thetomhub.github.io";

module.exports = function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const key = (req.query.key || "").trim();

  // Parse license map from env
  let licenseMap = {};
  try {
    if (process.env.LICENSES) {
      licenseMap = JSON.parse(process.env.LICENSES);
    }
  } catch {
    // Malformed env — treat as empty
  }

  const tier = (key && licenseMap[key]) || "free";
  const config = TIER_CONFIG[tier] || TIER_CONFIG.free;

  return res.status(200).json({
    valid: Boolean(key && licenseMap[key]),
    tier,
    features: config.features,
    lineLimit: config.lineLimit,
  });
}
