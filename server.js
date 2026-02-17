f/**
 * TGJU → Google Sheets (Standalone Project) — ESM (Option A)
 * - No changes needed to your WP plugin
 * - Fetches https://call2.tgju.org/ajax.json
 * - Writes rows into a defined Google Sheet tab
 * - Optional: ONLY_MAJOR=1 to keep only major fiat currencies
 *
 * Env required:
 *  PORT
 *  SHEET_ID
 *  WORKSHEET_TITLE
 *  GOOGLE_SERVICE_ACCOUNT_JSON_BASE64   (base64 of the service account JSON)
 * Optional:
 *  CACHE_TTL_MS
 *  ONLY_MAJOR
 */

import "dotenv/config";
import express from "express";
import cron from "node-cron";
import { GoogleSpreadsheet } from "google-spreadsheet";

const app = express();

const PORT = Number(process.env.PORT || 3000);
const TGJU_JSON_URL = "https://call2.tgju.org/ajax.json";

const SHEET_ID = process.env.SHEET_ID;
const WORKSHEET_TITLE = process.env.WORKSHEET_TITLE || "Rates";
const CACHE_TTL_MS = Number(process.env.CACHE_TTL_MS || 60_000);
const ONLY_MAJOR = String(process.env.ONLY_MAJOR || "0") === "1";

const SA_B64 = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "";

// -----------------------------
// Config: major fiat list only
// -----------------------------
const MAJOR_FIAT = new Set([
  "usd", "eur", "gbp", "cad", "aed", "try",
  "sar", "qar", "kwd", "bhd", "iqd",
  "cny", "jpy", "chf", "rub",
  "aud", "nzd", "sek", "nok", "dkk",
  "inr", "krw", "myr", "thb",
  "php", "mxn", "brl", "zar",
]);

// TGJU special fiat keys -> normalized codes
const SPECIAL_CODE_MAP = {
  dollar_rl: "usd",
  dollar_ex: "usd_official",
  dollar_dt: "usd_dt",
  dollar_sm: "usd_sm",
  eur_ex: "eur_official",
};

// Persian names (expand any time)
const FA_NAME_MAP = {
  usd: "دلار آمریکا",
  eur: "یورو",
  gbp: "پوند انگلیس",
  cad: "دلار کانادا",
  aed: "درهم امارات",
  try: "لیر ترکیه",
  sar: "ریال عربستان",
  qar: "ریال قطر",
  kwd: "دینار کویت",
  bhd: "دینار بحرین",
  iqd: "دینار عراق",
  cny: "یوان چین",
  jpy: "ین ژاپن",
  chf: "فرانک سوئیس",
  rub: "روبل روسیه",
  aud: "دلار استرالیا",
  nzd: "دلار نیوزیلند",
  sek: "کرون سوئد",
  nok: "کرون نروژ",
  dkk: "کرون دانمارک",
  inr: "روپیه هند",
  krw: "وون کره جنوبی",
  myr: "رینگیت مالزی",
  thb: "بات تایلند",
  php: "پزوی فیلیپین",
  mxn: "پزو مکزیک",
  brl: "رئال برزیل",
  zar: "رند آفریقای جنوبی",
};

// -----------------------------
// Helpers
// -----------------------------
function nowMs() {
  return Date.now();
}

function safeString(v) {
  return typeof v === "string" ? v : "";
}

function num(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const cleaned = v.replace(/,/g, "").trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function isNumericLike(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).replace(/,/g, "").trim();
  if (!s) return false;
  return !Number.isNaN(Number(s));
}

function pickTs(item) {
  const ts =
    safeString(item?.dt) ||
    safeString(item?.ts) ||
    safeString(item?.date) ||
    safeString(item?.time) ||
    (typeof item?.last === "string" ? item.last : "") ||
    "";
  return ts.trim() !== "" ? ts.trim() : new Date().toISOString();
}

function baseCurrency(code) {
  const c = String(code).toLowerCase().trim();
  return c.split("_")[0] || c;
}

function tgjuKeyToCode(priceKey) {
  const raw = String(priceKey).replace(/^price_/, "").toLowerCase();
  return SPECIAL_CODE_MAP[raw] || raw;
}

// -----------------------------
// Classification rules
// -----------------------------
const CRYPTO_KEYWORDS = [
  "btc","eth","usdt","tether","xrp","trx","ltc","bch","bnb","ada",
  "doge","dot","sol","matic","shib","avax","atom","link","xlm","eos",
  "etc","omg","xaut","ton",
];

const GOLD_KEYWORDS = [
  "gold","silver","xau","sekke","rob","nim","gerami","emami","bahar",
  "mesghal","ons","coin","tala","sime","abshode",
];

function isGoldKey(key) {
  const k = key.toLowerCase();
  return GOLD_KEYWORDS.some((w) => k.includes(w));
}

function isCryptoKey(key) {
  const k = key.toLowerCase();
  if (k.endsWith("-irr") || k.endsWith("_irr")) {
    return CRYPTO_KEYWORDS.some((c) => k.startsWith(c));
  }
  if (k.startsWith("price_")) {
    const sym = k.slice("price_".length);
    return CRYPTO_KEYWORDS.includes(sym);
  }
  return CRYPTO_KEYWORDS.some(
    (c) => k === c || k.includes(`${c}-`) || k.includes(`${c}_`)
  );
}

function isFiatKey(key) {
  const k = key.toLowerCase();
  if (!k.startsWith("price_")) return false;
  if (isCryptoKey(k)) return false;
  if (isGoldKey(k)) return false;
  return true;
}

// Normalize TGJU entry -> row object (label/name NEVER numeric)
function normalizeEntry(priceKey, item, group) {
  const code = tgjuKeyToCode(priceKey);
  const base = baseCurrency(code);

  const price =
    num(item?.current) ?? num(item?.price) ?? num(item?.p) ?? num(item) ?? 0;

  const low =
    num(item?.tolerance_low) ?? num(item?.low) ?? num(item?.l) ?? null;

  const high =
    num(item?.tolerance_high) ?? num(item?.high) ?? num(item?.h) ?? null;

  const rawName =
    safeString(item?.name) ||
    safeString(item?.title) ||
    safeString(item?.n) ||
    "";

  const rawLabel =
    safeString(item?.label) ||
    safeString(item?.s) ||
    "";

  const safeRawName = isNumericLike(rawName) ? "" : rawName.trim();
  const safeRawLabel = isNumericLike(rawLabel) ? "" : rawLabel.trim();

  const fa = FA_NAME_MAP[code] || FA_NAME_MAP[base] || "";

  const name_fa = fa || safeRawName || safeRawLabel || code.toUpperCase();

  return {
    group,
    code,
    name_fa,
    price,
    change: safeString(item?.diff) || safeString(item?.change) || "0",
    low,
    high,
    ts: pickTs(item),
  };
}

// -----------------------------
// Google Sheets write
// -----------------------------
function loadServiceAccountFromEnv() {
  if (!SA_B64) {
    throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON in env");
  }
  let jsonText = "";
  try {
    jsonText = Buffer.from(SA_B64, "base64").toString("utf8");
  } catch {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not valid base64");
  }

  let creds;
  try {
    creds = JSON.parse(jsonText);
  } catch {
    throw new Error("Decoded service account JSON is invalid");
  }

  if (!creds.client_email || !creds.private_key) {
    throw new Error("Service account JSON missing client_email/private_key");
  }
  return creds;
}

async function getSheet() {
  if (!SHEET_ID) throw new Error("Missing SHEET_ID in env");

  const creds = loadServiceAccountFromEnv();

  const doc = new GoogleSpreadsheet(SHEET_ID);

  await doc.useServiceAccountAuth({
    client_email: creds.client_email,
    private_key: String(creds.private_key).replace(/\\n/g, "\n"),
  });

  await doc.loadInfo();

  const sheet = doc.sheetsByTitle[WORKSHEET_TITLE] || doc.sheetsByIndex[0];
  if (!sheet) throw new Error("Worksheet not found");

  return sheet;
}

async function writeRowsToSheet(rows) {
  const sheet = await getSheet();

  const wantedHeaders = [
    "group","code","name_fa","price","change","low","high","ts","updated_at"
  ];

  // Ensure headers
  try {
    await sheet.loadHeaderRow();
  } catch {
    // ignore
  }
  if (!sheet.headerValues || sheet.headerValues.length === 0) {
    await sheet.setHeaderRow(wantedHeaders);
  }

  // Delete old rows (keeps header)
  const existing = await sheet.getRows();
  if (existing.length) {
    await Promise.all(existing.map((r) => r.delete()));
  }

  const updated_at = new Date().toISOString();
  const finalRows = rows.map((r) => ({ ...r, updated_at }));

  if (finalRows.length) {
    await sheet.addRows(finalRows);
  }

  return { count: finalRows.length, updated_at };
}

// -----------------------------
// Fetch + build rows (with cache)
/// -----------------------------
let lastRun = { ok: false, error: "Not run yet", updated_at: null, count: 0 };
let lastFetchMs = 0;

async function fetchAndWrite(force = false) {
  const age = nowMs() - lastFetchMs;
  if (!force && lastRun.ok && age < CACHE_TTL_MS) return lastRun;

  const res = await fetch(TGJU_JSON_URL, {
    headers: {
      "User-Agent": "tgju-to-sheets/1.0",
      Accept: "application/json,text/plain,*/*",
      "Accept-Language": "fa,en;q=0.8",
    },
  });

  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error("TGJU response is not valid JSON");
  }

  const current = json?.current;
  if (!current || typeof current !== "object") {
    throw new Error("TGJU JSON missing 'current' object");
  }

  const rows = [];

  for (const key of Object.keys(current)) {
    if (!key.startsWith("price_")) continue;

    const item = current[key];
    const k = key.toLowerCase();

    if (isGoldKey(k)) {
      rows.push(normalizeEntry(key, item, "gold"));
      continue;
    }
    if (isCryptoKey(k)) {
      rows.push(normalizeEntry(key, item, "crypto"));
      continue;
    }
    if (isFiatKey(k)) {
      const entry = normalizeEntry(key, item, "fiat");
      if (ONLY_MAJOR && !MAJOR_FIAT.has(baseCurrency(entry.code))) continue;
      rows.push(entry);
      continue;
    }
  }

  // Sort for nicer output
  rows.sort((a, b) => (a.group + a.code).localeCompare(b.group + b.code));

  const result = await writeRowsToSheet(rows);

  lastRun = { ok: true, error: null, updated_at: result.updated_at, count: result.count };
  lastFetchMs = nowMs();

  return lastRun;
}

// -----------------------------
// Routes
// -----------------------------
app.get("/", (_req, res) => {
  res.type("text/plain").send("tgju-to-sheets running ✅");
});

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    service: "tgju-to-sheets",
    lastRun,
    only_major: ONLY_MAJOR,
    worksheet: WORKSHEET_TITLE,
  });
});

// Manual trigger
app.get("/run", async (req, res) => {
  const force = req.query.force === "1" || req.query.force === "true";
  try {
    const out = await fetchAndWrite(force);
    res.json({ ok: true, ...out });
  } catch (e) {
    res.status(500).json({ ok: false, error: e?.message || "unknown" });
  }
});

// -----------------------------
// Scheduler (every 5 minutes)
// -----------------------------
cron.schedule("*/5 * * * *", async () => {
  try {
    await fetchAndWrite(false);
    console.log("✅ Sheet updated:", lastRun);
  } catch (e) {
    console.error("❌ Update failed:", e?.message || e);
    lastRun = { ok: false, error: e?.message || "unknown", updated_at: null, count: 0 };
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`tgju-to-sheets running on port ${PORT}`);
  console.log(`Manual run: /run?force=1`);
});
