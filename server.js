/**
 * service.js — TGJU → Google Sheets (ESM, Option A, Render-friendly)
 *
 * Fixes your error:
 *   "doc.useServiceAccountAuth is not a function"
 * by using google-auth-library JWT and passing it into GoogleSpreadsheet(...)
 *
 * Key requirement:
 *   Keep rows “fixed” by (group + code):
 *   - If a row with same group+code exists -> UPDATE it in place
 *   - If not -> ADD a new row
 *
 * ENV (Render > Environment Variables):
 *   SHEET_ID=xxxxxxxxxxxxxxxxxxxxxxxxxxxx
 *   WORKSHEET_TITLE=Rates
 *   GOOGLE_SERVICE_ACCOUNT_JSON_BASE64=base64(JSON)
 *   ONLY_MAJOR=0|1
 *   CACHE_TTL_MS=60000
 *   PORT=3000  (Render sets PORT automatically; you can omit)
 *
 * NPM deps needed:
 *   express, dotenv, node-cron, google-spreadsheet, google-auth-library
 */

import "dotenv/config";
import express from "express";
import cron from "node-cron";
import { GoogleSpreadsheet } from "google-spreadsheet";
import { JWT } from "google-auth-library";

const app = express();

const PORT = Number(process.env.PORT || 3000);
const TGJU_JSON_URL = "https://call2.tgju.org/ajax.json";

const SHEET_ID = process.env.SHEET_ID || "";
const WORKSHEET_TITLE = process.env.WORKSHEET_TITLE || "Rates";
const CACHE_TTL_MS = Number(process.env.CACHE_TTL_MS || 60_000);
const ONLY_MAJOR = String(process.env.ONLY_MAJOR || "0") === "1";

const SA_B64 = process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 || "";

// -----------------------------
// Major fiat only (optional)
// -----------------------------
const MAJOR_FIAT = new Set([
  "usd", "eur", "gbp", "cad", "aed", "try",
  "sar", "qar", "kwd", "bhd", "iqd",
  "cny", "jpy", "chf", "rub",
  "aud", "nzd", "sek", "nok", "dkk",
  "inr", "krw", "myr", "thb",
  "php", "mxn", "brl", "zar",
]);

// TGJU special keys -> normalized codes
const SPECIAL_CODE_MAP = {
  dollar_rl: "usd",
  dollar_ex: "usd_official",
  dollar_dt: "usd_dt",
  dollar_sm: "usd_sm",
  eur_ex: "eur_official",
};

// Persian names (extend anytime)
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
  "gold","silver","xau","sekke","sekee","sekeb","rob","nim","gerami",
  "emami","bahar","mesghal","ons","coin","tala","sime","abshode",
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

// -----------------------------
// Normalize TGJU entry -> sheet row
// -----------------------------
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

  // NEVER allow numeric “name/label”
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
// Google auth + sheet helpers
// -----------------------------
function loadServiceAccountFromEnv() {
  if (!SA_B64) throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 in env");

  let jsonText = "";
  try {
    jsonText = Buffer.from(SA_B64, "base64").toString("utf8");
  } catch {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 is not valid base64");
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

async function getWorksheet() {
  if (!SHEET_ID) throw new Error("Missing SHEET_ID in env");

  const creds = loadServiceAccountFromEnv();
  const jwt = new JWT({
    email: creds.client_email,
    key: String(creds.private_key).replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  // ✅ v4+ uses auth in constructor
  const doc = new GoogleSpreadsheet(SHEET_ID, jwt);

  await doc.loadInfo();

  const sheet = doc.sheetsByTitle[WORKSHEET_TITLE] || doc.sheetsByIndex[0];
  if (!sheet) throw new Error("Worksheet not found");

  return sheet;
}

async function ensureHeaders(sheet) {
  const wanted = ["group","code","name_fa","price","change","low","high","ts","updated_at"];

  // Important: load header row first (fixes "Header values are not yet loaded")
  try {
    await sheet.loadHeaderRow();
  } catch {
    // If sheet is brand-new, loadHeaderRow may fail; ignore.
  }

  const current = Array.isArray(sheet.headerValues) ? sheet.headerValues : [];

  if (!current.length) {
    await sheet.setHeaderRow(wanted);
    return wanted;
  }

  // If headers differ, set to wanted (optional but safer)
  const same =
    current.length === wanted.length &&
    current.every((h, i) => String(h) === String(wanted[i]));

  if (!same) {
    await sheet.setHeaderRow(wanted);
    return wanted;
  }

  return current;
}

/**
 * Upsert by group+code:
 * - Update existing row in place
 * - Add new row if not found
 * (Does NOT delete rows that disappeared from TGJU.)
 */
async function upsertRows(sheet, incomingRows) {
  await ensureHeaders(sheet);

  const existing = await sheet.getRows(); // reads all rows
  const byKey = new Map();

  for (const r of existing) {
    const g = String(r.get("group") ?? "").trim().toLowerCase();
    const c = String(r.get("code") ?? "").trim().toLowerCase();
    if (!g || !c) continue;
    byKey.set(`${g}|${c}`, r);
  }

  const updated_at = new Date().toISOString();
  let updated = 0;
  const toInsert = [];

  for (const row of incomingRows) {
    const g = String(row.group).toLowerCase();
    const c = String(row.code).toLowerCase();
    const key = `${g}|${c}`;

    const found = byKey.get(key);
    if (found) {
      // update in place
      found.set("name_fa", row.name_fa);
      found.set("price", row.price);
      found.set("change", row.change);
      found.set("low", row.low);
      found.set("high", row.high);
      found.set("ts", row.ts);
      found.set("updated_at", updated_at);

      await found.save();
      updated++;
    } else {
      // insert new
      toInsert.push({
        ...row,
        updated_at,
      });
    }
  }

  let inserted = 0;
  if (toInsert.length) {
    await sheet.addRows(toInsert);
    inserted = toInsert.length;
  }

  return { updated, inserted, updated_at, total_incoming: incomingRows.length };
}

// -----------------------------
// Fetch TGJU -> build incoming rows
// -----------------------------
let lastRun = {
  ok: false,
  error: "Not run yet",
  updated_at: null,
  total_incoming: 0,
  updated: 0,
  inserted: 0,
};
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
    }
  }

  // Stable sort (doesn’t affect row “fixedness”; key is group+code)
  rows.sort((a, b) => (a.group + "|" + a.code).localeCompare(b.group + "|" + b.code));

  const sheet = await getWorksheet();
  const result = await upsertRows(sheet, rows);

  lastRun = { ok: true, error: null, ...result };
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
    worksheet: WORKSHEET_TITLE,
    only_major: ONLY_MAJOR,
    cache_ttl_ms: CACHE_TTL_MS,
    lastRun,
  });
});

app.get("/run", async (req, res) => {
  const force = req.query.force === "1" || req.query.force === "true";
  try {
    const out = await fetchAndWrite(force);
    res.json({ ok: true, ...out });
  } catch (e) {
    lastRun = { ok: false, error: e?.message || "unknown", updated_at: null, total_incoming: 0, updated: 0, inserted: 0 };
    res.status(500).json({ ok: false, error: e?.message || "unknown" });
  }
});

// -----------------------------
// Scheduler (every 5 minutes)
// -----------------------------
cron.schedule("*/5 * * * *", async () => {
  try {
    await fetchAndWrite(false);
    console.log("✅ Sheet upsert OK:", lastRun);
  } catch (e) {
    console.error("❌ Sheet upsert failed:", e?.message || e);
    lastRun = { ok: false, error: e?.message || "unknown", updated_at: null, total_incoming: 0, updated: 0, inserted: 0 };
  }
});

// Start
app.listen(PORT, () => {
  console.log(`tgju-to-sheets listening on :${PORT}`);
  console.log(`Manual run: /run?force=1`);
});
