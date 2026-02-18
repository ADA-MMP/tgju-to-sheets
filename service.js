/**
 * service.js — Google Sheets sync (quota-safe, stable rows)
 *
 * Uses Google Sheets API directly (no google-spreadsheet).
 * Strategy (2 writes per run):
 *   1) Clear tab
 *   2) Update A1 with [header + all rows] in ONE request
 *
 * Rows are stable by: sort(group + ":" + code)
 *
 * Env required:
 *  SHEET_ID
 *  WORKSHEET_TITLE
 *  GOOGLE_SERVICE_ACCOUNT_JSON_BASE64
 */

import { GoogleAuth } from "google-auth-library";

const SHEET_ID = process.env.SHEET_ID;
const TAB = process.env.WORKSHEET_TITLE || "Rates";
const SA_B64 = process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 || "";

export const HEADERS = [
  "group",
  "code",
  "name_fa",
  "price",
  "change",
  "low",
  "high",
  "ts",
  "updated_at",
];

// -----------------------------
// Auth + HTTP helpers
// -----------------------------
function loadServiceAccountFromEnv() {
  if (!SA_B64) throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 in env");

  let jsonText;
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

async function getAccessToken() {
  const creds = loadServiceAccountFromEnv();

  const auth = new GoogleAuth({
    credentials: {
      client_email: creds.client_email,
      private_key: String(creds.private_key).replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const client = await auth.getClient();
  const tokenResponse = await client.getAccessToken();
  const token = typeof tokenResponse === "string" ? tokenResponse : tokenResponse?.token;

  if (!token) throw new Error("Could not obtain Google access token");
  return token;
}

async function gfetch(url, { method = "GET", token, body } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }

  if (!res.ok) {
    const msg =
      json?.error?.message ||
      json?.message ||
      `${res.status} ${res.statusText}` ||
      "Unknown Google API error";
    throw new Error(`Google API error - [${res.status}] ${msg}`);
  }

  return json;
}

// -----------------------------
// Core: stable, bulk sync
// -----------------------------
function stableSortRows(rows) {
  return [...rows].sort((a, b) => {
    const ak = `${a.group || ""}:${a.code || ""}`.toLowerCase();
    const bk = `${b.group || ""}:${b.code || ""}`.toLowerCase();
    return ak.localeCompare(bk);
  });
}

function normalizeRow(r, updated_at) {
  return [
    String(r.group ?? ""),
    String(r.code ?? ""),
    String(r.name_fa ?? ""),
    r.price ?? "",
    String(r.change ?? ""),
    r.low ?? "",
    r.high ?? "",
    String(r.ts ?? ""),
    updated_at,
  ];
}

/**
 * syncToSheet(rows)
 * - rows: array of objects: {group, code, name_fa, price, change, low, high, ts}
 * - stable by group+code (sorting)
 * - bulk clear + bulk update
 */
export async function syncToSheet(rows) {
  if (!SHEET_ID) throw new Error("Missing SHEET_ID in env");
  if (!TAB) throw new Error("Missing WORKSHEET_TITLE in env");

  const token = await getAccessToken();
  const updated_at = new Date().toISOString();

  const sorted = stableSortRows(rows);

  // Build 2D values array: header + rows
  const values = [
    HEADERS,
    ...sorted.map((r) => normalizeRow(r, updated_at)),
  ];

  // 1) CLEAR entire tab (1 write)
  // Note: clearing by tab name clears used cells in that sheet.
  await gfetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      SHEET_ID
    )}/values/${encodeURIComponent(TAB)}:clear`,
    { method: "POST", token, body: {} }
  );

  // 2) UPDATE all values starting at A1 (1 write)
  await gfetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      SHEET_ID
    )}/values/${encodeURIComponent(TAB)}!A1?valueInputOption=RAW`,
    {
      method: "PUT",
      token,
      body: { range: `${TAB}!A1`, majorDimension: "ROWS", values },
    }
  );

  return { count: sorted.length, updated_at };
}
