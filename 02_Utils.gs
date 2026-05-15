// ============================================================
// 02_Utils.gs
// Shared helpers: sheet I/O, caching, time/ID generators, phone
// normalisation, PIN matching, schema init.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── CONSTANT-TIME PIN COMPARISON ─────────────────────────────
// Prevents timing-oracle attacks where comparing a wrong-length PIN
// returns faster than a correct-length one, leaking PIN length.
// Pads both sides to 32 chars, XORs every character, checks all at once.
function _pinMatch(supplied, expected) {
  const a = String(supplied || "").padEnd(32, "\0");
  const b = String(expected  || "").padEnd(32, "\0");
  let diff = 0;
  for (let i = 0; i < 32; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  // Also require lengths match (padEnd would equalize lengths, so check originals)
  return diff === 0 && String(supplied).length === String(expected).length;
}
function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
// ── GOOGLE ANALYTICS 4 INTEGRATION ──────────────────────────
const GA4_HEADERS = ["Date", "Source", "Device", "Active_Users", "Sessions", "Page_Views", "Engagement_Rate", "Avg_Session_Duration", "Event_Count"];
// ── HELPERS ──────────────────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}
function getOrCreateTab(ss, name, headers) {
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
  }
  
  if (headers && headers.length > 0) {
    const lastCol = ws.getLastColumn();
    const currentHeaders = lastCol > 0 ? ws.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||"").trim()) : [];
    
    // Force header row synchronization by explicitly setting range if any mismatch
    headers.forEach((h, i) => {
      if (currentHeaders[i] !== h) {
        ws.getRange(1, i + 1).setValue(h)
          .setFontWeight("bold")
          .setBackground("#c0392b")
          .setFontColor("white");
        if (i === 0) ws.setFrozenRows(1);
        // Force certain columns to stay as Plain Text to preserve leading zeros
        if (h === "Phone" || h === "PIN") {
          ws.getRange(1, i + 1, ws.getMaxRows(), 1).setNumberFormat("@");
        }
      }
    });

    // CRITICAL: If headers were provided, ensure No Extra Columns exist beyond them
    // This prevents "Timestamp" duplicates if things drifted in legacy versions
    if (headers.length > 0 && ws.getLastColumn() > headers.length) {
      const extra = ws.getLastColumn() - headers.length;
      ws.deleteColumns(headers.length + 1, extra);
    }
  }
  return ws;
}
// ── CACHE HELPER ────────────────────────────────────────────
// Cross-execution cache using Apps Script CacheService.
// Falls back gracefully if value is too large (>100 KB) to store.
function _cachedData(key, ttlSeconds, fetchFn) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit !== null) {
    try { return JSON.parse(hit); } catch(e) {}
  }
  const data = fetchFn();
  try { cache.put(key, JSON.stringify(data), ttlSeconds); } catch(e) {
    // Value may exceed 100 KB limit — silent fallback to uncached
  }
  return data;
}
function _invalidateCache() {
  const keys = Array.from(arguments);
  if (!keys.length) return;
  try { CacheService.getScriptCache().removeAll(keys); } catch(e) {}
}
function getISTDate() {
  const now = new Date();
  // Cross-environment IST Date object
  return new Date(now.getTime() + (now.getTimezoneOffset() + 330) * 60000);
}
function getISTTimestamp() {
  return Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}
function generateSubmissionID() {
  const ist = getISTDate();
  const dateStr = Utilities.formatDate(ist, "Asia/Kolkata", "yyyyMMdd");
  const rand = Math.floor(Math.random() * 9000) + 1000;
  return `SK-${dateStr}-${rand}`;
}
function headerIndex(ws) {
  // Returns {colName: 1-based-index} for the given sheet
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i + 1; });
  return idx;
}
function getAllRows(ws) {
  const last = ws.getLastRow();
  if (last < 2) return [];
  const data = ws.getRange(2, 1, last - 1, ws.getLastColumn()).getValues();
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  return data.map((row, ri) => {
    const obj = {_row: ri + 2};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
function _get(obj, key) {
  if (!obj || !key) return undefined;
  if (obj[key] !== undefined) return obj[key];
  const nk = key.replace(/_/g, ' ').toLowerCase();
  for (let k in obj) {
    if (k.replace(/_/g, ' ').toLowerCase() === nk) return obj[k];
  }
  return undefined;
}
function _cleanNum(val) {
  if (typeof val === "number") return val;
  const s = String(val || "").replace(/[^\d.-]/g, '');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
function getRecentRows(ws, maxRows) {
  const last = ws.getLastRow();
  if (last < 2) return [];
  const startRow = Math.max(2, last - maxRows + 1);
  const data = ws.getRange(startRow, 1, last - startRow + 1, ws.getLastColumn()).getValues();
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  return data.map((row, ri) => {
    const obj = {_row: ri + startRow};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
// ── SCHEMA INIT ──────────────────────────────────────────────
function initSchema() {
  const ss = getSpreadsheet();
  getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner",
    "OOS_JSON","Orders_Closed","Stock_JSON"
  ]);
  getOrCreateTab(ss, TAB_BF_MASTER, ["ID","Name","Price","Active"]);
  getOrCreateTab(ss, TAB_SABJI,     ["ID","Name","Type","Active"]);
  getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  return {success: true, message: "Schema initialised"};
}
/**
 * Normalizes phone numbers for reliable comparison across Google Sheets.
 * Handles scientific notation (e.g., 9.87E+9) and trailing decimals (.0).
 */
// Returns true if the order should be excluded from kitchen/prep counts.
// "Cancelled (Verify UPI)" = soft-cancel pending admin verification —
// the customer already requested cancellation, do NOT include in kitchen prep.
function _isOrderCancelled(paymentStatus) {
  const s = String(paymentStatus || "").toLowerCase();
  return s === "cancelled" || s.startsWith("cancelled");
}
function _normalizePhone(phone) {
  let p = String(phone || "").trim();
  if (!p) return "";
  // Scientific notation (Sheets quirk: 9.87654321e+9)
  if (p.toUpperCase().includes("E+") && !isNaN(Number(p))) {
    p = String(Math.round(Number(p)));
  }
  // Trailing decimal from Sheets (9876543210.0)
  if (p.includes(".")) p = p.split(".")[0];
  // Strip everything that isn't a digit (removes +, spaces, dashes, parens, country-code prefixes)
  p = p.replace(/\D/g, "");
  // 12-digit with 91 country code → 10-digit
  if (p.length === 12 && p.startsWith("91")) p = p.substring(2);
  // 11-digit with leading zero → 10-digit
  if (p.length === 11 && p.startsWith("0")) p = p.substring(1);
  return p;
}
