// ============================================================
// 04_Menu_Admin.gs
// Daily menu, breakfast/sabji masters, free areas, admin dashboard data,
// stock-limit helpers.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── STOCK LIMIT HELPERS ─────────────────────────────────────
// Map admin-stock colKey to the name used in Items_JSON.
// Breakfast Curd is stored as "Breakfast Curd" (new rows) — old rows stored it
// as plain "Curd". countOrderedUnits handles both for backward compat.
function itemsJsonKey(colKey) { return colKey === "B_CURD" ? "Breakfast Curd" : colKey; }
// Count ordered units per meal/item for a given date, excluding cancelled orders.
function countOrderedUnits(ordersRows, dateStr) {
  const counts = { Breakfast: {}, Lunch: {}, Dinner: {} };
  ordersRows.forEach(row => {
    if (_isOrderCancelled(row.Payment_Status)) return;
    const d = row.Order_Date instanceof Date
      ? Utilities.formatDate(row.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(row.Order_Date || "").trim();
    if (d !== dateStr) return;
    const meal = String(row.Meal_Type || "");
    if (!counts[meal]) return;
    let items = {};
    try { items = JSON.parse(row.Items_JSON || "{}"); } catch(e) {}
    Object.entries(items).forEach(([name, qty]) => {
      // Backward compat: old Breakfast rows stored Curd as "Curd". Normalize to
      // "Breakfast Curd" so aggregates match the new canonical key.
      let k = name;
      if (meal === "Breakfast" && name === "Curd") k = "Breakfast Curd";
      counts[meal][k] = (counts[meal][k] || 0) + Number(qty || 0);
    });
  });
  return counts;
}
// ── GET MENU ─────────────────────────────────────────────────
// Count ACTIVE (non-cancelled) orders per meal type for one date, from a rows
// array. One order row = one order. Cancelled rows free their slot. Shared by
// getMenu (display) and the submitOrder cap guard (authoritative).
function _countActiveMealOrders(rows, dateStr) {
  const c = { Breakfast: 0, Lunch: 0, Dinner: 0 };
  for (var i = 0; i < rows.length; i++) {
    const r = rows[i];
    const d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date || "").trim();
    if (d !== dateStr) continue;
    if (_isOrderCancelled(r.Payment_Status)) continue;
    // The cap is a DELIVERY limit — Self Pickup / Porter orders don't use a
    // delivery slot, so they neither count toward the cap nor get blocked by it.
    const ar = String(r.Area || "").toLowerCase();
    if (ar.indexOf("pickup") !== -1 || ar === "porter") continue;
    const mt = String(r.Meal_Type || "").trim();
    if (c[mt] !== undefined) c[mt]++;
  }
  return c;
}
function getMenu(dateStr) {
  // Cache per-date for 60 s. The hard stock-block in submitOrder (under LockService)
  // prevents actual over-orders even when menu data is slightly stale.
  return _cachedData("menu_v2_" + dateStr, 60, function() { return _getMenuUncached(dateStr); });
}
function getMenuBatch(datesStr) {
  const dates = String(datesStr || "").split(',').map(d => d.trim()).filter(Boolean);
  const result = {};
  dates.forEach(d => {
    // Rely on the existing cached helper so we don't duplicate logic
    result[d] = getMenu(d);
  });
  return result;
}
// For the REORDER flow: given a comma-separated list of breakfast item names,
// returns whether the calendar should be restricted and to which dates.
//  - Everyday items (master Active, e.g. Poha/Upma) impose NO restriction.
//  - Special items (e.g. Aloo Paratha) restrict to upcoming dates whose daily
//    Breakfast_JSON includes them (intersection if several specials).
// Returns { restrict: bool, dates: ["yyyy-MM-dd", ...] }.
function getBreakfastItemDates(itemsStr) {
  const items = String(itemsStr || "").split(',').map(s => s.trim()).filter(Boolean);
  if (!items.length) return { restrict: false, dates: [] };
  const _norm = function(n){ return String(n||"").toLowerCase().replace(/\[[^\]]*\]/g,"").replace(/\([^)]*\)/g,"").replace(/\s+/g," ").trim(); };
  const ss = getSpreadsheet();

  // Everyday (master-Active) breakfast item names — these are on EVERY day.
  const bfWs = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const activeSet = new Set(getAllRows(bfWs).filter(function(x){ return String(x.Active).toLowerCase() !== "false"; }).map(function(x){ return _norm(x.Name); }));

  const specials = items.map(_norm).filter(function(n){ return n && !activeSet.has(n); });
  if (!specials.length) return { restrict: false, dates: [] }; // all everyday → no restriction

  const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const menuWs = getOrCreateTab(ss, TAB_MENU, []);
  const datesByItem = {};
  specials.forEach(function(s){ datesByItem[s] = {}; });
  getAllRows(menuWs).forEach(function(r){
    const d = r.Date instanceof Date ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd") : String(r.Date||"").trim();
    if (!d || d < today || !r.Breakfast_JSON) return;
    let parsed; try { parsed = JSON.parse(r.Breakfast_JSON); } catch(e) { return; }
    if (!Array.isArray(parsed)) return;
    const namesOnDay = {};
    parsed.forEach(function(x){ namesOnDay[_norm(x && x.name)] = true; });
    specials.forEach(function(s){ if (namesOnDay[s]) datesByItem[s][d] = true; });
  });

  // allowed = intersection of every special item's date set
  let allowed = null;
  specials.forEach(function(s){
    const ks = Object.keys(datesByItem[s]);
    if (allowed === null) allowed = ks;
    else allowed = allowed.filter(function(x){ return datesByItem[s][x]; });
  });
  return { restrict: true, dates: (allowed || []).sort() };
}
function _getMenuUncached(dateStr) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => {
    const d = x.Date instanceof Date
      ? Utilities.formatDate(x.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(x.Date).trim();
    return d === dateStr;
  });

  // Admin can mark a specific (non-Sunday) day as Kitchen Closed via the
  // Daily Menu tab. When set, customer calendar greys out the day and any
  // submitOrder attempt for it is rejected server-side.
  const _kitchenClosed = !!(r && (r.Kitchen_Closed === true ||
    String(r.Kitchen_Closed || "").toLowerCase() === "true"));

  // Breakfast master items
  const bfWs = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const bfRows = getAllRows(bfWs).filter(x => String(x.Active).toLowerCase() !== "false");
  
  const NAME_MAP = {
    "Kanda Poha": "Kanda Poha [175g]",
    "Ghee Upma": "Ghee Upma [200g]",
    "Sabudana Khichdi": "Sabudana Khichdi [200g]",
    "Tikhi Pudi": "5 x Tikhi Pudi with 100 ml coriander chutney",
    "Tikhi Puri": "5 x Tikhi Pudi with 100 ml coriander chutney",
    "Idli Chutney": "4 x Idli & 100ml Chutney",
    "Idli": "4 x Idli & 100ml Chutney",
    "4 x Idli & 100ml Chutney": "4 x Idli & 100ml Chutney",
    "Ghee Sheera": "Ghee Sheera [200g]"
  };

  const breakfast = bfRows.map(x => {
    const rawName = String(x.Name).trim();
    return {
      name: NAME_MAP[rawName] || rawName,
      price: Number(x.Price)
    };
  });

  // Determine if this date is a Sunday with no sabjis set.
  // Kitchen is closed on Sundays by default; admin can override by setting at least one sabji.
  const _dayName   = Utilities.formatDate(new Date(dateStr + "T12:00:00+05:30"), "Asia/Kolkata", "EEEE");
  const _isSunday  = _dayName === "Sunday";
  const _hasSabjis = r && (r.Lunch_Dry || r.Lunch_Curry || r.Dinner_Dry || r.Dinner_Curry);

  if (!r) {
    // No menu row at all — if Sunday, close everything; otherwise return open empty menu.
    return {
      breakfast, lunch_dry:"", lunch_curry:"", dinner_dry:"", dinner_curry:"",
      cutoff_overrides:{},
      oos_items: { Breakfast: [], Lunch: [], Dinner: [] },
      orders_closed: _isSunday ? { Breakfast: true, Lunch: true, Dinner: true } : {},
      stock_limits: {},
      units_remaining: {},
      sunday_closed: _isSunday
    };
  }

  // Menu row exists but it's a Sunday with no sabjis — still treat as closed.
  if (_isSunday && !_hasSabjis) {
    let ordersClosed2 = { Breakfast: true, Lunch: true, Dinner: true };
    return {
      breakfast, lunch_dry:"", lunch_curry:"", dinner_dry:"", dinner_curry:"",
      cutoff_overrides:{},
      oos_items: { Breakfast: [], Lunch: [], Dinner: [] },
      orders_closed: ordersClosed2,
      stock_limits: {},
      units_remaining: {},
      sunday_closed: true
    };
  }

  const co = {};
  if (r && r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
  if (r && r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
  if (r && r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);

  // MERGE LOGIC: Start with master active items, then merge daily overrides
  const masterActive = breakfast;
  let dailyBf = [];
  if (r && r.Breakfast_JSON) {
    try { 
      const parsed = JSON.parse(r.Breakfast_JSON); 
      dailyBf = parsed.map(d => ({
        ...d,
        name: d.name ? (NAME_MAP[d.name.trim()] || d.name) : ""
      }));
    } catch(e) {}
  }

  // Prioritize Daily selections (where specific prices or choices were made)
  // but ensure Master Active items are always present.
  const finalBreakfast = [...dailyBf];
  masterActive.forEach(m => {
    if (!finalBreakfast.some(d => d.name === m.name)) {
      finalBreakfast.push(m);
    }
  });

  let oosItems = { Breakfast: [], Lunch: [], Dinner: [] };
  try { if (r && r.OOS_JSON) oosItems = JSON.parse(r.OOS_JSON); } catch(e) {}

  let ordersClosed = {};
  try { if (r && r.Orders_Closed) ordersClosed = JSON.parse(r.Orders_Closed); } catch(e) {}

  let stockLimits = {};
  try { if (r && r.Stock_JSON) stockLimits = JSON.parse(r.Stock_JSON); } catch(e) {}

  // Per-meal max-order caps (e.g. {"Breakfast":50}). When a meal's active
  // (non-cancelled) order count reaches its cap it is SOLD OUT for the day.
  let orderCaps = {};
  try { if (r && r.Order_Cap_JSON) orderCaps = JSON.parse(r.Order_Cap_JSON); } catch(e) {}

  const ordersWs2   = getOrCreateTab(ss, TAB_ORDERS, []);
  // OPTIMIZATION: Only read the last 500 rows to compute stock limit (covers today and yesterday).
  // This prevents scanning thousands of old orders just to check today's stock.
  const ordersRows2 = getRecentRows(ordersWs2, 500);
  const orderedCounts = countOrderedUnits(ordersRows2, dateStr);
  const unitsRemaining = {};
  ["Breakfast","Lunch","Dinner"].forEach(meal => {
    Object.entries(stockLimits[meal] || {}).forEach(([colKey, limit]) => {
      if (!unitsRemaining[meal]) unitsRemaining[meal] = {};
      unitsRemaining[meal][colKey] = Math.max(0, limit - (orderedCounts[meal][itemsJsonKey(colKey)] || 0));
    });
  });

  // Cap evaluation — reuse the rows already read above (zero extra cost). The
  // cap is a DELIVERY limit: when reached we flag the meal sold_out so the order
  // page offers Self Pickup / Porter (which bypass the cap). We do NOT set
  // orders_closed — that path stays open. submitOrder is the authoritative guard:
  // it rejects DELIVERY orders past the cap (full-sheet count, under lock) while
  // letting Self Pickup / Porter through.
  const orderCounts = _countActiveMealOrders(ordersRows2, dateStr);
  const soldOut = {};
  ["Breakfast","Lunch","Dinner"].forEach(meal => {
    const cap = Number(orderCaps[meal] || 0);
    if (cap > 0 && (orderCounts[meal] || 0) >= cap) soldOut[meal] = true;
  });

  return {
    breakfast:    finalBreakfast,
    lunch_dry:    r ? (r.Lunch_Dry || "") : "",
    lunch_curry:  r ? (r.Lunch_Curry || "") : "",
    dinner_dry:   r ? (r.Dinner_Dry || "") : "",
    dinner_curry: r ? (r.Dinner_Curry || "") : "",
    cutoff_overrides: co,
    oos_items:    oosItems,
    orders_closed: ordersClosed,
    stock_limits: stockLimits,
    units_remaining: unitsRemaining,
    order_caps:    orderCaps,    // admin display: configured per-meal max
    order_counts:  orderCounts,  // admin display: active orders placed so far
    sold_out:      soldOut,      // customer display: meal hit its cap today
    kitchen_closed: _kitchenClosed
  };
}

// ── KITCHEN CLOSURE: list of admin-closed (non-Sunday) dates ─────
// Lightweight endpoint used by the customer calendar to grey out
// closed days without having to fetch every date's full menu.
function getKitchenClosedDates() {
  return _cachedData("kitchen_closed_dates_v1", 60, function() {
    const ss   = getSpreadsheet();
    const ws   = getOrCreateTab(ss, TAB_MENU, []);
    const rows = getAllRows(ws);
    const today = getISTDate();
    // Include the recent past (40 days) too — the loyalty streak looks backward
    // and must skip admin days-off so they don't break a customer's streak.
    const cutoff = Utilities.formatDate(new Date(Date.now() - 40 * 86400000), "Asia/Kolkata", "yyyy-MM-dd");
    const closed = [];
    rows.forEach(function(r) {
      const isClosed = (r.Kitchen_Closed === true ||
        String(r.Kitchen_Closed || "").toLowerCase() === "true");
      if (!isClosed) return;
      const d = r.Date instanceof Date
        ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd")
        : String(r.Date).trim();
      if (!d || d < cutoff) return;
      closed.push(d);
    });
    closed.sort();
    return { closedDates: closed };
  });
}

// ── GET WEEKLY MENU (next 7 days) ────────────────────────────
function getWeeklyMenu() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, []);
  const rows = getAllRows(ws);

  // Breakfast master items
  const bfWs = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const bfRows = getAllRows(bfWs).filter(x => String(x.Active).toLowerCase() !== "false");
  const defaultBreakfast = bfRows.map(x => ({name: String(x.Name), price: Number(x.Price)}));

  // Build a map: dateStr → row for quick lookup
  const menuMap = {};
  rows.forEach(x => {
    const d = x.Date instanceof Date
      ? Utilities.formatDate(x.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(x.Date).trim();
    menuMap[d] = x;
  });

  // Show all dates from today onwards that have a menu row set
  const today = getISTDate();
  const todayStr = Utilities.formatDate(today, "Asia/Kolkata", "yyyy-MM-dd");

  // Collect all future/today dates that have a menu row, sorted ascending
  const futureDates = Object.keys(menuMap)
    .filter(d => d >= todayStr)
    .sort();

  const days = [];
  futureDates.forEach(dateStr => {
    const d = new Date(dateStr + "T00:00:00+05:30");
    const dayName    = Utilities.formatDate(d, "Asia/Kolkata", "EEEE");
    const displayDate = Utilities.formatDate(d, "Asia/Kolkata", "dd MMM");

    const r = menuMap[dateStr];

    // Skip Sundays that have no sabjis set — kitchen is closed by default on Sundays.
    // A Sunday only appears in the weekly menu popup if the admin has explicitly
    // set at least one sabji (Lunch or Dinner), signalling the kitchen is open that day.
    const isSunday = dayName === "Sunday";
    const hasSabjis = r && (r.Lunch_Dry || r.Lunch_Curry || r.Dinner_Dry || r.Dinner_Curry);
    if (isSunday && !hasSabjis) return;

    let bfDaily = [];
    try {
      if (r && r.Breakfast_JSON) bfDaily = JSON.parse(r.Breakfast_JSON);
    } catch(e) {}

    // Merge Master + Daily
    const finalBf = [...bfDaily];
    defaultBreakfast.forEach(m => {
      if (!finalBf.some(x => x.name === m.name)) finalBf.push(m);
    });

    days.push({
      date: dateStr,
      dayName: dayName,
      displayDate: displayDate,
      breakfast: finalBf,
      lunch_dry:    r ? (r.Lunch_Dry    || "") : "",
      lunch_curry:  r ? (r.Lunch_Curry  || "") : "",
      dinner_dry:   r ? (r.Dinner_Dry   || "") : "",
      dinner_curry: r ? (r.Dinner_Curry || "") : "",
      menuSet: true  // only dates with a menu row are included
    });
  });

  return { success: true, days: days };
}
// ── ADMIN: GET ALL DATA ──────────────────────────────────────
function getAdminData() {
  return _cachedData("adminData_v1", 30, _getAdminDataUncached);
}

// Lightweight id→name map for the breakfast + sabji masters ONLY.
// submitOrder uses this to resolve item-id columns; it must NOT call the full
// getAdminData(), whose menuEntries pass scans every order row per menu date
// (O(orders × dates)) and took ~40s cold — the entire order-placement lag.
// Two small sheet reads, cached 5 min (masters change rarely).
function _getMastersMap() {
  return _cachedData("mastersMap_v1", 300, function() {
    const ss = getSpreadsheet();
    const map = {};
    getAllRows(getOrCreateTab(ss, TAB_BF_MASTER, [])).forEach(function(r) {
      if (r.ID !== "" && r.ID !== undefined) map[String(r.ID)] = String(r.Name || "");
    });
    getAllRows(getOrCreateTab(ss, TAB_SABJI, [])).forEach(function(r) {
      if (r.ID !== "" && r.ID !== undefined) map[String(r.ID)] = String(r.Name || "");
    });
    return map;
  });
}

function _getAdminDataUncached() {
  const ss = getSpreadsheet();

  const bfWs   = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const sabjiWs= getOrCreateTab(ss, TAB_SABJI,     []);
  const menuWs = getOrCreateTab(ss, TAB_MENU,       []);

  const bfRows    = getAllRows(bfWs);
  const sabjiRows = getAllRows(sabjiWs);
  const menuRows  = getAllRows(menuWs);

  const ordersWsAdm = getOrCreateTab(ss, TAB_ORDERS, []);
  const allOrdersAdm = getAllRows(ordersWsAdm);

  // Build per-date ordered-unit counts in ONE pass over all orders. Previously
  // countOrderedUnits(allOrders, date) was called per menu row → O(orders ×
  // dates) with a JSON.parse for every order each time (the ~40s hot spot).
  const countsByDate = {};
  const mealOrderCounts = {};   // dd → {Breakfast,Lunch,Dinner} active order-row counts (for the per-meal order cap)
  allOrdersAdm.forEach(function(row) {
    if (_isOrderCancelled(row.Payment_Status)) return;
    const dd = row.Order_Date instanceof Date
      ? Utilities.formatDate(row.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(row.Order_Date || "").trim();
    if (!dd) return;
    const meal = String(row.Meal_Type || "");
    if (!mealOrderCounts[dd]) mealOrderCounts[dd] = { Breakfast: 0, Lunch: 0, Dinner: 0 };
    if (mealOrderCounts[dd][meal] !== undefined) mealOrderCounts[dd][meal]++;
    if (!countsByDate[dd]) countsByDate[dd] = { Breakfast: {}, Lunch: {}, Dinner: {} };
    if (!countsByDate[dd][meal]) return;
    let items = {};
    try { items = JSON.parse(row.Items_JSON || "{}"); } catch (e) {}
    Object.entries(items).forEach(function(pair) {
      let k = pair[0];
      if (meal === "Breakfast" && k === "Curd") k = "Breakfast Curd";
      countsByDate[dd][meal][k] = (countsByDate[dd][meal][k] || 0) + Number(pair[1] || 0);
    });
  });

  const breakfastMaster = bfRows.map(r => ({
    id: String(r.ID), name: String(r.Name), price: Number(r.Price),
    default_on: r.Active === true || String(r.Active).toUpperCase() === "TRUE"
  }));

  const sabjiMaster = sabjiRows.map(r => ({
    id: String(r.ID), name: String(r.Name), type: String(r.Type),
    active: String(r.Active).toLowerCase() !== "false"
  }));

  const menuEntries = menuRows.map(r => {
    const d = r.Date instanceof Date
      ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Date).trim();
    const co = {};
    if (r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
    if (r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
    if (r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);
    let breakfast = [];
    try { if (r.Breakfast_JSON) breakfast = JSON.parse(r.Breakfast_JSON); } catch(e) {}
    let oosItems = { Breakfast: [], Lunch: [], Dinner: [] };
    try { if (r.OOS_JSON) oosItems = JSON.parse(r.OOS_JSON); } catch(e) {}
    let ordersClosed = {};
    try { if (r.Orders_Closed) ordersClosed = JSON.parse(r.Orders_Closed); } catch(e) {}
    let stockLimits = {};
    try { if (r.Stock_JSON) stockLimits = JSON.parse(r.Stock_JSON); } catch(e) {}
    let orderCaps = {};
    try { if (r.Order_Cap_JSON) orderCaps = JSON.parse(r.Order_Cap_JSON); } catch(e) {}
    const orderedCounts = countsByDate[d] || { Breakfast: {}, Lunch: {}, Dinner: {} };
    const unitsRemaining = {};
    ["Breakfast","Lunch","Dinner"].forEach(meal => {
      Object.entries(stockLimits[meal] || {}).forEach(([colKey, limit]) => {
        if (!unitsRemaining[meal]) unitsRemaining[meal] = {};
        unitsRemaining[meal][colKey] = Math.max(0, limit - (orderedCounts[meal][itemsJsonKey(colKey)] || 0));
      });
    });
    const kitchenClosed = (r.Kitchen_Closed === true ||
      String(r.Kitchen_Closed || "").toLowerCase() === "true");
    return {
      date:             d,
      breakfast:        breakfast,
      lunch_dry:        r.Lunch_Dry    || "",
      lunch_curry:      r.Lunch_Curry  || "",
      dinner_dry:       r.Dinner_Dry   || "",
      dinner_curry:     r.Dinner_Curry || "",
      cutoff_overrides: co,
      oos_items:        oosItems,
      orders_closed:    ordersClosed,
      stock_limits:     stockLimits,
      units_remaining:  unitsRemaining,
      order_caps:       orderCaps,
      order_counts:     mealOrderCounts[d] || { Breakfast: 0, Lunch: 0, Dinner: 0 },
      kitchen_closed:   kitchenClosed,
    };
  });

  return {breakfastMaster, sabjiMaster, menuEntries};
}
// ── ADMIN: SAVE MENU ─────────────────────────────────────────
function saveMenu(body) {
  const ss = getSpreadsheet();
  // Always pass full headers so schema self-heals if initSchema() was never run
  const ws = getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner",
    "OOS_JSON","Orders_Closed","Stock_JSON","Kitchen_Closed","Order_Cap_JSON"
  ]);
  const rows = getAllRows(ws);
  let hIdx = headerIndex(ws);

  // Self-heal: ensure Kitchen_Closed column exists for legacy sheets.
  if (!hIdx["Kitchen_Closed"]) {
    ws.getRange(1, ws.getLastColumn() + 1).setValue("Kitchen_Closed");
    SpreadsheetApp.flush();
    hIdx = headerIndex(ws);
  }
  // Self-heal: ensure Order_Cap_JSON column exists (per-meal max-order caps).
  if (!hIdx["Order_Cap_JSON"]) {
    ws.getRange(1, ws.getLastColumn() + 1).setValue("Order_Cap_JSON");
    SpreadsheetApp.flush();
    hIdx = headerIndex(ws);
  }

  const dateStr     = body.date;
  const existing    = rows.find(r => {
    const d = r.Date instanceof Date
      ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Date).trim();
    return d === dateStr;
  });

  // breakfast comes as array from admin, serialise to JSON string for storage
  const bfJson = body.breakfast
    ? JSON.stringify(body.breakfast)
    : (body.breakfastJson || "");

  // Preserve the existing Kitchen_Closed flag — regular menu saves
  // should never silently flip it. Use setKitchenClosed() to change it.
  const preservedKitchenClosed = existing
    ? (existing.Kitchen_Closed === true ||
       String(existing.Kitchen_Closed || "").toLowerCase() === "true")
    : false;

  const newRow = [
    dateStr,
    bfJson,
    body.lunch_dry        || body.lunchDry    || "",
    body.lunch_curry      || body.lunchCurry  || "",
    body.dinner_dry       || body.dinnerDry   || "",
    body.dinner_curry     || body.dinnerCurry || "",
    body.cutoff_breakfast || body.cutoffBf    || "",
    body.cutoff_lunch     || body.cutoffL     || "",
    body.cutoff_dinner    || body.cutoffD     || "",
    JSON.stringify(body.oos_items    || { Breakfast: [], Lunch: [], Dinner: [] }),
    JSON.stringify(body.orders_closed || {}),
    JSON.stringify(body.stock_limits || {}),
    preservedKitchenClosed ? "TRUE" : "",
    // Per-meal max-order caps, e.g. {"Breakfast":50,"Lunch":80}. Preserve any
    // existing caps if this save doesn't carry order_caps (don't silently reopen).
    (body.order_caps !== undefined)
      ? JSON.stringify(body.order_caps || {})
      : (existing && existing.Order_Cap_JSON ? String(existing.Order_Cap_JSON) : "{}"),
  ];

  if (existing) {
    ws.getRange(existing._row, 1, 1, newRow.length).setValues([newRow]);
  } else {
    ws.appendRow(newRow);
  }
  // Bust per-date menu cache and the aggregated admin-data cache
  _invalidateCache("menu_v2_" + dateStr, "adminData_v1", "kitchen_closed_dates_v1");
  return {success: true, action: existing ? "updated" : "saved"};
}

// ── ADMIN: KITCHEN CLOSURE TOGGLE ─────────────────────────────
// See Code.gs (prod) for the full behaviour contract. Mirrored here
// for the merged Apps Script project.
function setKitchenClosed(body) {
  const pin = String(body && body.pin || "").trim();
  if (pin !== ADMIN_PIN) return { success: false, error: "STRICT ADMIN PIN REQUIRED" };

  const dateStr = String(body.date || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return { success: false, error: "Invalid date format (expected YYYY-MM-DD)" };
  }
  const isClosed = (body.isClosed === true || String(body.isClosed) === "true");
  const confirmCancelOrders = (body.confirmCancelOrders === true ||
                               String(body.confirmCancelOrders) === "true");

  const ss = getSpreadsheet();
  const menuWs = getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner",
    "OOS_JSON","Orders_Closed","Stock_JSON","Kitchen_Closed","Order_Cap_JSON"
  ]);
  let mIdx = headerIndex(menuWs);
  if (!mIdx["Kitchen_Closed"]) {
    menuWs.getRange(1, menuWs.getLastColumn() + 1).setValue("Kitchen_Closed");
    SpreadsheetApp.flush();
    mIdx = headerIndex(menuWs);
  }

  if (isClosed) {
    const ordersWs = ss.getSheetByName(TAB_ORDERS);
    const oRows = ordersWs ? getAllRows(ordersWs) : [];
    const activeMatches = oRows.filter(function(r) {
      const od = r.Order_Date instanceof Date
        ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
        : String(r.Order_Date || "").trim();
      if (od !== dateStr) return false;
      return !_isOrderCancelled(r.Payment_Status);
    });

    if (activeMatches.length && !confirmCancelOrders) {
      const total = activeMatches.reduce(function(s, r) {
        return s + (Number(r.Net_Total) || 0);
      }, 0);
      const customers = {};
      activeMatches.forEach(function(r) { customers[String(r.Phone || "")] = true; });
      return {
        success: false,
        requires_confirm: true,
        orderCount: activeMatches.length,
        customerCount: Object.keys(customers).length,
        totalAmount: total,
        date: dateStr,
        message: "There are " + activeMatches.length + " active order(s) totaling ₹"
               + total + " across " + Object.keys(customers).length
               + " customer(s) for " + dateStr
               + ". Closing this day will cancel and refund all of them. Confirm?"
      };
    }

    // deleteOrder auto-detects On Account from row's Payment_Status, so
    // passing rType="none" for those is safe — it cancels the row and
    // excludes it from the next monthly bill, no payout needed.
    let cancelled = 0, refundedWallet = 0, refundedUpi = 0, onAccountAdjusted = 0;
    activeMatches.forEach(function(r) {
      const pStat = String(r.Payment_Status || "").toLowerCase();
      let rType = "none", bucket = "other";
      if (pStat === "wallet paid") { rType = "wallet"; bucket = "wallet"; }
      else if (pStat === "on account" || pStat === "onaccount") { rType = "none"; bucket = "on_account"; }
      else if (pStat === "paid") { rType = "manual_upi"; bucket = "upi"; }
      else if (pStat.indexOf("pending") !== -1) { rType = "none"; bucket = "other"; }  // unpaid — cancel only, no refund (don't inflate the UPI-refund summary total)
      try {
        const res = deleteOrder(String(r.Phone || ""), String(r.Submission_ID || ""),
                                rType, { isAdmin: true });
        if (res && res.success) {
          cancelled++;
          const amt = Number(r.Net_Total) || 0;
          if (bucket === "wallet")     refundedWallet    += amt;
          if (bucket === "upi")        refundedUpi       += amt;
          if (bucket === "on_account") onAccountAdjusted += amt;
        }
        SpreadsheetApp.flush();
      } catch(e) {
        console.error("setKitchenClosed: deleteOrder failed for " + r.Submission_ID + ": " + e.message);
      }
    });

    _writeKitchenClosedFlag(menuWs, mIdx, dateStr, true);
    _invalidateCache("menu_v2_" + dateStr, "kitchen_closed_dates_v1", "adminData_v1");

    var parts = [];
    if (refundedWallet > 0)    parts.push("₹" + refundedWallet + " refunded to wallets");
    if (refundedUpi > 0)       parts.push("₹" + refundedUpi + " queued for UPI refund");
    if (onAccountAdjusted > 0) parts.push("₹" + onAccountAdjusted + " removed from On-Account balances (no payout — just won't be billed)");
    var breakdown = parts.length ? (" — " + parts.join(", ") + ".") : ".";

    return {
      success: true, isClosed: true,
      cancelled: cancelled,
      refundedWallet: refundedWallet,
      refundedUpi: refundedUpi,
      onAccountAdjusted: onAccountAdjusted,
      message: "Kitchen closed for " + dateStr + ". " + cancelled
             + " order(s) cancelled" + breakdown
    };
  }

  _writeKitchenClosedFlag(menuWs, mIdx, dateStr, false);
  _invalidateCache("menu_v2_" + dateStr, "kitchen_closed_dates_v1", "adminData_v1");
  return { success: true, isClosed: false, message: "Kitchen re-opened for " + dateStr + "." };
}

function _writeKitchenClosedFlag(menuWs, mIdx, dateStr, isClosed) {
  const rows = getAllRows(menuWs);
  const existing = rows.find(function(x) {
    const d = x.Date instanceof Date
      ? Utilities.formatDate(x.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(x.Date || "").trim();
    return d === dateStr;
  });
  const colIdx = mIdx["Kitchen_Closed"];
  if (existing) {
    menuWs.getRange(existing._row, colIdx).setValue(isClosed ? "TRUE" : "");
  } else {
    const newRow = new Array(menuWs.getLastColumn()).fill("");
    newRow[mIdx["Date"] - 1] = dateStr;
    newRow[colIdx - 1] = isClosed ? "TRUE" : "";
    menuWs.appendRow(newRow);
  }
  SpreadsheetApp.flush();
}
// ── ADMIN: BREAKFAST MASTER CRUD ─────────────────────────────
function saveBreakfastItem(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const rows = getAllRows(ws);

  // Admin sends "default_on" (true/false); map to Active column
  const isActive = body.default_on !== false && body.default_on !== "false";

  if (body.id) {
    const r = rows.find(x => String(x.ID) === String(body.id));
    if (r) {
      const hIdx = headerIndex(ws);
      ws.getRange(r._row, hIdx["Name"]).setValue(body.name);
      ws.getRange(r._row, hIdx["Price"]).setValue(body.price);
      ws.getRange(r._row, hIdx["Active"]).setValue(isActive ? "true" : "false");
      _invalidateCache("adminData_v1");
      return {success: true};
    }
  }
  const newId = "BF-" + new Date().getTime();
  ws.appendRow([newId, body.name, body.price, isActive ? "true" : "false"]);
  _invalidateCache("adminData_v1");
  return {success: true, id: newId};
}
function deleteBreakfastItem(id) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => String(x.ID) === String(id));
  if (!r) return {success: false, error: "Not found"};
  ws.deleteRow(r._row);
  _invalidateCache("adminData_v1");
  return {success: true};
}
// ── ADMIN: SABJI MASTER CRUD ──────────────────────────────────
function saveSabjiItem(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_SABJI, []);
  const rows = getAllRows(ws);

  if (body.id) {
    const r = rows.find(x => String(x.ID) === String(body.id));
    if (r) {
      const hIdx = headerIndex(ws);
      ws.getRange(r._row, hIdx["Name"]).setValue(body.name);
      ws.getRange(r._row, hIdx["Type"]).setValue(body.type);
      ws.getRange(r._row, hIdx["Active"]).setValue(body.active !== false ? "true" : "false");
      _invalidateCache("adminData_v1");
      return {success: true};
    }
  }
  const newId = "SB-" + new Date().getTime();
  ws.appendRow([newId, body.name, body.type || "Dry", "true"]);
  _invalidateCache("adminData_v1");
  return {success: true, id: newId};
}
function deleteSabjiItem(id) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_SABJI, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => String(x.ID) === String(id));
  if (!r) return {success: false, error: "Not found"};
  ws.deleteRow(r._row);
  return {success: true};
}
function getAreas() {
  return _cachedData("areas_v1", 300, function() {
    const ss = getSpreadsheet();
    const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
    const rows = getAllRows(ws);
    // Seed defaults on first run
    if (rows.length === 0) {
      DEFAULT_AREAS.forEach(function(r) { ws.appendRow(r); });
      return DEFAULT_AREAS.map(function(r) { return {name:r[0], label:r[1], free:true}; });
    }
    return rows.map(function(r) {
      return {name: r.Area_Name, label: r.Area_Label, free: r.Free_Delivery === true || String(r.Free_Delivery).toUpperCase() === "TRUE"};
    });
  });
}
function saveArea(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
  const data = ws.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf("Area_Name");
  const labelIdx = headers.indexOf("Area_Label");
  const freeIdx = headers.indexOf("Free_Delivery");
  // Update if exists
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]).toLowerCase() === String(body.name).toLowerCase()) {
      data[i][labelIdx] = body.label;
      data[i][freeIdx]  = body.free ? "TRUE" : "FALSE";
      ws.getDataRange().setValues(data);
      _invalidateCache("areas_v1");
      return {success: true};
    }
  }
  // Add new
  ws.appendRow([body.name, body.label, body.free ? "TRUE" : "FALSE"]);
  _invalidateCache("areas_v1");
  return {success: true};
}
function deleteArea(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
  const data = ws.getDataRange().getValues();
  const nameIdx = data[0].indexOf("Area_Name");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]) === String(body.name)) {
      ws.deleteRow(i + 1);
      _invalidateCache("areas_v1");
      return {success: true};
    }
  }
  return {success: false, error: "Area not found"};
}
