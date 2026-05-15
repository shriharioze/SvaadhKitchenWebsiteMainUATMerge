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
    units_remaining: unitsRemaining
  };
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
    const orderedCounts = countOrderedUnits(allOrdersAdm, d);
    const unitsRemaining = {};
    ["Breakfast","Lunch","Dinner"].forEach(meal => {
      Object.entries(stockLimits[meal] || {}).forEach(([colKey, limit]) => {
        if (!unitsRemaining[meal]) unitsRemaining[meal] = {};
        unitsRemaining[meal][colKey] = Math.max(0, limit - (orderedCounts[meal][itemsJsonKey(colKey)] || 0));
      });
    });
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
    "OOS_JSON","Orders_Closed","Stock_JSON"
  ]);
  const rows = getAllRows(ws);
  const hIdx = headerIndex(ws);

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
  ];

  if (existing) {
    ws.getRange(existing._row, 1, 1, newRow.length).setValues([newRow]);
  } else {
    ws.appendRow(newRow);
  }
  // Bust per-date menu cache and the aggregated admin-data cache
  _invalidateCache("menu_v2_" + dateStr, "adminData_v1");
  return {success: true, action: existing ? "updated" : "saved"};
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
