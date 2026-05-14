// ============================================================
// 07_Kitchen_Delivery.gs
// Kitchen production summary, driver delivery rows, label PDFs,
  packaging expenses, en-route/delivered state.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── ROTI PACKING UTILITY ──────────────────────────────────────
function calculatePackets(total, max) {
  if (total <= 0) return [];
  if (total <= max) return [total];
  var numPacks = Math.ceil(total / max);
  var baseSize = Math.floor(total / numPacks);
  var remainder = total % numPacks;
  var packs = [];
  for (var i = 0; i < numPacks; i++) {
    packs.push(i < remainder ? baseSize + 1 : baseSize);
  }
  return packs;
}
// ── KITCHEN SUMMARY ──────────────────────────────────────────
function getKitchenSummary(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getRecentRows(ws, 1500);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date;
  });

  var meals = {};
  var ROTI_COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri"];
  var ROTI_LIMITS = {
    "Chapati":6, "Without_Oil_Chapati":6,
    "Phulka":12, "Ghee_Phulka":12,
    "Jowar_Bhakri":2, "Bajra_Bhakri":2
  };
  var menu = getMenu(date);

  var orders = [];

  dayRows.forEach(function(r) {
    if (_isOrderCancelled(r.Payment_Status)) return; // exclude cancelled/verify-pending orders
    var meal = String(r.Meal_Type || "");
    if (!meal) return;
    if (!meals[meal]) meals[meal] = {count: 0};
    var m = meals[meal];
    m.count++;

    // For Labels Tab
    var summaryParts = [];
    if (meal === "Breakfast") {
      if (!m.items) m.items = {};
      for (var n = 1; n <= 4; n++) {
        var item = String(r["BF_Item_"+n] || "").trim();
        var qty  = Number(r["BF_Qty_"+n]) || 0;
        if (item && qty > 0) {
          m.items[item] = (m.items[item] || 0) + qty;
          summaryParts.push(qty + " " + item);
        }
      }
      var curdBf = Number(r.Curd) || 0;
      if (curdBf > 0) {
        m.items["Curd"] = (m.items["Curd"] || 0) + curdBf;
        summaryParts.push(curdBf + " Curd");
      }
    } else {
      if (!m.rotis) m.rotis = {};
      if (!m.rotiMatrix) {
        m.rotiMatrix = {};
        ROTI_COLS.forEach(function(c) { m.rotiMatrix[c] = {}; });
      }
      ROTI_COLS.forEach(function(c) {
        var q = Number(r[c]) || 0;
        if (q > 0) {
          m.rotis[c] = (m.rotis[c] || 0) + q;
          summaryParts.push(q + " " + c.replace(/_/g, " "));
          var packs = calculatePackets(q, ROTI_LIMITS[c]);
          packs.forEach(function(p) {
            m.rotiMatrix[c][p] = (m.rotiMatrix[c][p] || 0) + 1;
          });
        }
      });
      if (!m.sabji) {
        m.sabji = {
          dry_kg: 0, curry_kg: 0,
          dry_name: (meal === "Lunch" ? menu.lunch_dry : menu.dinner_dry) || "Sabji (Dry)",
          curry_name: (meal === "Lunch" ? menu.lunch_curry : menu.dinner_curry) || "Sabji (Curry)",
          dry_mini: 0, dry_full: 0, curry_mini: 0, curry_full: 0
        };
      }
      var dMini = Number(r.Dry_Sabji_Mini)||0;
      var dFull = Number(r.Dry_Sabji_Full)||0;
      var cMini = Number(r.Curry_Sabji_Mini)||0;
      var cFull = Number(r.Curry_Sabji_Full)||0;
      m.sabji.dry_mini  += dMini;
      m.sabji.dry_full  += dFull;
      m.sabji.curry_mini += cMini;
      m.sabji.curry_full += cFull;
      
      if (dMini > 0) summaryParts.push(dMini + " Mini Dry");
      if (dFull > 0) summaryParts.push(dFull + " Full Dry");
      if (cMini > 0) summaryParts.push(cMini + " Mini Curry");
      if (cFull > 0) summaryParts.push(cFull + " Full Curry");

      if (!m.other) m.other = {Dal:{kg:0, count:0}, Rice:{count:0}, Salad:{count:0}, Curd:{count:0}};
      if (!m.riceMatrix)  m.riceMatrix  = {};
      if (!m.saladMatrix) m.saladMatrix = {};
      if (!m.curdMatrix)  m.curdMatrix  = {};

      var dalQ = Number(r.Dal)   || 0;
      var riceQ = Number(r.Rice)  || 0;
      var saladQ = Number(r.Salad) || 0;
      var curdQ = Number(r.Curd)  || 0;
      
      m.other.Dal.kg      += dalQ * 1.33;
      m.other.Dal.count   += dalQ;
      m.other.Rice.count  += riceQ;
      m.other.Salad.count += saladQ;
      m.other.Curd.count  += curdQ;

      // Matrix calculations
      if (riceQ > 0) {
        var rPacks = calculatePackets(riceQ, 3); // RICE_LIMIT = 3
        rPacks.forEach(function(p) { m.riceMatrix[p] = (m.riceMatrix[p] || 0) + 1; });
      }
      if (saladQ > 0) {
        var sPacks = calculatePackets(saladQ, 4); // SALAD_LIMIT = 4
        sPacks.forEach(function(p) { m.saladMatrix[p] = (m.saladMatrix[p] || 0) + 1; });
      }
      if (curdQ > 0) {
        var cPacks = calculatePackets(curdQ, 2); // CURD_LIMIT = 2 (50g cups: 1 or 2 per packet)
        cPacks.forEach(function(p) { m.curdMatrix[p] = (m.curdMatrix[p] || 0) + 1; });
      }

      if (dalQ > 0) summaryParts.push(dalQ + " Dal");
      if (riceQ > 0) summaryParts.push(riceQ + " Rice");
      if (saladQ > 0) summaryParts.push(saladQ + " Salad");
      if (curdQ > 0) summaryParts.push(curdQ + " Curd");
    }

    orders.push({
      Submission_ID: String(r.Submission_ID || ""),
      Customer_Name: String(r.Customer_Name || ""),
      Meal_Type: meal,
      summary: summaryParts.join(", "),
      items: {
        Chapati: Number(r.Chapati)||0, Without_Oil_Chapati: Number(r.Without_Oil_Chapati)||0,
        Phulka: Number(r.Phulka)||0, Ghee_Phulka: Number(r.Ghee_Phulka)||0,
        Jowar_Bhakri: Number(r.Jowar_Bhakri)||0, Bajra_Bhakri: Number(r.Bajra_Bhakri)||0,
        Dry_Sabji_Mini: Number(r.Dry_Sabji_Mini)||0, Dry_Sabji_Full: Number(r.Dry_Sabji_Full)||0,
        Curry_Sabji_Mini: Number(r.Curry_Sabji_Mini)||0, Curry_Sabji_Full: Number(r.Curry_Sabji_Full)||0,
        Dal: Number(r.Dal)||0, Rice: Number(r.Rice)||0, Salad: Number(r.Salad)||0, Curd: Number(r.Curd)||0,
        "Kanda Poha": Number(r["Kanda Poha"])||0, "Ghee Upma": Number(r["Ghee Upma"])||0,
        "Thalipeeth": Number(r["Thalipeeth"])||0, "Palak Paratha": Number(r["Palak Paratha"])||0,
        "Paneer Paratha": Number(r["Paneer Paratha"])||0, "Methi Thepla": Number(r["Methi Thepla"])||0,
        "Sabudana Khichdi": Number(r["Sabudana Khichdi"])||0
      },
      Special_Notes_Kitchen: String(r.Special_Notes_Kitchen || ""),
      Special_Notes_Delivery: String(r.Special_Notes_Delivery || ""),
      Delivery_Point: String(r.Delivery_Point || ""),
      marathiNotes: String(r.marathiNotes || ""),
      Packed: r.Packed === true || String(r.Packed).toLowerCase() === "true"
    });
  });

  ["Lunch","Dinner"].forEach(function(meal) {
    if (!meals[meal]) return;
    var m = meals[meal];
    if (m.other && m.other.Dal) m.other.Dal.kg = Math.round(m.other.Dal.kg * 100) / 100;
  });

  return {
    date: date,
    meals: meals,
    orders: orders,
    cutoffs: menu.cutoff_overrides || {}
  };
}
// ── DRIVER ORDERS ─────────────────────────────────────────────
function getDriverOrders(date) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getRecentRows(ws, 1500);
  var meals = {Breakfast: [], Lunch: [], Dinner: []};

  // Load delivery status from SK_Deliveries tab (both EnRoute_At and Delivered_At)
  var delMap = {};
  var delWs  = ss.getSheetByName("SK_Deliveries");
  if (delWs) {
    getRecentRows(delWs, 1500).forEach(function(r) {
      var sid = String(r.Submission_ID || "").trim();
      if (sid) delMap[sid] = {
        deliveredAt: String(r.Delivered_At || ""),
        enRouteAt:   String(r.EnRoute_At   || "")
      };
    });
  }

  // Load customer meal preferences (Source of Truth)
  var custMap = {};
  var custWs = ss.getSheetByName(TAB_CUSTOMERS);
  if (custWs) {
    getAllRows(custWs).forEach(function(r) {
      var ph = _normalizePhone(r.Phone);
      if (ph) {
        custMap[ph] = {
          mealAddresses: r.Meal_Addresses || ""
        };
      }
    });
  }

  rows.forEach(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    if (d !== date) return;
    if (_isOrderCancelled(r.Payment_Status)) return;
    // var area = String(r.Area || "").trim();
    // if (area.toLowerCase().includes("pickup")) return;
    var area = String(r.Area || "").trim();
    var meal = String(r.Meal_Type || "");
    if (!meals[meal]) return;
    var sid = String(r.Submission_ID || "");
    var normP = _normalizePhone(r.Phone);
    meals[meal].push({
      submissionId:  sid,
      name:          String(r.Customer_Name || ""),
      phone:         String(r.Phone || ""),
      area:          area,
      address:       String(r.Full_Address || ""),
      landmark:      String(r.Landmark || ""),
      deliveryPoint: String(r.Delivery_Point || ""),
      maps:          String(r.Maps_Link || ""),
      notes:         String(r.Special_Notes_Delivery || ""),
      deliveredAt:   (delMap[sid] && delMap[sid].deliveredAt) || "",
      enRouteAt:     (delMap[sid] && delMap[sid].enRouteAt)   || "",
      amount:        Number(r.Net_Total || r.Food_Subtotal || 0),
      paymentStatus: String(r.Payment_Status || ""),
      mealAddresses: custMap[normP] ? custMap[normP].mealAddresses : ""
    });
  });



  return {date: date, meals: meals};
}
function createDeliverySheet(date, meal) {
  var data = getDriverOrders(date);
  var orders = (data.meals && data.meals[meal]) || [];

  var mealLabel = meal;
  var dateParts = date.split("-"); // yyyy-mm-dd
  var displayDate = dateParts[2] + "/" + dateParts[1] + "/" + dateParts[0];
  var title = "Delivery — " + mealLabel + " " + displayDate;

  var ss   = SpreadsheetApp.create(title);
  var sheet = ss.getActiveSheet();
  sheet.setName(mealLabel);

  // Headers
  var headers = ["Name", "Phone", "Address", "Maps Link", "Landmark", "Delivery Point", "Notes"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1E1240");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);

  // Data rows
  if (orders.length > 0) {
    var rows = orders.map(function(o) {
      return [
        o.name        || "",
        o.phone       || "",
        o.address     || "",
        o.maps        || "",
        o.landmark    || "",
        o.deliveryPoint || "",
        o.notes       || ""
      ];
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize columns for readability
  headers.forEach(function(_, i) { sheet.autoResizeColumn(i + 1); });

  // Freeze header row
  sheet.setFrozenRows(1);

  // Make the sheet accessible to anyone with the link (view + comment)
  var file = DriveApp.getFileById(ss.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  return { success: true, url: ss.getUrl(), title: title, count: orders.length };
}
// ── LABEL ORDERS ──────────────────────────────────────────────
function getLabelOrders(date, meal) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);
  var COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
              "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full","Dal","Rice","Salad"];

  var orders = rows
    .filter(function(r) {
      var d = r.Order_Date instanceof Date
        ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
        : String(r.Order_Date).trim();
      return d === date && String(r.Meal_Type) === meal && !_isOrderCancelled(r.Payment_Status);
    })
    .map(function(r) {
      var obj = {
        name:  String(r.Customer_Name || ""),
        area:  String(r.Area || ""),
        notes: String(r.Special_Notes || ""),
        Curd:  Number(r.Curd) || 0
      };
      if (meal === "Breakfast") {
        for (var n = 1; n <= 4; n++) {
          obj["BF_Item_"+n] = String(r["BF_Item_"+n] || "");
          obj["BF_Qty_"+n]  = Number(r["BF_Qty_"+n])  || 0;
        }
      } else {
        COLS.forEach(function(col) { obj[col] = Number(r[col]) || 0; });
      }
      return obj;
    });

  return {orders: orders};
}
function getPackagingExpenses(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date;
  });

  if (dayRows.length === 0) return {date: date, orderCount: 0, meals: {}, items: [], total: 0};

  var counts = {};
  var mealCounts = {Breakfast:0, Lunch:0, Dinner:0};
  function add(key, qty) { if (qty > 0) counts[key] = (counts[key]||0) + qty; }

  dayRows.forEach(function(r) {
    var meal = String(r.Meal_Type || "");
    if (mealCounts[meal] !== undefined) mealCounts[meal]++;

    add("Label / Sticker", 1);
    if (meal === "Breakfast") {
      add("Breakfast Box", 1);
      add("Curd Container", Number(r.Curd) || 0);
    } else {
      add("Delivery Bag", 1);
      var breadCols = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri"];
      var hasBread = breadCols.some(function(c) { return (Number(r[c])||0) > 0; });
      if (hasBread) add("Bread Packet", 1);
      add("Sabji Container (Mini)", (Number(r.Dry_Sabji_Mini)||0) + (Number(r.Curry_Sabji_Mini)||0));
      add("Sabji Container (Full)", (Number(r.Dry_Sabji_Full)||0) + (Number(r.Curry_Sabji_Full)||0));
      add("Dal Container",          Number(r.Dal)   || 0);
      add("Rice Container",         Number(r.Rice)  || 0);
      add("Salad Container",        Number(r.Salad) || 0);
      add("Curd Container",         Number(r.Curd)  || 0);
    }
  });

  var itemOrder = ["Breakfast Box","Delivery Bag","Label / Sticker","Bread Packet",
                   "Sabji Container (Mini)","Sabji Container (Full)",
                   "Dal Container","Rice Container","Salad Container","Curd Container"];
  var items = [];
  var total = 0;
  itemOrder.forEach(function(key) {
    var qty = counts[key] || 0;
    if (!qty) return;
    var unitCost = PKG_UNIT_COSTS[key] || 0;
    var t = qty * unitCost;
    items.push({name: key, qty: qty, unitCost: unitCost, total: t});
    total += t;
  });

  var mealsOut = {};
  Object.keys(mealCounts).forEach(function(m) { if (mealCounts[m] > 0) mealsOut[m] = mealCounts[m]; });

  return {date: date, orderCount: dayRows.length, meals: mealsOut, items: items, total: total};
}
// ── PACKAGING EXPENSES — RANGE ───────────────────────────────
function getPackagingExpensesRange(from, to) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var rangeRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d >= from && d <= to && !_isOrderCancelled(r.Payment_Status);
  });

  // Group by date
  var byDate = {};
  rangeRows.forEach(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(r);
  });

  var PKG_COSTS = PKG_UNIT_COSTS;
  var itemOrder = ["Breakfast Box","Delivery Bag","Label / Sticker","Bread Packet",
                   "Sabji Container (Mini)","Sabji Container (Full)",
                   "Dal Container","Rice Container","Salad Container","Curd Container"];

  function calcDay(dateStr, dayRows) {
    var counts = {}, mealCounts = {Breakfast:0, Lunch:0, Dinner:0};
    function add(key, qty) { if (qty > 0) counts[key] = (counts[key]||0) + qty; }
    dayRows.forEach(function(r) {
      var meal = String(r.Meal_Type || "");
      if (mealCounts[meal] !== undefined) mealCounts[meal]++;
      add("Label / Sticker", 1);
      if (meal === "Breakfast") {
        add("Breakfast Box", 1);
        add("Curd Container", Number(r.Curd) || 0);
      } else {
        add("Delivery Bag", 1);
        var breadCols = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri"];
        if (breadCols.some(function(c){ return (Number(r[c])||0)>0; })) add("Bread Packet", 1);
        add("Sabji Container (Mini)", (Number(r.Dry_Sabji_Mini)||0)+(Number(r.Curry_Sabji_Mini)||0));
        add("Sabji Container (Full)", (Number(r.Dry_Sabji_Full)||0)+(Number(r.Curry_Sabji_Full)||0));
        add("Dal Container",  Number(r.Dal)||0);
        add("Rice Container", Number(r.Rice)||0);
        add("Salad Container",Number(r.Salad)||0);
        add("Curd Container", Number(r.Curd)||0);
      }
    });
    var items = [], total = 0;
    itemOrder.forEach(function(key) {
      var qty = counts[key]||0; if (!qty) return;
      var unitCost = PKG_COSTS[key]||0, t = qty*unitCost;
      items.push({name:key, qty:qty, unitCost:unitCost, total:t});
      total += t;
    });
    var mealsOut = {};
    Object.keys(mealCounts).forEach(function(m){ if(mealCounts[m]>0) mealsOut[m]=mealCounts[m]; });
    return {date:dateStr, orderCount:dayRows.length, meals:mealsOut, items:items, total:total};
  }

  // Build per-day results
  var days = Object.keys(byDate).sort().map(function(d){ return calcDay(d, byDate[d]); });

  // Aggregate totals
  var aggCounts = {}, aggMeals = {Breakfast:0,Lunch:0,Dinner:0}, aggTotal = 0, aggOrders = 0;
  days.forEach(function(day) {
    aggOrders += day.orderCount;
    aggTotal  += day.total;
    Object.keys(day.meals).forEach(function(m){ aggMeals[m]=(aggMeals[m]||0)+day.meals[m]; });
    day.items.forEach(function(it){ aggCounts[it.name]=(aggCounts[it.name]||0)+it.qty; });
  });
  var aggItems = [];
  itemOrder.forEach(function(key) {
    var qty = aggCounts[key]||0; if (!qty) return;
    var unitCost = PKG_COSTS[key]||0, t=qty*unitCost;
    aggItems.push({name:key, qty:qty, unitCost:unitCost, total:t});
  });

  return {
    from: from, to: to,
    orderCount: aggOrders,
    total: aggTotal,
    meals: aggMeals,
    items: aggItems,
    days: days
  };
}
// ── LABEL DRIVE SAVE ─────────────────────────────────────────
function saveLabels(body) {
  var date   = body.date;  // "2026-03-18"
  var meal   = body.meal;  // "Lunch"
  var pdfB64 = body.pdf;   // base64-encoded PDF bytes

  var parts      = date.split("-");
  var year       = parts[0];                                        // "2026"
  var monthNum   = parseInt(parts[1], 10);                          // 3
  var monthNames = ["January","February","March","April","May","June",
                    "July","August","September","October","November","December"];
  var monthName  = monthNames[monthNum - 1];                        // "March"
  var mealAbbrev = meal.toLowerCase().substring(0, 7);              // "breakfa" / "lunch" / "dinner"
  var langCode   = (body.lang === "Devanagari") ? "mar" : "eng";    // "eng" / "mar"
  var filename   = "labels_" + mealAbbrev + "_" + langCode + "_" + date + "_58x25.pdf";
  // e.g. "labels_breakfa_eng_2026-03-05_58x25.pdf"

  var folder = getOrCreateFolderPath([
    "Svaadh Kitchen", "Accounting", "Tally Form Daily Sheets",
    "Processed_Orders", "Labels", year, monthName
  ]);

  // Replace existing file to avoid duplicates
  var existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var pdfBlob = Utilities.newBlob(Utilities.base64Decode(pdfB64), "application/pdf", filename);
  var file = folder.createFile(pdfBlob);
  return {url: file.getUrl(), name: filename};
}
// ── DELETED OBSOLETE ADMIN CANCEL ORDER (Merged with main) ──
/**
 * Marks a specific order as 'Packed' in the SK_Orders sheet.
 * Called by the Kitchen Dashboard (kitchen.html) via APPS_SCRIPT_URL.
 */
function markOrderPacked(body) {
  var id = body.submissionId;
  if (!id) return {success:false, error: "submissionId required"};

  var ss    = getSpreadsheet();
  var ws    = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var hIdx  = headerIndex(ws);
  var rows  = getAllRows(ws);
  
  if (hIdx.Packed === undefined) return {success:false, error: "Packed column not found"};

  var order = rows.find(function(r) {
    return String(r.Submission_ID) === String(id);
  });

  if (order) {
    ws.getRange(order._row, hIdx.Packed).setValue(true);
    return {success:true};
  }
  
  return {success:false, error: "Order not found"};
}
// ── LIVE TRACKER LOGIC: En Route & Delivered ──────────────────────────────────
function batchMarkEnRoute(body) {
  var sids      = body.submissionIds;
  var enRouteAt = body.enRouteAt;
  if (!sids || !sids.length) return {success:false, error:"submissionIds required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  var data = ws.getDataRange().getValues();
  var headers = data[0];
  var erIdx   = headers.indexOf("EnRoute_At");
  var sidIdx  = headers.indexOf("Submission_ID");

  if (erIdx === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
    erIdx = headers.length;
  }

  // Create lookup for existing rows
  var sidToRowMap = {};
  for (var i = 1; i < data.length; i++) {
    sidToRowMap[String(data[i][sidIdx])] = i + 1;
  }

  sids.forEach(function(sid) {
    var row = sidToRowMap[String(sid)];
    if (row) {
      ws.getRange(row, erIdx + 1).setValue(enRouteAt);
    } else {
      // Append if not found (though usually we expect them to be found if already rendered)
      ws.appendRow([sid, "", enRouteAt]);
    }
  });

  return {success:true};
}
function markEnRoute(body) {
  var sid         = body.submissionId;
  var enRouteAt   = body.enRouteAt;
  if (!sid) return {success:false, error:"submissionId required"};

  var ss  = getSpreadsheet();
  var ws  = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  var rows = getAllRows(ws);
  
  // Ensure "EnRoute_At" header exists
  var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  if (headers.indexOf("EnRoute_At") === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
  }
  var hIdx = headerIndex(ws);

  var existing = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Submission_ID) === String(sid)) { existing = rows[i]; break; }
  }
  if (existing) {
    ws.getRange(existing._row, hIdx["EnRoute_At"]).setValue(enRouteAt);
  } else {
    // Note: Append Row pushes based on sheet width. Best to be safe using indices.
    var newRowArr = [];
    newRowArr[hIdx["Submission_ID"] - 1] = sid;
    newRowArr[hIdx["Delivered_At"] - 1] = "";
    newRowArr[hIdx["EnRoute_At"] - 1] = enRouteAt;
    ws.appendRow(newRowArr);
  }
  return {success:true, submissionId:sid, enRouteAt:enRouteAt};
}
// ── MARK DELIVERED ────────────────────────────────────────────────────────────
function markDelivered(body) {
  var sid         = body.submissionId;
  var deliveredAt = body.deliveredAt;
  if (!sid) return {success:false, error:"submissionId required"};

  var ss  = getSpreadsheet();
  // Ensure we define the tab properly
  var ws  = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  
  var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  if (headers.indexOf("EnRoute_At") === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
  }
  
  var rows = getAllRows(ws);
  var hIdx = headerIndex(ws);

  var existing = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Submission_ID) === sid) { existing = rows[i]; break; }
  }
  if (existing) {
    ws.getRange(existing._row, hIdx["Delivered_At"]).setValue(deliveredAt);
  } else {
    var newRowArr = [];
    newRowArr[hIdx["Submission_ID"] - 1] = sid;
    newRowArr[hIdx["Delivered_At"] - 1] = deliveredAt;
    newRowArr[hIdx["EnRoute_At"] - 1] = "";
    ws.appendRow(newRowArr);
  }
  return {success:true, submissionId:sid, deliveredAt:deliveredAt};
}
function _getDeliveryPointLabel(key) {
  if (!key) return "Handover at Doorstep";
  const map = {
    door: "Handover at Doorstep / Office Door",
    bell_keep: "Keep outside & Ring bell",
    lobby_handoff: "Handover at Lobby / Reception",
    lobby_keep: "Keep at Lobby / Reception",
    gate_handoff: "Handover at Security Gate",
    gate_keep: "Keep at security cabin",
    comedown: "Customer will come down",
    other: "Other (see instructions)"
  };
  return map[key] || key;
}
function _deriveMapsLink(addr, society) {
  const dict = {
    "Laburnum Park": "https://maps.app.goo.gl/nEApFaLe5x4PzuHd8",
    "Magarpatta City": "https://maps.app.goo.gl/wEndRh6jnkjL1GRC9",
    "Amanora Mall": "https://maps.app.goo.gl/Wd2FxrytcABk9Xty6",
    "Pentagon 1": "https://maps.app.goo.gl/BQKFrtdmLv9sK8tF8",
    "IZiel": "https://maps.app.goo.gl/BQKFrtdmLv9sK8tF8",
    "Shree Lakshmi Vihar": "https://maps.app.goo.gl/LWr5zhbiXHN9sh1F8",
    "Desire Tower": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 18": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 21": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 22": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Greenville": "https://maps.app.goo.gl/FvpENaW8aFzXTj916",
    "Cosmos": "https://maps.app.goo.gl/bYu8gnb1ZVQgAwse9",
    "Prime Wing Cosmos": "https://maps.app.goo.gl/eUHhhV9oS1MPQutq8",
    "Zinnia": "https://maps.app.goo.gl/iWcMqnXCDRkesdnT7",
    "Trillium": "https://maps.app.goo.gl/cnYLtqjxaqDtexzKA",
    "Sudarshan Heritage": "https://maps.app.goo.gl/QJAAaugYVcD18eg4A",
    "Daffodils": "https://maps.app.goo.gl/vxXz2LuxFCZPcvTZ8",
    "Amanora Gold Tower": "https://maps.app.goo.gl/PgXAJ8nz7xKjnUZ7A",
    "Gateway Tower": "https://maps.app.goo.gl/YCh4dHET1ZsrLZqv6",
    "T100 Gateway Towers": "https://maps.app.goo.gl/YCh4dHET1ZsrLZqv6",
    "Wework Futura": "https://maps.app.goo.gl/u8JaNJYbpHPXBxdp9",
    "Jasminium": "https://maps.app.goo.gl/HiGZSmvbb3SFnTKP6",
    "Jasminium Society": "https://maps.app.goo.gl/pMXdyNSMu3kEMuec6",
    "Vascon Ela": "https://maps.app.goo.gl/MXhnTLvycEENcrNA7",
    "Marvel Fuego": "https://maps.app.goo.gl/izc9t6cXWWYiciyv8",
    "Heliconia": "https://maps.app.goo.gl/tFu79S7KHv6L48XG8",
    "Kumar Paradise": "https://maps.app.goo.gl/utQB7yFo4jMgVKAv6",
    "Vrindavan Heights": "https://maps.app.goo.gl/qcS8v4dVtbx1rBg46",
    "Sai Tower": "https://maps.app.goo.gl/dWigWD2KHXsE5YrT7",
    "Future Towers": "https://maps.app.goo.gl/pvhvaMsWS8n3x5Cq8",
    "Annexe Society": "https://maps.app.goo.gl/uAuA67dHxgLyfVX29",
    "Imperial Heights": "https://maps.app.goo.gl/Dt9HZVgBzB7iNdhu9",
    "Tulja Tower": "https://maps.app.goo.gl/VeUw6EciJckATaZU6",
    "Torana Kamdhenu": "https://maps.app.goo.gl/9PMRz86iGNjv8UZEA",
    "Cybercity": "https://maps.app.goo.gl/5ASLF1yDeoKH7Wq86",
    "Cyber City": "https://maps.app.goo.gl/5ASLF1yDeoKH7Wq86",
    "Samarth Shrushti": "https://maps.app.goo.gl/1eLo6vz3mTWu1BXB8",
    "Grevillea": "https://maps.app.goo.gl/ppTArZ12auPnRR5E8",
    "Orient Garden": "https://maps.app.goo.gl/vFrsEoijMPVV7a3U7",
    "Bhosale Nagar": "https://maps.app.goo.gl/5MzXtAZmtZvD9D9o6",
    "Amar Ornate": "https://maps.app.goo.gl/mtFcV35i5gpPz4BG7",
    "DSK Sunderban": "https://maps.app.goo.gl/xu4fiGtFLbgp3Kxs9",
    "Aruna Girls PG": "https://maps.app.goo.gl/FUxKBQ6iQKt64ji36",
    "Aruna PG": "https://maps.app.goo.gl/FUxKBQ6iQKt64ji36",
    "Mams Bungalow": "https://maps.app.goo.gl/gX94MR52LDXzbq7h8",
    "Gardenia": "https://maps.app.goo.gl/tVsv9NZvBzxqg1aU6",
    "Palazzo": "https://maps.app.goo.gl/fd1Na7Lenjh1AwiH6",
    "Kumar Picasso": "https://maps.app.goo.gl/WpDwFWDGJVSRx9r88",
    "Roystonea": "https://maps.app.goo.gl/6xfQ2VNEc4CMnRgD7",
    "Pankaj Avenue": "https://maps.app.goo.gl/HzLBbS8L5o8zkHgx8",
    "Unika": "https://maps.app.goo.gl/CvgQwPzkTRUT5dHz6",
    "Amanora Ascent": "https://maps.app.goo.gl/xgMEJAbrEbD2PpJ69",
    "Vanshree": "https://maps.app.goo.gl/cW1gjcsVkStwudfZ7",
    "Aspire Towers": "https://maps.app.goo.gl/dZ7V6SVs7SX63BhG8",
    "Naren Bliss": "https://maps.app.goo.gl/RJbXCqWcZSZTieBS7",
    "Solitaire": "https://maps.app.goo.gl/RsrdFp6gYQNbWzzq7",
    "Sylvania": "https://maps.app.goo.gl/j5UhPibRVDxhkdRY7",
    "Erica": "https://maps.app.goo.gl/yTpThD413d3q2FL38",
    "Om Balaji Darshan": "https://maps.app.goo.gl/KLDrXo2DZcUUkJzV9",
    "Adreno Towers": "https://maps.app.goo.gl/725jneTokL1ahFsF9",
    "Neo Towers": "https://maps.app.goo.gl/sJy9YiqFcEEDMgWz9",
    "Leisure Town": "https://maps.app.goo.gl/gvVkgfZtLfhpRk1W8",
    "Sundar Sankul": "https://share.google/aOAcluoLyIMzc3I1L",
    "Kumar Purab": "https://maps.app.goo.gl/FGez7Rv63NzNRLeaA",
    "Marvel Azure": "https://maps.app.goo.gl/YaeQHEbg8D4HAALDA",
    "Hrishikesh housing": "https://maps.app.goo.gl/PsyUKTGt4Rj9qPsVA"
  };

  const str = (addr + " " + society).toLowerCase();
  for (let key in dict) {
    if (str.includes(key.toLowerCase())) {
      return dict[key];
    }
  }
  return "";
}
