// ============================================================
// 08_Analytics_Inventory.gs
// Business analytics, churn report, GA4 sync, inventory entries,
  expense tracking.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

function syncGA4Data() {
  const propertyId = GA4_PROPERTY_ID;
  if (!propertyId) return "Error: GA4_PROPERTY_ID not set.";

  const request = {
    dimensions: [
      { name: 'date' },
      { name: 'sessionSource' },
      { name: 'deviceCategory' }
    ],
    metrics: [
      { name: 'activeUsers' },
      { name: 'sessions' },
      { name: 'screenPageViews' },
      { name: 'engagementRate' },
      { name: 'averageSessionDuration' },
      { name: 'eventCount' }
    ],
    dateRanges: [{ startDate: '30daysAgo', endDate: 'today' }]
  };

  try {
    const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
    if (!response.rows || response.rows.length === 0) return "No data found in GA4.";

    const ss = getSpreadsheet();
    const ws = getOrCreateTab(ss, TAB_GA4_METRICS, GA4_HEADERS);
    
    // Snapshot approach: overwrite the tab with the latest window
    if (ws.getLastRow() > 1) {
      ws.getRange(2, 1, ws.getLastRow() - 1, GA4_HEADERS.length).clearContent();
    }

    const rows = response.rows.map(row => {
      // Format YYYYMMDD to YYYY-MM-DD
      const rawDate = row.dimensionValues[0].value;
      const formattedDate = rawDate.substring(0,4) + "-" + rawDate.substring(4,6) + "-" + rawDate.substring(6,8);
      
      return [
        formattedDate,
        row.dimensionValues[1].value, // Source
        row.dimensionValues[2].value, // Device
        ...row.metricValues.map(mv => mv.value)
      ];
    });
    
    // Sort by date descending, then source
    rows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
      return a[1].localeCompare(b[1]);
    });

    if (rows.length > 0) {
      ws.getRange(2, 1, rows.length, GA4_HEADERS.length).setValues(rows);
    }
    
    return "Successfully synced " + rows.length + " data points (Date/Source/Device combinations).";
  } catch (e) {
    console.error("GA4 Sync Error:", e);
    return "Error: " + e.message;
  }
}
function setupAnalyticsTrigger() {
  // Remove existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'syncGA4Data') ScriptApp.deleteTrigger(t);
  });
  
  // Create new daily trigger
  ScriptApp.newTrigger('syncGA4Data')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
    
  return "Daily GA4 sync trigger set for 1:00 AM.";
}
// ── ANALYTICS ─────────────────────────────────────────────────────────────────
function getAnalytics(p) {
  var dateFrom = p.dateFrom, dateTo = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};
  var ss  = getSpreadsheet();
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };
  // Pull from BOTH the live sheet and any archived monthly files that
  // overlap this date range. 10-min CacheService cache keeps repeat
  // queries fast.
  var combined = getOrdersInRangeWithArchive(dateFrom, dateTo);
  var rows = combined.filter(function(r) {
    return !_isOrderCancelled(r.Payment_Status);
  });
  var liveCount = 0;
  try {
    var ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
    liveCount = getAllRows(ws).filter(function(r) {
      var d = fmtDate(r.Order_Date);
      return d >= dateFrom && d <= dateTo;
    }).length;
  } catch(_) {}
  var archivedCount = combined.length - liveCount;

  // ── Option B: Exact Small Order Fee backfill ──────────────────────────────
  // Pre-pass 1: build VIP set from profiles (Fee_Exempt = Yes)
  var profWs   = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var profRows = getAllRows(profWs);
  var vipSet   = {};
  profRows.forEach(function(pr) {
    if (pr.Fee_Exempt === "Yes" || pr.Fee_Exempt === true) {
      vipSet[String(pr.Phone||"").trim()] = true;
    }
  });

  // Pre-pass 2: for every phone+date combo, sum food subtotals & count distinct meals
  // This lets us know if the combined day total reached the free-delivery threshold,
  // which also waives the small order fee.
  var dayTotals = {}; // key = phone+"_"+date  →  { foodTotal, mealCount }
  rows.forEach(function(r) {
    var ph   = String(r.Phone||"").trim();
    var d    = fmtDate(r.Order_Date);
    var food = Number(r.Food_Subtotal)||0;
    var key  = ph + "_" + d;
    if (!dayTotals[key]) dayTotals[key] = { foodTotal:0, meals:{} };
    dayTotals[key].foodTotal += food;
    dayTotals[key].meals[String(r.Meal_Type||"")] = true;
  });

  // Helper: calculate small fee for a row using exact rules
  function calcSmallFee(r) {
    var stored = r.Small_Order_Fee;
    // If the column exists and has a numeric value, trust it
    if (stored !== undefined && stored !== null && stored !== "" && !isNaN(Number(stored))) {
      return Number(stored);
    }
    // Backfill for old rows
    var meal = String(r.Meal_Type||"");
    if (meal !== "Lunch" && meal !== "Dinner") return 0; // Breakfast: never charged
    var food = Number(r.Food_Subtotal)||0;
    if (food <= 0 || food >= 50) return 0;              // Only charged when sub < ₹50
    var area = String(r.Area||"").trim();
    if (area === "Self Pickup") return 0;                // Pickup: waived
    var ph  = String(r.Phone||"").trim();
    if (vipSet[ph]) return 0;                            // VIP: waived
    // Check if combined day food total crossed free-delivery threshold
    var d   = fmtDate(r.Order_Date);
    var key = ph + "_" + d;
    var dt  = dayTotals[key] || {foodTotal:0, meals:{}};
    var mealCount    = Object.keys(dt.meals).length;
    var threshold    = mealCount <= 1 ? 100 : 150;
    if (dt.foodTotal >= threshold) return 0;             // Day crossed threshold: waived
    return 10;
  }
  // ── End backfill helper ───────────────────────────────────────────────────
  var LUNCH_COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
    "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full","Dal","Rice","Salad","Curd"];
  var COL_DISP = {"Chapati":"Chapati","Without_Oil_Chapati":"WO Chapati","Phulka":"Phulka","Ghee_Phulka":"Ghee Phulka",
    "Jowar_Bhakri":"Jowar Bhakri","Bajra_Bhakri":"Bajra Bhakri","Dry_Sabji_Mini":"Dry Sabji Mini",
    "Dry_Sabji_Full":"Dry Sabji Full","Curry_Sabji_Mini":"Curry Sabji Mini","Curry_Sabji_Full":"Curry Sabji Full",
    "Dal":"Dal","Rice":"Rice","Salad":"Salad","Curd":"Curd"};
  var totalRev=0, totalPaid=0, totalDelivery=0, totalSurcharge=0, totalSmallFee=0;
  var custSet={}, dayMap={};
  var mealStats={Breakfast:{count:0,revenue:0},Lunch:{count:0,revenue:0},Dinner:{count:0,revenue:0}};
  var itemCounts={};
  rows.forEach(function(r) {
    var d=fmtDate(r.Order_Date), net=Number(r.Net_Total)||0;
    var delivery=Number(r.Delivery_Charge)||0;
    var food=Number(r.Food_Subtotal)||0;
    // Backfill surcharge for old rows where Inflation_Surcharge column was blank
    var surchargeRaw=Number(r.Inflation_Surcharge);
    var surcharge = (!isNaN(surchargeRaw) && surchargeRaw > 0) ? surchargeRaw : (food > 0 ? Math.ceil(food/20) : 0);
    // Small_Order_Fee: exact backfill using Option B (checks VIP, pickup, day threshold)
    var smallFee = calcSmallFee(r);
    var payStatus = String(r.Payment_Status || "").trim();
    totalRev+=net;
    totalDelivery+=delivery; totalSurcharge+=surcharge; totalSmallFee+=smallFee;
    if(payStatus==="Paid"||payStatus==="Wallet Paid"||payStatus==="Collected") totalPaid+=net;
    var ph=String(r.Phone||"").trim(); if(ph) custSet[ph]=true;
    var meal=String(r.Meal_Type||"");
    if(mealStats[meal]){mealStats[meal].count++;mealStats[meal].revenue+=net;}
    if(!dayMap[d]) dayMap[d]={orders:0,revenue:0,delivery:0,surcharge:0,smallFee:0};
    dayMap[d].orders++; dayMap[d].revenue+=net;
    dayMap[d].delivery+=delivery; dayMap[d].surcharge+=surcharge; dayMap[d].smallFee+=smallFee;
    if(meal==="Breakfast"){
      for(var n=1;n<=4;n++){var bi=String(r["BF_Item_"+n]||"").trim(),bq=Number(r["BF_Qty_"+n])||0;if(bi&&bq>0)itemCounts[bi]=(itemCounts[bi]||0)+bq;}
      var cu=Number(r.Curd)||0; if(cu>0)itemCounts["Curd"]=(itemCounts["Curd"]||0)+cu;
    } else {
      LUNCH_COLS.forEach(function(col){var q=Number(r[col])||0;if(q>0){var dn=COL_DISP[col]||col;itemCounts[dn]=(itemCounts[dn]||0)+q;}});
    }
  });
  var days=Object.keys(dayMap).sort().map(function(d){
    return{date:d,orders:dayMap[d].orders,revenue:Math.round(dayMap[d].revenue),
           delivery:Math.round(dayMap[d].delivery),surcharge:Math.round(dayMap[d].surcharge),smallFee:Math.round(dayMap[d].smallFee)};
  });
  var allItems=Object.keys(itemCounts).map(function(k){return{name:k,count:Math.round(itemCounts[k])};}).sort(function(a,b){return b.count-a.count;});
  var topItems=allItems.slice(0,15);
  Object.keys(mealStats).forEach(function(m){mealStats[m].revenue=Math.round(mealStats[m].revenue);});
  return {success:true,
    summary:{orders:rows.length,customers:Object.keys(custSet).length,revenue:Math.round(totalRev),
      paid:Math.round(totalPaid),pending:Math.round(totalRev-totalPaid),
      avgPerDay:days.length>0?Math.round(totalRev/days.length):0,
      delivery:Math.round(totalDelivery),surcharge:Math.round(totalSurcharge),smallFee:Math.round(totalSmallFee)},
    meals:mealStats,days:days,topItems:topItems,allItems:allItems,
    // Lets the admin UI show "Including X archived orders" so they know
    // the report pulled across archive files (which is slower than live-only).
    archived:{count: archivedCount, included: archivedCount > 0}};
}
function saveInventoryEntry(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var now  = new Date();
  var id   = "INV-" + Utilities.formatDate(now,"Asia/Kolkata","yyyyMMdd") + "-" + Math.floor(Math.random()*9000+1000);
  var hIdx = headerIndex(ws);
  var totalCols = Math.max(ws.getLastColumn(), INVENTORY_HEADERS.length);
  var row  = new Array(totalCols).fill("");
  var set  = function(col, val) { if (hIdx[col]) row[hIdx[col]-1] = val; };

  set("Entry_ID",   id);
  set("Date",       String(body.date || Utilities.formatDate(now,"Asia/Kolkata","yyyy-MM-dd")));
  set("Item",       String(body.item || "").trim());
  set("Unit",       String(body.unit || "kg"));
  set("Quantity",   Number(body.quantity) || 0);
  set("Price_Paid", Number(body.price)    || 0);
  set("Notes",      String(body.notes     || ""));
  set("Timestamp",  getISTTimestamp());

  ws.appendRow(row);
  return { success: true, id: id };
}
function getInventoryData(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var rows = getAllRows(ws);

  // Sort ascending by date for correct duration calculation
  rows.sort(function(a,b){ return String(a.Date).localeCompare(String(b.Date)); });

  // Group by item
  var byItem = {};
  rows.forEach(function(r) {
    var item = String(r.Item || "").trim();
    if (!item) return;
    if (!byItem[item]) byItem[item] = [];
    byItem[item].push({
      id:       String(r.Entry_ID || ""),
      date:     String(r.Date     || ""),
      unit:     String(r.Unit     || "kg"),
      qty:      Number(r.Quantity) || 0,
      price:    Number(r.Price_Paid) || 0,
      notes:    String(r.Notes    || ""),
      timestamp:String(r.Timestamp || "")
    });
  });

  // For each item, calculate durations between entries + consumption stats
  var items = [];
  Object.keys(byItem).sort().forEach(function(item) {
    var entries = byItem[item];
    var totalDays = 0, totalQty = 0, durationCount = 0;

    // Annotate each entry with how long it lasted (days until next purchase)
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      e.lasted_days = null;
      e.daily_rate  = null;
      if (i < entries.length - 1) {
        var d1 = new Date(e.date);
        var d2 = new Date(entries[i+1].date);
        var days = Math.round((d2 - d1) / 86400000);
        if (days > 0) {
          e.lasted_days = days;
          e.daily_rate  = Math.round((e.qty / days) * 100) / 100;
          totalDays += days;
          totalQty  += e.qty;
          durationCount++;
        }
      }
    }

    var avgDays       = durationCount > 0 ? Math.round(totalDays / durationCount) : null;
    var avgDailyRate  = (avgDays && totalQty) ? Math.round((totalQty / durationCount / avgDays) * 100) / 100 : null;
    var lastEntry     = entries[entries.length - 1];

    // Predict next purchase date
    var nextBuyDate = null;
    if (avgDays && lastEntry.date) {
      var d = new Date(lastEntry.date);
      d.setDate(d.getDate() + avgDays);
      nextBuyDate = Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    }

    items.push({
      item:          item,
      unit:          lastEntry.unit,
      entries:       entries.reverse(), // newest first for display
      entry_count:   entries.length,
      avg_days:      avgDays,
      avg_daily_rate:avgDailyRate,
      last_purchased:lastEntry.date,
      last_qty:      lastEntry.qty,
      next_buy_est:  nextBuyDate
    });
  });

  return { success: true, items: items };
}
function deleteInventoryEntry(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var hIdx = headerIndex(ws);
  var rows = ws.getDataRange().getValues();
  var idCol = (hIdx["Entry_ID"] || 1) - 1;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]).trim() === String(body.id || "").trim()) {
      ws.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Entry not found" };
}
// ── EXPENSE CUSTOM CATEGORIES ─────────────────────────────────────────────────
// Stored in Script Properties as JSON so no extra sheet is needed.
function getCustomExpenseCategories() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty("CUSTOM_EXP_CATS");
    return raw ? JSON.parse(raw) : {};
  } catch(e) { return {}; }
}
function saveCustomExpenseCategory(body) {
  var category = String(body.category || "").trim();
  var item     = String(body.item     || "").trim();
  if (!category) return { success:false, error:"Category required" };
  var cats = getCustomExpenseCategories();
  if (!cats[category]) cats[category] = [];
  if (item && !cats[category].includes(item)) cats[category].push(item);
  PropertiesService.getScriptProperties().setProperty("CUSTOM_EXP_CATS", JSON.stringify(cats));
  return { success:true, categories: cats };
}
function deleteCustomExpenseCategory(body) {
  var category = String(body.category || "").trim();
  var item     = String(body.item     || "").trim();
  var cats = getCustomExpenseCategories();
  if (item && cats[category]) {
    cats[category] = cats[category].filter(function(i){ return i !== item; });
  } else {
    delete cats[category];
  }
  PropertiesService.getScriptProperties().setProperty("CUSTOM_EXP_CATS", JSON.stringify(cats));
  return { success:true, categories: cats };
}
function saveExpense(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var hIdx = headerIndex(ws);
  var now  = new Date();
  var id   = "EXP-" + Utilities.formatDate(now, "Asia/Kolkata", "yyyyMMdd") + "-" + Math.floor(Math.random()*9000+1000);

  var totalCols = ws.getLastColumn();
  var row = new Array(totalCols).fill("");
  var set = function(col, val) { if (hIdx[col]) row[hIdx[col]-1] = val; };

  set("Expense_ID",   id);
  set("Date",         String(body.date || Utilities.formatDate(now,"Asia/Kolkata","yyyy-MM-dd")));
  set("Category",     String(body.category  || ""));
  set("Item",         String(body.item      || ""));
  set("Amount",       Number(body.amount)   || 0);
  set("Frequency",    String(body.frequency || "One-time"));
  set("Payment_Mode", String(body.payment_mode || "Cash"));
  set("Notes",        String(body.notes     || ""));
  set("Timestamp",    getISTTimestamp());

  ws.appendRow(row);
  return { success: true, id: id };
}
function deleteExpense(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var hIdx = headerIndex(ws);
  var rows = ws.getDataRange().getValues();
  var idCol = (hIdx["Expense_ID"] || 1) - 1;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]).trim() === String(body.id || "").trim()) {
      ws.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Expense not found" };
}
function getExpenses(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var rows = getAllRows(ws);
  var from = String(body.from || "");
  var to   = String(body.to   || "");
  var filtered = rows.filter(function(r) {
    var d = String(r.Date || "").trim();
    if (from && d < from) return false;
    if (to   && d > to)   return false;
    return true;
  });
  // Sort newest first
  filtered.sort(function(a,b){ return String(b.Date).localeCompare(String(a.Date)); });
  return {
    success: true,
    expenses: filtered.map(function(r) {
      return {
        id:           r.Expense_ID,
        date:         r.Date,
        category:     r.Category,
        item:         r.Item,
        amount:       Number(r.Amount) || 0,
        frequency:    r.Frequency,
        payment_mode: r.Payment_Mode,
        notes:        r.Notes,
        timestamp:    r.Timestamp
      };
    }),
    categories: EXPENSE_CATEGORIES
  };
}
function getExpenseAnalytics(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var rows = getAllRows(ws);
  var from = String(body.from || "");
  var to   = String(body.to   || "");

  var filtered = rows.filter(function(r) {
    var d = String(r.Date || "").trim();
    if (!d) return false;
    if (from && d < from) return false;
    if (to   && d > to)   return false;
    return true;
  });

  var total       = 0;
  var byCat       = {};   // category → total
  var byFreq      = {};   // frequency → total
  var byPayMode   = {};   // payment mode → total
  var byDay       = {};   // date → total
  var topItems    = {};   // item → total
  var monthlyFixed = 0;   // sum of Monthly-tagged expenses

  filtered.forEach(function(r) {
    var amt  = Number(r.Amount) || 0;
    var cat  = String(r.Category || "Other");
    var freq = String(r.Frequency || "One-time");
    var pm   = String(r.Payment_Mode || "Cash");
    var d    = String(r.Date || "").trim();
    var item = String(r.Item || "Other");

    total += amt;
    byCat[cat]     = (byCat[cat]     || 0) + amt;
    byFreq[freq]   = (byFreq[freq]   || 0) + amt;
    byPayMode[pm]  = (byPayMode[pm]  || 0) + amt;
    byDay[d]       = (byDay[d]       || 0) + amt;
    topItems[item] = (topItems[item] || 0) + amt;
    if (freq === "Monthly") monthlyFixed += amt;
  });

  var days = Object.keys(byDay).sort().map(function(d) {
    return { date: d, amount: Math.round(byDay[d]) };
  });

  var catArr = Object.keys(byCat).sort(function(a,b){ return byCat[b]-byCat[a]; }).map(function(c) {
    return { category: c, amount: Math.round(byCat[c]) };
  });

  var itemArr = Object.keys(topItems).sort(function(a,b){ return topItems[b]-topItems[a]; }).slice(0,10).map(function(i) {
    return { item: i, amount: Math.round(topItems[i]) };
  });

  return {
    success:      true,
    total:        Math.round(total),
    monthlyFixed: Math.round(monthlyFixed),
    count:        filtered.length,
    byCategory:   catArr,
    byFrequency:  byFreq,
    byPayMode:    byPayMode,
    byDay:        days,
    topItems:     itemArr,
    categories:   EXPENSE_CATEGORIES
  };
}
function logClientError(body) {
  try {
    var ss  = getSpreadsheet();
    var ws  = getOrCreateTab(ss, TAB_ERROR_LOG, ERROR_LOG_HEADERS);
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");

    // Extract known fields; stash everything else in Extra_JSON for debugging
    var known = ["phone","version","type","action","attempt","ms","msg","url"];
    var extra  = {};
    Object.keys(body).forEach(function(k) { if (known.indexOf(k) === -1) extra[k] = body[k]; });

    ws.appendRow([
      getISTTimestamp(),
      dateStr,
      String(body.phone    || "unknown"),
      String(body.version  || ""),
      String(body.type     || "error"),
      String(body.action   || "unknown"),
      Number(body.attempt  || 1),
      Number(body.ms       || 0),
      String(body.msg      || ""),
      String(body.url      || ""),
      Object.keys(extra).length ? JSON.stringify(extra) : ""
    ]);
    return { success: true };
  } catch(e) {
    return { success: false }; // never throw — this is logging only
  }
}
// ── CHURN REPORT ──────────────────────────────────────────────────────────────
function getChurnReport(sinceDate) {
  if (!sinceDate) return {success:false, error:"sinceDate required"};
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };
  // To detect churn we need every customer's MOST RECENT order — so we must
  // scan all archives too, not just the live sheet. Otherwise customers whose
  // last order was in an archived month would always appear "churned".
  var today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  var earliest = "2024-01-01"; // conservative — covers everything ever archived
  var allRows = getOrdersInRangeWithArchive(earliest, today);

  var map = {};
  allRows.forEach(function(r) {
    if (_isOrderCancelled(r.Payment_Status)) return;
    var phone = String(r.Phone||"").trim();
    if (!phone) return;
    var d = fmtDate(r.Order_Date);
    if (!map[phone]) map[phone] = {
      phone:phone,
      name:String(r.Customer_Name||"").trim(),
      area:String(r.Area||"").trim(),
      lastDate:"",
      orderCount:0
    };
    map[phone].orderCount++;
    if (d > map[phone].lastDate) {
      map[phone].lastDate = d;
      map[phone].name = String(r.Customer_Name||map[phone].name).trim();
    }
  });
  var churned = Object.values(map).filter(function(c) { return c.lastDate < sinceDate; })
                                  .sort(function(a,b) { return b.lastDate.localeCompare(a.lastDate); });
  return {success:true, sinceDate:sinceDate, customers:churned, count:churned.length, archive_inclusive:true};
}
