// ============================================================
// 03_Customer.gs
// Customer registration, profile, order history, customer list,
// reviews/ratings, standard-order shortcuts.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── GET CUSTOMER ─────────────────────────────────────────────
function getCustomer(phone) {
  if (!phone) return {found: false};
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const rows = getAllRows(ws);
  const pStr = _normalizePhone(phone);
  const r = rows.find(x => _normalizePhone(x.Phone) === pStr);
  if (!r) return {found: false, hasPin: false, wallet_balance: 0};
  
  const hasPin = (String(r.PIN || "").trim() !== "");
  
  if (hasPin) {
    // Return early without profile details to secure them.
    return { found: true, hasPin: true };
  }
  
  return {
    found: true,
    hasPin: false,
    name:               r.Customer_Name || "",
    area:               r.Area || "",
    wing:               r.Wing || "",
    flat:               r.Flat || "",
    floor:              r.Floor || "",
    society:            r.Society || "",
    maps:               r.Maps_Link || "",
    landmark:           r.Landmark || "",
    payment_preference: r.Payment_Freq || "Daily Payment",
    meal_addresses:     r.Meal_Addresses || "",
    promoCount: (function(v){
      if (v === "" || v === null || v === undefined) return null;
      var num = Number(v);
      return isNaN(num) ? v : num;
    })(r.Review_Promo_Count),
    wallet_balance:     _calculateWalletBalance(phone),
    feeExempt:          (r.Fee_Exempt === "Yes" || r.Fee_Exempt === true),
    onAccount:          r.On_Account || "No",
    billingCycle:       r.Billing_Cycle || "Daily"
  };
}
// ── VERIFY LOGIN ─────────────────────────────────────────────
function verifyLogin(phone, pin) {
  if (!phone || !pin) return {success: false, error: "Missing Phone or PIN."};
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const rows = getAllRows(ws);
  const pStr = _normalizePhone(phone);
  const r = rows.find(x => _normalizePhone(x.Phone) === pStr);
  
  if (!r) return {success: false, error: "Account not found."};
  if (String(r.PIN).trim() !== String(pin).trim()) return {success: false, error: "Incorrect PIN."};
  
  let pendingAmount = 0;
  const isOnAccount = String(_get(r, "On_Account") || "").trim().toLowerCase() === "yes";
  
  if (isOnAccount) {
    const wsOrders = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
    const orderRows = getAllRows(wsOrders);
    for (const ord of orderRows) {
      if (_normalizePhone(_get(ord, "Phone")) === pStr) {
        const ps = String(_get(ord, "Payment_Status") || "").trim().toLowerCase();
        if (ps === "on account" || ps === "onaccount" || ps === "pending" || ps === "") {
          pendingAmount += _cleanNum(_get(ord, "Net_Total"));
        }
      }
    }
  }

  return {
    success: true,
    profile: {
      name:               r.Customer_Name || "",
      area:               r.Area || "",
      wing:               r.Wing || "",
      flat:               r.Flat || "",
      floor:              r.Floor || "",
      society:            r.Society || "",
      maps:               r.Maps_Link || "",
      landmark:           r.Landmark || "",
      payment_preference: r.Payment_Freq || "Daily Payment",
      meal_addresses:     r.Meal_Addresses || "",
      promoCount: (function(v){
        if (v === "" || v === null || v === undefined) return null;
        var num = Number(v);
        return isNaN(num) ? v : num;
      })(r.Review_Promo_Count),
      wallet_balance:     _calculateWalletBalance(phone),
      feeExempt:          (r.Fee_Exempt === "Yes" || r.Fee_Exempt === true),
      onAccount:          r.On_Account || "No",
      billingCycle:       r.Billing_Cycle || "Daily",
      pending_amount:     pendingAmount
    }
  };
}
// ── UPSERT CUSTOMER ──────────────────────────────────────────
function _upsertCustomer(ss, profile) {
  // Ensure tab exists and headers are correct before doing anything
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  SpreadsheetApp.flush(); // Lock in the headers before indexing

  const rows = getAllRows(ws);
  const pStr = _normalizePhone(profile.phone);
  const existing = rows.find(r => _normalizePhone(r.Phone) === pStr);
  
  const fullAddr = [
    profile.wing    && `Wing ${profile.wing}`,
    profile.flat    && `Flat ${profile.flat}`,
    profile.floor   && `${profile.floor} Floor`,
    profile.society, profile.area
  ].filter(Boolean).join(", ");

  if (existing) {
    const rowNum = existing._row;
    const hIdx = headerIndex(ws);
    const update = (col, val) => {
      // ONLY update if val is effectively provided (not undefined)
      if (hIdx[col] && val !== undefined) {
        ws.getRange(rowNum, hIdx[col]).setValue(val);
      }
    };
    if (profile.name !== undefined) update("Customer_Name", profile.name);
    if (profile.area !== undefined) update("Area",          profile.area);
    if (profile.wing !== undefined) update("Wing",          profile.wing);
    if (profile.flat !== undefined) update("Flat",          profile.flat);
    if (profile.floor !== undefined) update("Floor",         profile.floor);
    if (profile.society !== undefined) update("Society",       profile.society);
    if (profile.area !== undefined || profile.society !== undefined) update("Full_Address",  fullAddr);
    
    // Auto-derive Maps Link if missing
    let finalMaps = profile.maps || "";
    if (!finalMaps) {
      finalMaps = _deriveMapsLink(fullAddr, profile.society || "");
    }
    update("Maps_Link", finalMaps);

    if (profile.landmark !== undefined) update("Landmark",      profile.landmark || "");
    if (profile.delivery_point !== undefined) update("Delivery_Point", _getDeliveryPointLabel(profile.delivery_point));
    if (profile.payment_preference !== undefined) update("Payment_Freq",  profile.payment_preference);
    if (profile.pin) update("PIN", profile.pin);
    if (profile.meal_addresses) update("Meal_Addresses", profile.meal_addresses);
    if (profile.standardOrder !== undefined) update("Standard_Order", profile.standardOrder);
    if (profile.onAccount !== undefined) update("On_Account", profile.onAccount);
    if (profile.billingCycle !== undefined) update("Billing_Cycle", profile.billingCycle);
    
    SpreadsheetApp.flush(); // Ensure writes are committed
  } else {
    // For new records, construct a clean Row Array mapping directly to our schema
    const newRow = CUSTOMERS_HEADERS.map(h => {
      let val = "";
      switch(h) {
        case "Phone":           val = _normalizePhone(profile.phone); break;
        case "Customer_Name":   val = profile.name || ""; break;
        case "Area":            val = profile.area || ""; break;
        case "Wing":            val = profile.wing || ""; break;
        case "Flat":            val = profile.flat || ""; break;
        case "Floor":           val = profile.floor || ""; break;
        case "Society":         val = profile.society || ""; break;
        case "Full_Address":    val = fullAddr; break;
        case "Maps_Link":       val = profile.maps || _deriveMapsLink(fullAddr, profile.society || ""); break;
        case "Landmark":        val = profile.landmark || ""; break;
        case "Delivery_Point":  val = _getDeliveryPointLabel(profile.delivery_point); break;
        case "Payment_Freq":    val = profile.payment_preference || "Daily Payment"; break;
        case "Created_At":      val = getISTTimestamp(); break;
        case "PIN":             val = profile.pin || ""; break;
        case "Meal_Addresses":  val = profile.meal_addresses || ""; break;
        case "Standard_Order":  val = profile.standardOrder || ""; break;
        case "Billing_Cycle":   val = profile.billingCycle || "Daily"; break;
        case "On_Account":      val = profile.onAccount || "No"; break;
        case "Review_Promo_Count": val = ""; break;
        default:                val = "";
      }
      // Force leading zeros to be preserved for Phone and PIN by prepending '
      if (h === "Phone" || h === "PIN") return "'" + String(val).trim();
      return val;
    });
    
    // Safety check: Ensure we NEVER write to Row 1 (header row)
    const nextRow = Math.max(2, ws.getLastRow() + 1);
    ws.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);
  }
}
function getCustomerOrders(phone) {
  if (!phone) return {orders:[], past_orders:[], wallet_balance: 0};
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

  const fmtD = function(r) {
    return r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
  };

  const delWs = ss.getSheetByName("SK_Deliveries");
  const deliveryMap = {};
  if (delWs) {
    const delRows = getAllRows(delWs);
    delRows.forEach(d => {
      const sid = String(d.Submission_ID || "");
      if (sid) deliveryMap[sid] = {
        deliveredAt: d.Delivered_At || null,
        enRouteAt: d.EnRoute_At || null
      };
    });
  }

  const allFiltered = rows.filter(r => String(r.Phone).trim() === String(phone).trim());
  
  const upcoming = allFiltered
    .filter(r => fmtD(r) >= today)
    .sort((a,b) => fmtD(a).localeCompare(fmtD(b)))
    .map(r => {
      const delTracker = deliveryMap[String(r.Submission_ID)] || {};
      let itemsRaw = {};
      try { itemsRaw = JSON.parse(r.Items_JSON || "{}"); } catch(e) {}
      return {
        rowId:              r.Submission_ID,
        date:               fmtD(r),
        meal:               r.Meal_Type,
        summary:            _buildSummary(r),
        items_raw:          itemsRaw,
        total:              r.Net_Total,
        inflation_surcharge: Number(r.Inflation_Surcharge) || 0,
        loyalty_discount:   String(r.Loyalty_Discount || "").trim().toLowerCase() === "yes",
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
        wallet_credit:      Number(r.Wallet_Credit) || 0,
        deliveredAt:        delTracker.deliveredAt,
        enRouteAt:          delTracker.enRouteAt
      };
    });

  const past = allFiltered
    .filter(r => fmtD(r) < today)
    .sort((a,b) => fmtD(b).localeCompare(fmtD(a))) // newest first
    .slice(0, 10)
    .map(r => {
      const delTracker = deliveryMap[String(r.Submission_ID)] || {};
      // items_raw: structured {itemName: qty} for "Order Again" feature on frontend
      let itemsRaw = {};
      try { itemsRaw = JSON.parse(r.Items_JSON || "{}"); } catch(e) {}
      return {
        rowId:              r.Submission_ID,
        date:               fmtD(r),
        meal:               r.Meal_Type,
        summary:            _buildSummary(r),
        items_raw:          itemsRaw,
        total:              r.Net_Total,
        inflation_surcharge: Number(r.Inflation_Surcharge) || 0,
        loyalty_discount:   String(r.Loyalty_Discount || "").trim().toLowerCase() === "yes",
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
        wallet_credit:      Number(r.Wallet_Credit) || 0,
        deliveredAt:        delTracker.deliveredAt,
        enRouteAt:          delTracker.enRouteAt
      };
    });

  const onAccountBalance = allFiltered
    .filter(r => String(r.Payment_Status || "").toLowerCase() === "on account")
    .reduce((sum, r) => sum + (Number(r.Net_Total) || 0), 0);

  // ── Monthly spending summary ──────────────────────────────────
  // Compute current calendar month's total spend + order count.
  // Used by order.html to show "You spent ₹X in April across N orders".
  const nowIST = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM");
  let monthTotal = 0, monthCount = 0;
  allFiltered.forEach(r => {
    const d = fmtD(r);
    // Exclude cancelled orders from monthly spend summary
    if (d.startsWith(nowIST) && !_isOrderCancelled(r.Payment_Status)) {
      monthTotal += Number(r.Net_Total) || 0;
      monthCount++;
    }
  });
  const monthName = Utilities.formatDate(new Date(), "Asia/Kolkata", "MMMM");

  return {
    orders: upcoming,
    past_orders: past,
    wallet_balance: _calculateWalletBalance(phone),
    on_account_balance: onAccountBalance,
    month_summary: {
      month: monthName,
      total: monthTotal,
      count: monthCount
    }
  };
}
function _buildSummary(r) {
  try {
    const obj = JSON.parse(r.Items_JSON || "{}");
    return Object.entries(obj)
      .filter(([,q]) => q > 0)
      .map(([n,q]) => `${q}×${n}`)
      .join(", ") || "—";
  } catch(e) { return "—"; }
}
// ── GET CUSTOMER LIST ─────────────────────────────────────────────────────────
function getCustomerList() {
  var ss   = getSpreadsheet();
  var ordersWs = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  // Include archived orders so lifetime stats (order_count, total_spent)
  // don't reset every month after archiving. Live + archived merged.
  var today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  var ordRows = getOrdersInRangeWithArchive("2024-01-01", today);

  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var custRows = getAllRows(custWs);
  var cMap = {};
  custRows.forEach(function(c) {
    var p = _normalizePhone(c.Phone);
    if (p) {
      cMap[p] = {
        count: Number(c.Review_Promo_Count) || 0,
        claimed: (String(c.Review_Reward_Claimed) === "TRUE" || String(c.Review_Reward_Claimed) === "true"),
        standardOrder: c.Standard_Order || "",
        feeExempt: (String(c.Fee_Exempt).trim() === "Yes") ? "Yes" : "No",
        onAccount: (String(c.On_Account).trim() === "Yes") ? "Yes" : "No",
        billingCycle: c.Billing_Cycle || "Daily",
        // Address profiles
        wing:    c.Wing || "",
        flat:    c.Flat || "",
        floor:   c.Floor || "",
        society: c.Society || "",
        maps:    c.Maps_Link || "",
        landmark: c.Landmark || "",
        delivery_point: c.Delivery_Point || ""
      };
    }
  });

  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var map = {};
  ordRows.forEach(function(r) {
    if (_isOrderCancelled(r.Payment_Status)) return;
    var phone = String(r.Phone||"").trim();
    if (!phone) return;
    var normP = _normalizePhone(phone);
    var d = fmtDate(r.Order_Date);
    if (!map[phone]) {
      map[phone] = {
        phone:phone, 
        name:String(r.Customer_Name||"").trim(),
        area:String(r.Area||"").trim(), 
        payFreq:String(r.Payment_Freq||"").trim(),
        orderCount:0, 
        totalSpent:0, 
        pendingAmt:0, 
        lastDate:"",
        promoCount: cMap[normP] ? cMap[normP].count : 0,
        reviewClaimed: cMap[normP] ? cMap[normP].claimed : false,
        standardOrder: cMap[normP] ? cMap[normP].standardOrder : "",
        Fee_Exempt:    cMap[normP] ? cMap[normP].feeExempt : "No",
        onAccount:     cMap[normP] ? cMap[normP].onAccount : "No",
        billingCycle:  cMap[normP] ? cMap[normP].billingCycle : "Daily",
        wing:    cMap[normP] ? cMap[normP].wing : "",
        flat:    cMap[normP] ? cMap[normP].flat : "",
        floor:   cMap[normP] ? cMap[normP].floor : "",
        society: cMap[normP] ? cMap[normP].society : "",
        maps:    cMap[normP] ? cMap[normP].maps : "",
        landmark: cMap[normP] ? cMap[normP].landmark : "",
        delivery_point: cMap[normP] ? cMap[normP].delivery_point : ""
      };
    }
    map[phone].orderCount++;
    map[phone].totalSpent += Number(r.Net_Total)||0;
    const ps = String(r.Payment_Status || "").trim();
    if (ps !== "Paid" && ps !== "Wallet Paid" && ps !== "Collected") map[phone].pendingAmt += Number(r.Net_Total)||0;
    if (d > map[phone].lastDate) {
      map[phone].lastDate = d;
      map[phone].name = String(r.Customer_Name||map[phone].name).trim();
    }
  });

  // Also include SK_Customers entries that have never placed an order (e.g. pre-registered VIPs)
  custRows.forEach(function(c) {
    var p = String(c.Phone || "").trim();
    if (!p) return;
    var normP = _normalizePhone(p);
    if (map[p]) return; // already in map from orders
    // Only surface pre-registered VIPs (Fee_Exempt = Yes) to keep the list clean
    if (String(c.Fee_Exempt).trim() !== "Yes") return;
    map[p] = {
      phone: p,
      name: String(c.Customer_Name || "").trim(),
      area: String(c.Area || "").trim(),
      payFreq: String(c.Payment_Freq || "").trim(),
      orderCount: 0,
      totalSpent: 0,
      pendingAmt: 0,
      lastDate: String(c.Created_At instanceof Date
        ? Utilities.formatDate(c.Created_At, "Asia/Kolkata", "yyyy-MM-dd")
        : (c.Created_At || "")).slice(0, 10),
      promoCount: 0,
      reviewClaimed: false,
      standardOrder: "",
      Fee_Exempt: "Yes",
      onAccount: String(c.On_Account).trim() === "Yes" ? "Yes" : "No",
      billingCycle: c.Billing_Cycle || "Daily"
    };
  });

  // Load wallet rows ONCE — avoids N API calls inside the map loop
  var walletWs = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  var walletRows = getAllRows(walletWs);

  var customers = Object.values(map)
    .map(function(c){
      const wb = _calculateWalletBalance(c.phone, walletRows);
      const net = Math.round(c.pendingAmt - wb);
      return Object.assign({}, c, {
        totalSpent: Math.round(c.totalSpent),
        pendingAmt: net,
        walletBalance: Math.round(wb)
      });
    })
    .sort(function(a,b){return b.lastDate.localeCompare(a.lastDate);});

  return {success:true, customers:customers};
}
function markReviewed(body) {
  var phone = body.phone;
  if (!phone) return {success:false, error:"phone required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var hIdx = headerIndex(ws);
  var rows = getAllRows(ws);
  var normP = _normalizePhone(phone);

  var r = rows.find(function(x) { return _normalizePhone(x.Phone) === normP; });
  if (!r) return {success:false, error:"Customer not found"};

  var col = hIdx["Review_Promo_Count"];
  if (!col) return {success:false, error:"Review_Promo_Count column missing"};

  var current = Number(r.Review_Promo_Count) || 0;
  ws.getRange(r._row, col).setValue(current + 3);

  // Mark as claimed
  var claimCol = hIdx["Review_Reward_Claimed"];
  if (claimCol) {
    ws.getRange(r._row, claimCol).setValue("TRUE");
  }

  return {success:true, newCount: current + 3};
}
// ── GET CUSTOMER HISTORY ──────────────────────────────────────────────────────
function getCustomerHistory(phone) {
  if (!phone) return {success:false, error:"phone required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  // Pull live + archived orders for this customer so the timeline shows
  // their full history. Archive helper takes a date range — we use a
  // wide window (2020 through today + 1y) so every archive month is in scope.
  var todayPlusOneYear = new Date();
  todayPlusOneYear.setFullYear(todayPlusOneYear.getFullYear() + 1);
  var rangeTo = Utilities.formatDate(todayPlusOneYear, "Asia/Kolkata", "yyyy-MM-dd");
  var allRows;
  try {
    allRows = getOrdersInRangeWithArchive("2020-01-01", rangeTo) || [];
  } catch (e) {
    console.warn("getCustomerHistory: archive lookup failed, falling back to live: " + e.message);
    allRows = getAllRows(ws);
  }
  var rows = allRows.filter(function(r){return String(r.Phone||"").trim()===phone;});

  var orders = rows.map(function(r) {
    return {
      id:       r.Submission_ID,
      date:     fmtDate(r.Order_Date),
      meal:     r.Meal_Type,
      area:     r.Area,
      total:    Number(r.Net_Total)||0,
      subtotal: Number(r.Food_Subtotal)||0,
      delivery: Number(r.Delivery_Charge)||0,
      discount: Number(r.Discount_Amount)||0,
      status:   r.Payment_Status||"Pending",
      items:    r.Items_JSON,
      notes:    r.Special_Notes||""
    };
  }).sort(function(a,b){return b.date.localeCompare(a.date);});

  var name       = rows.length ? String(rows[0].Customer_Name||"").trim() : "";
  var area       = rows.length ? String(rows[0].Area||"").trim() : "";
  var payFreq    = rows.length ? String(rows[0].Payment_Freq||"").trim() : "";
  var activeOrders = orders.filter(function(o){return !_isOrderCancelled(o.status);});
  var totalSpent = Math.round(activeOrders.reduce(function(s,o){return s+o.total;},0));
  var pending    = Math.round(activeOrders.filter(function(o){return String(o.status)!=="Paid" && String(o.status)!=="Wallet Paid";}).reduce(function(s,o){return s+o.total;},0));

  // Fetch Standard_Order from customer sheet
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var cRows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var standardOrder = "";
  var custMatch = cRows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  if (custMatch) standardOrder = custMatch.Standard_Order || "";

  var feeExempt = custMatch ? (String(custMatch.Fee_Exempt).trim() === "Yes" ? "Yes" : "No") : "No";
  var onAccount = custMatch ? (String(custMatch.On_Account).trim() === "Yes" ? "Yes" : "No") : "No";
  var billingCycle = custMatch ? (custMatch.Billing_Cycle || "Daily") : "Daily";

  return {
    success:true, phone:phone, name:name, area:area, payFreq:payFreq,
    orders:orders, totalSpent:totalSpent, pending:pending, orderCount:orders.length,
    standardOrder: standardOrder, Fee_Exempt: feeExempt, On_Account: onAccount, Billing_Cycle: billingCycle,
    // Add full address profile
    wing:    custMatch ? (custMatch.Wing || "") : "",
    flat:    custMatch ? (custMatch.Flat || "") : "",
    floor:   custMatch ? (custMatch.Floor || "") : "",
    society: custMatch ? (custMatch.Society || "") : "",
    maps:    custMatch ? (custMatch.Maps_Link || "") : "",
    landmark: custMatch ? (custMatch.Landmark || "") : "",
    delivery_point: custMatch ? (custMatch.Delivery_Point || "") : ""
  };
}
// ═══════════════════════════════════════════════════════
// LIVE GOOGLE REVIEWS
// ═══════════════════════════════════════════════════════
function getReviews() {
  if (!GOOGLE_PLACES_API_KEY) {
    return { error: true, message: "Missing GOOGLE_PLACES_API_KEY. Please configure in Code.gs." };
  }
  
  try {
    var url = "https://maps.googleapis.com/maps/api/place/details/json?place_id=" + PLACE_ID + "&fields=url,rating,user_ratings_total,reviews&key=" + GOOGLE_PLACES_API_KEY;
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(res.getContentText());
    
    if (json.status !== "OK") {
      return { error: true, message: json.error_message || json.status };
    }
    
    var place = json.result;
    var liveReviews = [];
    
    if (place.reviews) {
      liveReviews = place.reviews.map(function(r) {
        return {
          name: r.author_name,
          rating: r.rating,
          date: r.relative_time_description, 
          text: r.text || ""
        };
      });
    }
    
    // Sort reviews randomly so they don't look repetitive
    liveReviews.sort(function() { return 0.5 - Math.random() });
    
    return {
      success: true,
      rating: place.rating || 5.0,
      total: place.user_ratings_total || 85,
      reviewUrl: place.url || 'https://g.page/r/CasEH8gGAhzLEBM/review',
      reviews: liveReviews
    };
  } catch(e) {
    return { error: true, message: e.message };
  }
}
// ── TEST DATA GENERATOR ──────────────────────────────────────
/**
 * ADMIN: Grant Review Promo (Manual)
 */
function markReviewed(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const rows = getAllRows(ws);
  const phone = _normalizePhone(body.phone);
  
  const hIdx = headerIndex(ws);
  if (!hIdx["Review_Promo_Count"]) return {success: false, error: "Review column not initialized. Please refresh sheet."};
  
  const rowIdx = rows.findIndex(x => _normalizePhone(x.Phone) === phone);
  if (rowIdx === -1) return {success: false, message: "Customer not found."};
  
  // Set Review_Promo_Count to 3
  const realRow = rowIdx + 2;
  ws.getRange(realRow, hIdx["Review_Promo_Count"]).setValue(3);
  
  return {success: true, message: "10% Discount (3x) gifted successfully!"};
}
function setStandardOrder(phone, itemsJSON, templateName, meal) {
  var ss = getSpreadsheet();
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var rows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var cust = rows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  
  var currentStr = (cust && cust.Standard_Order) ? cust.Standard_Order : "[]";
  var list = [];
  try { list = JSON.parse(currentStr); if(!Array.isArray(list)) list=[]; } catch(e){ list=[]; }
  
  // Ensure items is an object, not a string, before saving
  var finalItems = itemsJSON;
  if (typeof itemsJSON === "string") {
    try { finalItems = JSON.parse(itemsJSON); } catch(e) { finalItems = itemsJSON; }
  }
  
  // Remove existing with same name if any
  list = list.filter(function(x){ return x.name !== templateName; });
  list.push({ 
    name: templateName, 
    meal: meal || "Other",
    items: finalItems, 
    createdAt: new Date().toISOString() 
  });
  
  _upsertCustomer(ss, { phone: phone, standardOrder: JSON.stringify(list) });
  return { success: true };
}
function removeStandardOrder(phone, templateName) {
  var ss = getSpreadsheet();
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var rows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var cust = rows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  
  if (!cust || !cust.Standard_Order) return { success: true };
  
  var list = [];
  try { list = JSON.parse(cust.Standard_Order); if(!Array.isArray(list)) list=[]; } catch(e){ list=[]; }
  list = list.filter(function(x){ return x.name !== templateName; });
  
  _upsertCustomer(ss, { phone: phone, standardOrder: JSON.stringify(list) });
  return { success: true };
}
