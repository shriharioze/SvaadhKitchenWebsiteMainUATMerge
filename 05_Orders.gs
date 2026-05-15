// ============================================================
// 05_Orders.gs
// Order submission flow, loyalty streak, day totals, delete/cancel,
// manual order admin, missed-order safety net, retroactive credit logic.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── AUTO-SETTLE PENDING ORDERS ──────────────────────────────
function _autoSettlePendingOrders(phone) {
  const pStr = _normalizePhone(phone);
  
  const ss = getSpreadsheet();
  const profWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const profRows = getAllRows(profWs);
  const profile = profRows.find(r => _normalizePhone(r.Phone) === pStr);
  
  // Rule 1: Only for On Account users
  if (!profile || (String(profile.On_Account).trim().toLowerCase() !== "yes")) {
    return { settled: 0, msg: "" };
  }

  const wsOrders = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(wsOrders);
  const hIdx = headerIndex(wsOrders);

  // Rule 2: Only target "on account" orders (ignore normal Pending/UPI)
  const pendingOrders = rows.filter(r => {
    if (_normalizePhone(_get(r, "Phone")) !== pStr) return false;
    const ps = String(_get(r, "Payment_Status") || "").trim().toLowerCase();
    if (ps !== "on account" && ps !== "onaccount") return false;
    return _cleanNum(_get(r, "Net_Total")) > 0;
  });

  if (pendingOrders.length === 0) return { settled: 0, msg: "" };

  pendingOrders.sort((a, b) => String(_get(a, "Order_Date")).localeCompare(String(_get(b, "Order_Date"))));

  let walletBalance = _calculateWalletBalance(phone);
  if (walletBalance <= 0) return { settled: 0, msg: "" };

  let totalSettled = 0;
  let ordersSettledCount = 0;
  let originalPendingAmount = pendingOrders.reduce((sum, o) => sum + _cleanNum(_get(o, "Net_Total")), 0);
  
  let currentWallet = walletBalance;

  for (let order of pendingOrders) {
    let amount = _cleanNum(_get(order, "Net_Total"));
    if (currentWallet >= amount) {
      wsOrders.getRange(order._row, hIdx["Payment_Status"]).setValue("Paid");
      _appendWalletTransaction(phone, _get(order, "Customer_Name") || "Customer", "Auto-deducted for On Account order " + (_get(order, "Submission_ID") || _get(order, "Order_Date")), amount, true, "AUTO-" + Date.now() + "-" + Math.floor(Math.random()*1000));
      currentWallet -= amount;
      totalSettled += amount;
      ordersSettledCount++;
    } else {
      break;
    }
  }

  if (ordersSettledCount > 0) {
    if (originalPendingAmount <= walletBalance) {
      return { 
        settled: totalSettled, 
        msg: `Wallet recharge used against the pending orders. Balance is now: Wallet ₹${currentWallet}` 
      };
    } else {
      // Wallet < Pending overall
      return { 
        settled: totalSettled, 
        msg: `Wallet recharge applied! Note: ₹${originalPendingAmount - totalSettled} is still pending on account.` 
      };
    }
  }

  // If we couldn't settle even one full order but they have wallet balance
  if (originalPendingAmount > 0 && walletBalance > 0 && walletBalance < originalPendingAmount) {
    return {
      settled: 0,
      msg: `Recharge added to wallet (₹${walletBalance}). You still have ₹${originalPendingAmount} pending on account.`
    };
  }

  return { settled: 0, msg: "" };
}
// ── MISSED-ORDER SAFETY NET ───────────────────────────────────
/**
 * Called immediately after appendRow for each order row.
 * Saves the order payload to Script Properties as a backup.
 * A separate cleanup pass (called at the end of submitOrder after flush)
 * verifies the row landed in the sheet; if not, it emails admin.
 *
 * This closes the 0.5% gap where GAS buffered writes silently failed.
 */
function _missedOrderSafetyNet(ss, sid, row, phone) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const raw    = props.getProperty("PENDING_ORDER_ROWS") || "{}";
    const store  = JSON.parse(raw);
    // Expire entries older than 10 minutes
    const now    = Date.now();
    Object.keys(store).forEach(k => { if (now - store[k].ts > 10 * 60 * 1000) delete store[k]; });
    store[sid]   = { ts: now, phone: String(phone || ""), row: row };
    props.setProperty("PENDING_ORDER_ROWS", JSON.stringify(store));
  } catch(e) {
    console.error("_missedOrderSafetyNet save failed:", e.message);
  }
}
function _verifyAndAlertMissedOrders(ss, submissionIds) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const raw    = props.getProperty("PENDING_ORDER_ROWS") || "{}";
    const store  = JSON.parse(raw);
    if (!Object.keys(store).length) return;

    const ws     = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
    const hIdx   = headerIndex(ws);
    const sidCol = hIdx["Submission_ID"];
    if (!sidCol) return;

    // Read all Submission_IDs from sheet (last 200 rows for speed)
    const lastRow  = ws.getLastRow();
    const startRow = Math.max(2, lastRow - 200);
    const count    = lastRow - startRow + 1;
    if (count <= 0) return;
    const sidValues = ws.getRange(startRow, sidCol, count, 1).getValues().flat().map(String);
    const inSheet   = new Set(sidValues);

    const missed = [];
    Object.entries(store).forEach(([sid, entry]) => {
      if (!inSheet.has(sid)) {
        console.error("MISSED ORDER DETECTED — " + sid + " not found in sheet after flush!");
        missed.push({ sid, phone: entry.phone, row: entry.row });
        // Emergency re-append
        try {
          ws.appendRow(entry.row);
          console.log("Emergency re-append succeeded for " + sid);
        } catch(e2) {
          console.error("Emergency re-append FAILED for " + sid + ": " + e2.message);
        }
      }
      delete store[sid]; // clear from queue regardless
    });

    props.setProperty("PENDING_ORDER_ROWS", JSON.stringify(store));

    if (missed.length > 0) {
      // Email admin alert
      try {
        const adminEmail = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL");
        if (adminEmail) {
          const body = missed.map(m =>
            `SK Order ID: ${m.sid}\nPhone: ${m.phone}\nRow data: ${JSON.stringify(m.row)}`
          ).join("\n\n---\n\n");
          GmailApp.sendEmail(adminEmail, "⚠️ Svaadh: Missed Order Row Detected & Auto-Recovered", body);
        }
      } catch(e) { console.error("Alert email failed:", e.message); }
    }
  } catch(e) {
    console.error("_verifyAndAlertMissedOrders failed:", e.message);
  }
}
// ── SUBMIT ORDER ─────────────────────────────────────────────
function submitOrder(body) {
  // Serialize submitOrder calls to prevent stock-race + wallet-race between concurrent customers.
  // ALSO: enforce Gateway_Order_ID idempotency before any writes happen — if a row
  // already exists in SK_Orders for this HDFC order, return the existing Submission_IDs
  // without writing again. This is the fix for the duplicate-order bug found in UAT
  // (SK-20260513-6512 / SK-20260513-6666 — same Gateway_Order_ID, 1 second apart, both
  // rows landed because submitOrder had no gateway-level dedup).
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch(e) { return { error: "Server busy — please retry in a few seconds." }; }
  try {
    const gatewayOrderIdGuard = String(body.gateway_order_id || "").trim();
    if (gatewayOrderIdGuard) {
      const ss       = getSpreadsheet();
      const ordersWs = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
      const data     = ordersWs.getDataRange().getValues();
      const headers  = data[0] || [];
      const gCol     = headers.indexOf("Gateway_Order_ID");
      const sidCol   = headers.indexOf("Submission_ID");
      if (gCol !== -1 && sidCol !== -1) {
        const existingSids = [];
        for (let r = 1; r < data.length; r++) {
          if (String(data[r][gCol] || "").trim() === gatewayOrderIdGuard) {
            existingSids.push(String(data[r][sidCol] || "").trim());
          }
        }
        if (existingSids.length > 0) {
          console.log("submitOrder: idempotent skip — Gateway_Order_ID "
            + gatewayOrderIdGuard + " already has " + existingSids.length
            + " row(s): " + existingSids.join(", "));
          return {
            success: true,
            idempotent: true,
            submissionIds: existingSids,
            submissionId:  existingSids[0],
            message: "Order already recorded for this gateway transaction."
          };
        }
      }
    }
    return _submitOrderInternal(body);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}
function _submitOrderInternal(body) {
  const ss = getSpreadsheet();
  const ordersWs = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const profile   = body.profile || {};
  const orders    = body.orders  || [];   // [{date, meals:[{type,items,notes,subtotal,area}]}]

  const submittedAt  = getISTTimestamp();
  let   payMethod    = body.payment_method  || "UPI";
  let   payStatus    = body.payment_status  || "Pending";
  const firstTime    = profile.isFirstTime ? "Yes" : "No";
  const payFreq      = profile.payment_preference || "Daily Payment";

  // Build the header→index map once
  const hIdx = headerIndex(ordersWs);

  // Fetch free areas dynamically (replaces hardcoded FREE_AREA = "Bhosale Nagar")
  const freeAreaNames = getAreas().filter(function(a){ return a.free; }).map(function(a){ return a.name; });
  const DELIVERY  = 10;

  const submissionIds = [];

  // ── ONE-SHOT ROW FETCHES ────────────────────────────────────
  // Fetch once, share everywhere. Previously these tabs were re-read 5+ times
  // per submitOrder (day totals, loyalty, duplicate check, stock check, wallet).
  const allOrderRows  = getAllRows(ordersWs);
  const walletWsRef   = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const allWalletRows = getAllRows(walletWsRef);
  // Menu rows read once here — reused by stock check below (avoids duplicate sheet fetch)
  const menuWsOnce  = getOrCreateTab(ss, TAB_MENU, []);
  const menuRowsAll = getAllRows(menuWsOnce);

  // Fetch existing orders once for all dates in this submission to calculate combined-day fees/discounts
  const submissionDates = orders.map(o => o.date);
  const existingDayTotals = getDayTotalsForDates(profile.phone, submissionDates.join(','), allOrderRows).dayTotals || {};

  // Fetch current promo state
  const custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const cIdx   = headerIndex(custWs);
  const cRows  = getAllRows(custWs);
  const phoneStr = _normalizePhone(profile.phone);
  const cRowIdx = cRows.findIndex(r => _normalizePhone(r.Phone) === phoneStr);
  let promoCount = null;
  if (cRowIdx !== -1) {
    const rawVal = cRows[cRowIdx].Review_Promo_Count;
    promoCount = (rawVal === "" || rawVal === undefined) ? null : rawVal;
    if (promoCount !== null && !isNaN(promoCount)) promoCount = Number(promoCount);

    // ── On Account override (server-enforced) ──────────────────
    // If the customer is flagged On_Account in SK_Customers, every order
    // is automatically set to method "On Account" / status "On Account"
    // regardless of what the frontend sends.
    if (String(cRows[cRowIdx].On_Account || "").trim() === "Yes") {
      payMethod = "On Account";
      payStatus = "On Account";
    }
  }

  // Pre-fetch masters once for ID -> Name resolution in sheet columns
  const masterMap = {};
  try {
    const masters = getAdminData();
    (masters.breakfastMaster || []).forEach(m => masterMap[String(m.id)] = m.name);
    (masters.sabjiMaster || []).forEach(m => masterMap[String(m.id)] = m.name);
  } catch(e) { console.error("Master fetch failed in submitOrder", e); }

  // Strip weight/measure suffixes like [175g], [200g], [100ml], (2 pieces) etc.
  // so backend always stores the clean item name regardless of what frontend shows.
  const stripDisplaySuffix = (name) => {
    return String(name)
      .replace(/\s*\[.*?\]\s*/g, '')   // removes [175g], [200ml], [2 pcs] etc.
      .replace(/\s*\(.*?\)\s*/g, '')   // removes (2 pieces), (100ml) etc.
      .trim();
  };

  const resolveName = (k) => {
    let name;
    if (ITEM_COL_MAP[k]) name = ITEM_COL_MAP[k].replace(/_/g, ' ');
    else if (masterMap[k]) name = masterMap[k];
    else name = k.replace(/_/g, ' ');
    return stripDisplaySuffix(name);
  };

  // Sort orders by date to ensure virtual streak runs chronologically
  orders.sort((a,b) => a.date.localeCompare(b.date));
  const initialStreakInfo = _calculateLoyaltyStreak(profile.phone, allOrderRows);
  let virtualStreakCount = initialStreakInfo.streak;
  let virtualPastSurcharge = initialStreakInfo.pastSurcharge;

  // ════ STOCK LIMIT PRE-FLIGHT ════
  // Hard-block submission if any requested item exceeds remaining stock.
  // Runs under LockService so concurrent submissions see each other's counts.
  {
    const menuRowsStk = menuRowsAll;   // reuse the already-fetched menu rows
    const stockConflicts = [];
    for (const dateOrder of orders) {
      const dateStrStk = dateOrder.date;
      const menuRowStk = menuRowsStk.find(mr => {
        const d = mr.Date instanceof Date
          ? Utilities.formatDate(mr.Date, "Asia/Kolkata", "yyyy-MM-dd")
          : String(mr.Date).trim();
        return d === dateStrStk;
      });
      let stockLimitsStk = {};
      try { if (menuRowStk && menuRowStk.Stock_JSON) stockLimitsStk = JSON.parse(menuRowStk.Stock_JSON); } catch(e) {}
      if (!Object.keys(stockLimitsStk).length) continue;

      const countedStk = countOrderedUnits(allOrderRows, dateStrStk);
      for (const mealStk of (dateOrder.meals || [])) {
        const mealLimits = stockLimitsStk[mealStk.type] || {};
        let mealItems = mealStk.items || [];
        if (typeof mealItems === "string") {
          try { mealItems = JSON.parse(mealItems); } catch(e) { mealItems = []; }
        }
        if (!Array.isArray(mealItems)) mealItems = [];
        for (const it of mealItems) {
          const colKeyStk = it.colKey;
          const qtyStk = Number(it.qty) || 0;
          if (qtyStk <= 0) continue;
          const limitStk = mealLimits[colKeyStk];
          if (limitStk === undefined) continue;
          const usedStk = countedStk[mealStk.type][itemsJsonKey(colKeyStk)] || 0;
          if (usedStk + qtyStk > limitStk) {
            stockConflicts.push({
              date: dateStrStk,
              meal: mealStk.type,
              colKey: colKeyStk,
              available: Math.max(0, limitStk - usedStk)
            });
          }
        }
      }
    }
    if (stockConflicts.length) {
      const first = stockConflicts[0];
      const nm = first.colKey === "B_CURD" ? "Curd (Breakfast)" : first.colKey;
      return {
        error: `Only ${first.available} of "${nm}" left for ${first.meal} on ${first.date}. Please reduce your quantity.`,
        stock_conflicts: stockConflicts
      };
    }
  }

  const _dupNowMs = Date.now();
  const _FIVE_MIN_MS = 5 * 60 * 1000;
  const _normPhone = _normalizePhone(profile.phone);
  let loyaltyExcessCredit = 0; // accumulates surplus when 6th-day discount exceeds the bill
  // Normalize an items object to a stable JSON signature (sorted keys)
  const _itemsSig = (obj) => JSON.stringify(
    Object.keys(obj).sort().reduce((a, k) => { a[k] = obj[k]; return a; }, {})
  );
  // Normalize a date value that may be a Date object or a string
  const _normDate = (d) => {
    if (!d) return "";
    if (d instanceof Date) return Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    return String(d).trim().substring(0, 10);
  };

  for (const order of orders) {
    const orderDate = order.date;
    const is6thDay = (virtualStreakCount === 5); // Hits 6 on this day
    const existingDateInfo = (existingDayTotals[orderDate] || {});

    // Calculate meal count for this date to determine dynamic free delivery threshold
    const mealsThisSubmission = order.meals.filter(m => (Number(m.subtotal) || 0) > 0).map(m => m.type);
    const existingMeals = Object.keys(existingDateInfo).filter(mType => (Number(existingDateInfo[mType].subtotal) || 0) > 0);
    const allMealsOnDate = Array.from(new Set([...mealsThisSubmission, ...existingMeals]));
    const totalMealsCount = allMealsOnDate.length;
    const dynamicFreeThreshold = totalMealsCount <= 1 ? 100 : 150;

    // Calculate total food subtotal for this specific submission's date
    const submissionDayFoodTotal = order.meals.reduce((s, m) => s + (Number(m.subtotal) || 0), 0);
    // Combine with already placed orders for this date
    const prevDayFoodTotal = Object.values(existingDateInfo).reduce((s, m) => s + (Number(m.subtotal) || 0), 0);
    const combinedDayTotal = submissionDayFoodTotal + prevDayFoodTotal;

    // Calculate day-level discount once across all meals for this date (including previous ones)
    let discRate = 0;
    if (combinedDayTotal >= 450) discRate = 0.075;
    else if (combinedDayTotal >= 300) discRate = 0.05;
    
    const totalDayDiscAmt = Math.round(combinedDayTotal * discRate);
    // Find how much discount was already applied to previous orders for this date
    const prevDayDiscAmt = Object.values(existingDateInfo).reduce((s, m) => s + (Number(m.discount_applied) || 0), 0);
    // The discount to apply to this ENTIRE submission for this date = entitled - already_applied
    const submissionDateDiscAmt = Math.max(0, totalDayDiscAmt - prevDayDiscAmt);

    // Pro-rate the submission-level discount across meals in this submission
    const getDisc = (sub) => {
      if (is6thDay) {
        // Loyalty Discount: Waive all 6 days of surcharge
        const currentSurcharge = Math.ceil(submissionDayFoodTotal / 20);
        const totalWaiver = virtualPastSurcharge + currentSurcharge;
        return submissionDayFoodTotal > 0 ? Math.round(totalWaiver * (sub / submissionDayFoodTotal)) : 0;
      }
      return submissionDayFoodTotal > 0 ? Math.round(submissionDateDiscAmt * (sub / submissionDayFoodTotal)) : 0;
    };

    // Update virtual streak state for NEXT loop iteration
    const currentDaySurcharge = Math.ceil(submissionDayFoodTotal / 20);
    if (is6thDay) {
      virtualStreakCount = 0;
      virtualPastSurcharge = 0;
    } else {
      virtualStreakCount++;
      virtualPastSurcharge += currentDaySurcharge;
    }

    for (const meal of order.meals) {
      const sid = generateSubmissionID();
      submissionIds.push(sid);
      meal._sid = sid; // carry sid for ledger
      
      const mealType = meal.type;
      const sub = Number(meal.subtotal) || 0;
      const mealArea = meal.area || profile.area || "";
      
      let items  = meal.items || [];   // [{colKey, qty}]
      // Safety fix: If items is a stringified JSON, parse it (prevents character-distortion crash)
      if (typeof items === "string") {
        try { items = JSON.parse(items); } catch(e) { items = []; }
      }
      if (!Array.isArray(items)) items = [];

      const nKitchen = meal.notesKitchen || "";
      const nDelivery = meal.notesDelivery || "";
      
      // Get combined totals for THIS specific meal type (prev + current)
      const prevMealSub = (existingDateInfo[mealType] || {}).subtotal || 0;
      const combinedMealSub = sub + prevMealSub;
      
      // Delivery & Fee logic (matches frontend)
      const isPickup  = (mealArea.toLowerCase().includes("pickup"));
      const isDayFree = (combinedDayTotal >= dynamicFreeThreshold);
      const isFreeArea = freeAreaNames.includes(mealArea);

      // VIP Fee Exemption
      const isFeeExempt = (cRowIdx !== -1 && (cRows[cRowIdx].Fee_Exempt === "Yes" || cRows[cRowIdx].Fee_Exempt === true));

      let delCharge = 0;
      if (!isFeeExempt && !isDayFree && !isPickup && !isFreeArea && sub > 0) {
        delCharge = DELIVERY;
      }

      let smallOrderFee = 0;
      if (!isFeeExempt && !isDayFree && !isPickup && (mealType === "Lunch" || mealType === "Dinner") && sub > 0 && combinedMealSub < 50) {
        smallOrderFee = 10;
      }

      // Calculation of credits for previously paid fees on the same day (Retroactive waiver).
      //
      // BUG-FIX (UAT v14.8 — SK-20260513-5042 / 7492 / 9482 forensics):
      // Previously we summed prior delivery_charged + small_fee_charged every time
      // isDayFree triggered, so the SAME ₹X in prior fees got refunded multiple
      // times (once per subsequent meal of the day). Net effect: customer received
      // many copies of the same refund.
      //
      // Fix: also sum the prior Meal_Credit already given, and subtract from the
      // available pool so each rupee of prior fee gets refunded EXACTLY ONCE
      // across the day.
      //
      //   available_credit = prior_delivery + prior_small_fee − prior_meal_credit_already_given
      //
      // getDayTotalsForDates surfaces this as meal_credit_applied per existing row.
      // Existing rows without this column read as 0, preserving backward compat.
      let dateDeliveryCredit = 0;
      let dateSmallFeeCredit = 0;
      let dateMealCreditAlreadyGiven = 0;
      if (isDayFree) {
        Object.keys(existingDateInfo).forEach(mType => {
          dateDeliveryCredit         += (Number(existingDateInfo[mType].delivery_charged)    || 0);
          dateSmallFeeCredit         += (Number(existingDateInfo[mType].small_fee_charged)   || 0);
          dateMealCreditAlreadyGiven += (Number(existingDateInfo[mType].meal_credit_applied) || 0);
        });
      }
      const totalPriorFees      = dateDeliveryCredit + dateSmallFeeCredit;
      const availableDateCredit = Math.max(0, totalPriorFees - dateMealCreditAlreadyGiven);
      const mealCredit = submissionDayFoodTotal > 0
        ? Math.round(availableDateCredit * (sub / submissionDayFoodTotal))
        : 0;

      const discAmt = getDisc(sub);
      // On the 6th day, the surcharge IS charged (consistency), and the loyalty discount
      // (totalWaiver) includes all 6 days of surcharge so it covers it. Net effect:
      // the 6th-day surcharge charge and refund cancel each other; customer gets back days 1–5.
      const inflationSurcharge = Math.ceil(sub / 20);

      // Google Review Promo Logic (10% OFF per meal)
      let reviewDiscount = 0;
      const isNumeric = (typeof promoCount === "number" && !isNaN(promoCount));
      if (isNumeric && promoCount > 0 && sub > 0) {
        reviewDiscount = Math.round(sub * 0.10);
        promoCount--;
      }

      let netTotal = Math.round(sub + delCharge + smallOrderFee + inflationSurcharge - discAmt - mealCredit - reviewDiscount);
      // If the 6th-day loyalty discount exceeds this meal's bill, clamp to ₹0 and
      // accumulate the surplus — it gets credited to the customer's wallet after all rows are written.
      if (is6thDay && netTotal < 0) {
        loyaltyExcessCredit += Math.abs(netTotal);
        netTotal = 0;
      }
      meal._reviewDiscount = reviewDiscount; // carry for set() below


      // Build items JSON
      // Breakfast Curd gets a distinct key ("Breakfast Curd") so kitchen prep
      // and admin reports can tell it apart from Lunch/Dinner Curd.
      const itemsObj = {};
      items.forEach(({colKey, qty}) => {
        let canonical;
        if (colKey === "B_CURD") canonical = "Breakfast Curd";
        else canonical = resolveName(colKey);
        itemsObj[canonical] = qty;
      });

      // Address fields handling (Sanitized for Pickup)
      const wing    = isPickup ? "" : (meal.wing    || profile.wing    || "");
      const flat    = isPickup ? "" : (meal.flat    || profile.flat    || "");
      const floor   = isPickup ? "" : (meal.floor   || profile.floor   || "");
      const society = isPickup ? "" : (meal.society || profile.society || "");
      const area    = isPickup ? "Self Pickup" : mealArea;

      const fullAddr = isPickup
                        ? "Self Pickup (A 104, Shree laxmi vihar society)"
                        : [wing && `Wing ${wing}`, flat && `Flat ${flat}`, floor && `${floor} Floor`, society, area].filter(Boolean).join(", ");
      const mapsLink = isPickup ? "" : (meal.maps || profile.maps || "");
      const landmark = isPickup ? "" : (meal.landmark || profile.landmark || "");

      // Build row array aligned to ORDERS_HEADERS
      const row = new Array(ORDERS_HEADERS.length).fill("");
      const set = (colName, val) => {
        const idx = hIdx[colName];
        if (idx) row[idx - 1] = val;
      };

      set("Submission_ID",       sid);
      set("Submitted_At",        submittedAt);
      set("Order_Date",          orderDate);
      set("Meal_Type",           mealType);
      set("Customer_Name",       profile.name     || "");
      set("Phone",               _normalizePhone(profile.phone));
      set("Area",                area);
      set("Wing",                wing);
      set("Flat",                flat);
      set("Floor",               floor);
      set("Society",             society);
      set("Full_Address",        fullAddr);
      set("Maps_Link",           mapsLink);
      set("Landmark",            landmark);
      set("Delivery_Point",      _getDeliveryPointLabel(meal.delivery_point || profile.delivery_point));
      if (!hIdx["Small_Order_Fee"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Small_Order_Fee");
        hIdx["Small_Order_Fee"] = ordersWs.getLastColumn();
      }
      if (!hIdx["Inflation_Surcharge"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Inflation_Surcharge");
        hIdx["Inflation_Surcharge"] = ordersWs.getLastColumn();
      }
      if (!hIdx["Loyalty_Discount"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Loyalty_Discount");
        hIdx["Loyalty_Discount"] = ordersWs.getLastColumn();
      }
      // Auto-heal Meal_Credit column (added v14.8 to fix double-refund bug).
      // Tracks retroactive delivery/small-fee refunds so each rupee of prior
      // fee gets credited back EXACTLY ONCE across the day's meals.
      if (!hIdx["Meal_Credit"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Meal_Credit");
        hIdx["Meal_Credit"] = ordersWs.getLastColumn();
      }
      set("Items_JSON",          JSON.stringify(itemsObj));
      set("Special_Notes_Kitchen",  nKitchen);
      set("Special_Notes_Delivery", nDelivery);
      set("Food_Subtotal",       sub);
      set("Small_Order_Fee",     smallOrderFee);
      set("Inflation_Surcharge", inflationSurcharge);
      set("Loyalty_Discount",    is6thDay ? "Yes" : "No");
      set("Delivery_Charge",     delCharge);
      set("Discount_Amount",     discAmt);
      set("Meal_Credit",         mealCredit);
      if (hIdx["Review_Discount"]) {
        set("Review_Discount",   meal._reviewDiscount || 0);
      }
      set("Net_Total",           netTotal);

      // ════ DUPLICATE CHECK — must run BEFORE wallet deduction ════
      // Three layers of protection (any one catching is enough):
      //
      //   1. CacheService (fast, in-memory, atomic across script invocations).
      //      Bulletproof against sheet-read staleness — if A wrote here in the
      //      last 5 min, B's cache.get(key) will see it instantly even if
      //      A's appendRow hasn't propagated to a fresh getAllRows yet.
      //   2. Fresh sheet re-read (catches anything cache evicted under load).
      //   3. The original allOrderRows snapshot (legacy, kept for safety).
      //
      // After the row is written, we cache.put(key) so future calls hit layer 1.
      const _incomingSig = _itemsSig(itemsObj);
      const _dupKey      = `dup_${_normPhone}_${_normDate(orderDate)}_${mealType}_${_incomingSig}`;
      const _cache       = CacheService.getScriptCache();

      // Layer 1: cache lookup
      const _cachedSid = _cache.get(_dupKey);
      if (_cachedSid) {
        submissionIds[submissionIds.length - 1] = _cachedSid;
        console.log("Duplicate caught by cache: " + _dupKey + " → " + _cachedSid);
        continue;
      }

      // Layer 2 + 3: fresh sheet re-read (covers cache eviction edge cases)
      const _freshRows  = getAllRows(ordersWs);
      const _nowMsFresh = Date.now();
      const _dupRow = _freshRows.find(r => {
        if (_normalizePhone(r.Phone) !== _normPhone) return false;
        if (_normDate(r.Order_Date) !== _normDate(orderDate)) return false;
        if (r.Meal_Type !== mealType) return false;
        const rMs = r.Submitted_At ? new Date(r.Submitted_At).getTime() : 0;
        if (!rMs || (_nowMsFresh - rMs) > _FIVE_MIN_MS) return false;
        try {
          const stored = typeof r.Items_JSON === "string" ? JSON.parse(r.Items_JSON) : (r.Items_JSON || {});
          return _itemsSig(stored) === _incomingSig;
        } catch(e) { return false; }
      });
      if (_dupRow) {
        submissionIds[submissionIds.length - 1] = _dupRow.Submission_ID || sid;
        // Backfill cache so subsequent calls hit layer 1 (faster + more reliable)
        try { _cache.put(_dupKey, _dupRow.Submission_ID || sid, 300); } catch(e) {}
        console.log("Duplicate order skipped (sheet check): " + _normPhone + " / " + orderDate + " / " + mealType);
        continue;
      }

      // Reserve the cache key BEFORE the wallet deduction + row write so any
      // concurrent retry that arrives during this meal's processing hits layer 1.
      try { _cache.put(_dupKey, sid, 300); } catch(e) {}

      let pStat = payStatus;
      let walletCreditUsed = 0;
      // ════ WALLET DEDUCTION LOGIC ════
      if (payMethod === "Wallet") {
        let currentBalance = _calculateWalletBalance(profile.phone, allWalletRows);

        if (currentBalance >= netTotal) {
          _appendWalletTransaction(profile.phone || "", profile.name || "Customer", "Order Deduction", netTotal, true, sid);
          // Reflect the new debit in our in-memory wallet cache so subsequent
          // meals in the same submission see the updated balance.
          allWalletRows.push({ Phone: _normalizePhone(profile.phone), Txn_Type: "Order Deduction", Amount: netTotal, Verified: "TRUE" });
          pStat = "Wallet Paid";
          walletCreditUsed = netTotal;
        } else {
          pStat = "Pending"; // Wallet failed, fallback to pending
        }
      } else if (payMethod === "Split") {
        // Split: deduct wallet portion now, UPI portion remains pending
        const requestedCredit = Math.min(Number(body.wallet_credit) || 0, netTotal);
        if (requestedCredit > 0) {
          const currentBalance = _calculateWalletBalance(profile.phone, allWalletRows);
          if (currentBalance >= requestedCredit) {
            _appendWalletTransaction(profile.phone || "", profile.name || "Customer", "Order Deduction (Wallet Part)", requestedCredit, true, sid);
            allWalletRows.push({ Phone: _normalizePhone(profile.phone), Txn_Type: "Order Deduction", Amount: requestedCredit, Verified: "TRUE" });
            walletCreditUsed = requestedCredit;
            pStat = "Pending"; // UPI portion still outstanding
          } else {
            // Not enough wallet — fall back to full UPI
            payMethod = "UPI";
            pStat = "Pending";
          }
        }
      } else if (payMethod === "On Account") {
        pStat = "On Account";
      }

      // Self-heal Wallet_Credit column if it doesn't exist yet (no initSchema needed)
      if (walletCreditUsed > 0 && !hIdx["Wallet_Credit"]) {
        const newCol = ordersWs.getLastColumn() + 1;
        ordersWs.getRange(1, newCol).setValue("Wallet_Credit");
        SpreadsheetApp.flush();
        // Refresh hIdx so set() can find it
        Object.assign(hIdx, headerIndex(ordersWs));
      }

      set("Payment_Method",      payMethod);
      set("Payment_Status",      pStat);
      if (walletCreditUsed > 0) set("Wallet_Credit", walletCreditUsed);
      // Stamp the HDFC gateway order id so the idempotency check in
      // submitOrder() can detect duplicate writes for the same charge.
      const _gOrderId = String(body.gateway_order_id || "").trim();
      if (_gOrderId) set("Gateway_Order_ID", _gOrderId);
      set("Payment_Freq",        payFreq);
      set("First_Time",          firstTime);
      set("Source",              "WebApp");

      // Fill individual item columns
      if (mealType === "Breakfast") {
        // Breakfast: dynamic items go to BF_Item_N/BF_Qty_N
        let bfSlot = 1;
        items.forEach(({colKey, qty}) => {
          if (bfSlot > 4) return;
          const displayName = (colKey === "B_CURD") ? "Curd" : resolveName(colKey);
          set(`BF_Item_${bfSlot}`, displayName);
          set(`BF_Qty_${bfSlot}`,  qty);
          bfSlot++;
        });
        // Curd goes to Curd column too
        const curdItem = items.find(x => x.colKey === "B_CURD");
        if (curdItem) set("Curd", curdItem.qty);
      } else {
        // Lunch/Dinner: map colKeys to named columns
        items.forEach(({colKey, qty}) => {
          const canonical = ITEM_COL_MAP[colKey] || colKey;
          // If canonical is still an ID, try masterMap
          const finalCol = (masterMap[canonical]) ? masterMap[canonical] : canonical;
          set(finalCol, qty);
        });
      }

      ordersWs.appendRow(row);
      _missedOrderSafetyNet(ss, sid, row, profile.phone);  // safety net — verify write succeeded
    }
  }

  // Force all buffered Sheets writes to disk before returning success
  SpreadsheetApp.flush();

  // Verify every row we just wrote actually landed; auto-recover + email if not
  _verifyAndAlertMissedOrders(ss, submissionIds);

  // Upsert customer record
  _upsertCustomer(ss, profile);

  // If user requested to settle ALL pending dues in this same transaction
  if (body.settle_all && payMethod === "Wallet") {
    _settlePendingInternal(ss, profile.phone, profile.name || "Customer");
  }

  if (payFreq === "Prepaid Wallet" || payFreq.includes("10 days") || payFreq.includes("Wallet")) {
    try { _updateLedger(ss, profile, orders); } catch(e) { /* non-fatal */ }
  }

  // Sync final promoCount back to customer sheet
  if (cRowIdx !== -1 && cIdx["Review_Promo_Count"]) {
    // Pro transition: 0 -> "Exhausted"
    let finalValue = promoCount;
    if (finalValue === 0) finalValue = "Exhausted";
    else if (finalValue === null) finalValue = "";
    
    const realRow = cRowIdx + 2;
    custWs.getRange(realRow, cIdx["Review_Promo_Count"]).setValue(finalValue);
  }

  // If 6th-day loyalty discount exceeded the bill, credit the excess to wallet (server-computed)
  if (loyaltyExcessCredit > 0) {
    try {
      _appendWalletTransaction(
        profile.phone || "", profile.name || "Customer",
        "Loyalty Streak Reward (Excess Credit)",
        loyaltyExcessCredit, true, submissionIds[0] || ""
      );
    } catch(e) { /* non-fatal */ }
  }

  // Invalidate menu cache for all ordered dates so units_remaining is fresh on next getMenu call.
  // (Cache had 60s TTL — without this, customers would see stale stock counts after placing an order.)
  if (submissionDates.length) {
    _invalidateCache(...submissionDates.map(d => "menu_v2_" + d));
  }

  return {success: true, submissionId: submissionIds[0] || "", wallet_bonus: loyaltyExcessCredit};
}
function markOnAccount(phone, cycle, status) {
  const ss = getSpreadsheet();
  const phoneStr = _normalizePhone(phone);
  const profile = {
    phone: phoneStr,
    onAccount: status,
    billingCycle: cycle
  };
  _upsertCustomer(ss, profile);
  return { success: true, phone: phoneStr, status: status, cycle: cycle };
}
// ── GET CUSTOMER ORDERS ──────────────────────────────────────
// ── GET DAY TOTALS FOR DATES (used to compute combined-day fees) ─
// Returns existing meal subtotals per date for the given phone,
// excluding the current cart being built (which is not yet placed).
function getDayTotalsForDates(phone, datesParam, preloadedRows) {
  if (!phone || !datesParam) return { dayTotals: {} };
  const dates = String(datesParam).split(',').map(d => d.trim()).filter(Boolean);
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  // Allow caller to pass pre-fetched rows (submitOrder) so we don't re-hit the sheet.
  const rows = Array.isArray(preloadedRows) ? preloadedRows : getAllRows(ws);

  const result = {};
  dates.forEach(d => { result[d] = {}; });

  // Canonicalize target phone so +91/space/decimal/scientific variants all match.
  const targetPhone = _normalizePhone(phone);

  rows.filter(r => {
    // Always group by Order_Date column (never submission timestamp) — this is what
    // keeps bills for a single meal-day intact even if the customer hits "place order"
    // on either side of IST midnight.
    if (_normalizePhone(r.Phone) !== targetPhone) return false;
    if (_isOrderCancelled(r.Payment_Status)) return false;
    const rDate = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r.Order_Date).trim();
    return dates.includes(rDate);
  }).forEach(r => {
    const rDate = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r.Order_Date).trim();
    const meal = String(r.Meal_Type).trim();
    if (!result[rDate][meal]) result[rDate][meal] = {
      subtotal:            0,
      delivery_charged:    0,
      discount_applied:    0,
      small_fee_charged:   0,
      meal_credit_applied: 0,   // v14.8 — tracks retroactive refunds already given
      count:               0
    };
    result[rDate][meal].subtotal            += Number(r.Food_Subtotal    || 0);
    result[rDate][meal].delivery_charged    += Number(r.Delivery_Charge  || 0);
    result[rDate][meal].discount_applied    += Number(r.Discount_Amount  || 0);
    result[rDate][meal].small_fee_charged   += Number(r.Small_Order_Fee  || 0);
    result[rDate][meal].meal_credit_applied += Number(r.Meal_Credit      || 0);
    result[rDate][meal].count++;
  });

  return { dayTotals: result };
}
function _calculateLoyaltyStreak(phone, preloadedRows) {
  if (!phone) return { streak: 0, pastSurcharge: 0 };
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = Array.isArray(preloadedRows) ? preloadedRows : getAllRows(ws);
  const phoneStr = _normalizePhone(phone);
  const todayISO = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

  const dailyTotals  = {}; // date → total surcharge that day
  const rewardDays   = new Set(); // dates where Loyalty_Discount = "Yes"

  rows.forEach(r => {
    if (_normalizePhone(r.Phone) !== phoneStr) return;
    // Cancelled rows (soft or hard) must NOT contribute to streak count —
    // otherwise a user could cancel days 3/4 and still hit day-6 reward.
    if (_isOrderCancelled(r.Payment_Status)) return;
    const stat = String(r.Payment_Status || "").toLowerCase();
    if (stat.includes("deleted")) return;

    const d = r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd") : String(r.Order_Date).trim();

    // For today's rows: only check if the loyalty reward was already given this morning
    // (so a separate lunch submission doesn't double-apply the day-6 discount)
    if (d === todayISO) {
      if (String(r.Loyalty_Discount || "").trim().toLowerCase() === "yes") {
        rewardDays.add(d); // today's breakfast already got the reward → block second meal
      }
      return; // Don't count today in the backward streak totals
    }
    if (d > todayISO) return; // future dates — ignore

    if (!dailyTotals[d]) dailyTotals[d] = 0;
    dailyTotals[d] += (Number(r.Inflation_Surcharge) || (Math.ceil((Number(r.Food_Subtotal)||0)/20)));

    // Track days where the 6-day loyalty reward was already given
    if (String(r.Loyalty_Discount || "").trim().toLowerCase() === "yes") {
      rewardDays.add(d);
    }
  });

  let streakCount = 0;
  let accumulatedSurcharge = 0;

  let d = new Date(); d.setDate(d.getDate() - 1); // start from yesterday
  let safety = 0;
  while (safety < 30) {
    safety++;
    if (d.getDay() === 0) { // Skip Sunday (closed)
      d.setDate(d.getDate() - 1);
      continue;
    }
    const iso = Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    if (dailyTotals[iso] !== undefined) {
      if (rewardDays.has(iso)) {
        // This day was a 6th-day reward day — it marks the END of the previous cycle.
        // Don't count it; the new cycle starts from the day after it.
        break;
      }
      streakCount++;
      accumulatedSurcharge += dailyTotals[iso];
    } else {
      break; // gap in ordering — streak broken
    }
    d.setDate(d.getDate() - 1);
  }

  // If today itself already received the loyalty reward (e.g. breakfast was submitted
  // first and marked Loyalty_Discount=Yes), treat it as a cycle already reset —
  // return streak=0 so any subsequent meal on the same day doesn't get a double reward.
  if (rewardDays.has(todayISO)) {
    return { streak: 0, pastSurcharge: 0 };
  }

  return { streak: streakCount, pastSurcharge: accumulatedSurcharge };
}
// ── VERIFY ORDER PLACED (timeout recovery) ───────────────────
// Called by the frontend after a network timeout to check if the order
// actually landed on the backend. Matches by phone + every date/meal combo
// in the cart, within a 10-minute recency window.
// Returns { found: true, submissionId } or { found: false }.
function verifyOrderPlaced(body) {
  const phone = _normalizePhone(String(body.phone || ""));
  if (!phone) return { found: false };

  // cart = [{date: "yyyy-MM-dd", meal: "Breakfast"|"Lunch"|"Dinner"}]
  // Derived from body.orders (same format as submitOrder)
  const orders = body.orders || [];
  const cartEntries = []; // [{date, meal}]
  for (const dateOrder of orders) {
    for (const meal of (dateOrder.meals || [])) {
      if ((Number(meal.subtotal) || 0) > 0) {
        cartEntries.push({ date: dateOrder.date, meal: meal.type });
      }
    }
  }
  if (!cartEntries.length) return { found: false };

  const ss  = getSpreadsheet();
  const ws  = getOrCreateTab(ss, TAB_ORDERS, []);
  const rows = getAllRows(ws);

  const nowMs     = Date.now();
  const TEN_MIN   = 10 * 60 * 1000;
  const normDate  = (d) => {
    if (!d) return "";
    if (d instanceof Date) return Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    return String(d).trim().substring(0, 10);
  };

  // Recent rows for this phone (last 10 min)
  const recent = rows.filter(r => {
    if (_normalizePhone(String(r.Phone || "")) !== phone) return false;
    const rMs = r.Submitted_At ? new Date(r.Submitted_At).getTime() : 0;
    return rMs > 0 && (nowMs - rMs) <= TEN_MIN;
  });

  if (!recent.length) return { found: false };

  // Every cart entry must have a matching recent row
  let firstId = null;
  for (const entry of cartEntries) {
    const match = recent.find(r =>
      normDate(r.Order_Date) === entry.date && r.Meal_Type === entry.meal
    );
    if (!match) return { found: false };
    if (!firstId) firstId = String(match.Submission_ID || "");
  }

  return firstId ? { found: true, submissionId: firstId } : { found: false };
}
// ── DELETE ORDER (with Refund Logic) ─────────────────────────
function deleteOrder(phone, rowId, refundType, opts) {
  // ─── CONCURRENCY GUARD ─────────────────────────────────────────────
  // Prevents parallel deletes (double-clicks, retries) from both finding
  // the same row, both appending refunds, and both calling deleteRow on
  // shifted indices. Without this, the second call deleted the wrong row.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, error: "System busy. Please try again in a moment." };
  }
  try {
    return _deleteOrderInternal(phone, rowId, refundType, opts);
  } catch (e) {
    // Top-level safety net so transient Drive/Sheets errors don't surface
    // as raw "Service error: Drive" to the user. Logged for diagnosis.
    console.error(`deleteOrder failed for rowId=${rowId} phone=${phone}: ${e && e.message}\n${e && e.stack}`);
    return {
      success: false,
      error: "Could not cancel right now (a Google service blip). Please try again in a few seconds."
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
function _deleteOrderInternal(phone, rowId, refundType, opts) {
  opts = opts || {};
  const isAdminCall = !!opts.isAdmin;
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const now = getISTDate();
  let msg = "Order deleted successfully";
  const today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
  const hourIST = now.getHours() + now.getMinutes() / 60;
  const CUTOFFS = { Breakfast: 7, Lunch: 9, Dinner: 16.5 };

  // ── Ownership guard ──────────────────────────────────────────
  // Customers: must pass BOTH the exact Submission_ID and the matching phone.
  // Admin: can delete by Submission_ID alone (phone not required).
  // Submission_ID is compared as a full exact string (case-insensitive) to
  // prevent the old "digits-only" collision bug where SK-20250101-XYZ and
  // SK-20250101-ABC both reduced to "20250101".
  const targetId = String(rowId || "").trim().toUpperCase();
  if (!targetId) {
    return { success: false, error: "Missing order identifier." };
  }
  const normTargetPhone = _normalizePhone(phone);

  const r = rows.find(x => {
    const sheetId = String(x.Submission_ID || "").trim().toUpperCase();
    if (sheetId !== targetId) return false;
    if (isAdminCall) return true; // Admin bypass — PIN already verified
    // Customer must also match phone
    return _normalizePhone(x.Phone) === normTargetPhone;
  });
  if (!r) {
    console.error(`CANCELLATION FAILED: Submission ID "${rowId}" not found or phone mismatch for ${phone} (admin=${isAdminCall}).`);
    return {success: false, error: "Order record not found or you do not have permission to cancel it."};
  }
  // ── ALREADY CANCELLED GUARD ────────────────────────────────────────
  // Row is kept forever now. Prevent double-cancellation attempts.
  if (_isOrderCancelled(r.Payment_Status)) {
    return { success: false, error: "This order has already been cancelled." };
  }
  const orderDateStr = r.Order_Date instanceof Date
    ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
    : String(r.Order_Date).trim();
  if (orderDateStr < today) return {success: false, error: "Cannot delete past orders"};

  // Block deletion if cutoff has passed for today's orders
  // Normalize meal type — strip whitespace + title-case — to avoid silent skip when value is " breakfast" or "BREAKFAST"
  const mealNorm = String(r.Meal_Type || "").trim().toLowerCase();
  const mealKey  = mealNorm.charAt(0).toUpperCase() + mealNorm.slice(1);
  if (orderDateStr === today) {
    const cutoffHour = CUTOFFS[mealKey];
    if (cutoffHour !== undefined && hourIST >= cutoffHour) {
      return {success: false, error: `Cutoff for ${mealKey} has already passed`};
    }
  }

  // ─── IDEMPOTENCY / RECOVERY GUARD ─────────────────────────────────
  // If a pending refund already exists for this Submission_ID, DON'T add
  // another, but DO finish the cancellation by deleting the order row.
  // (Previous attempt may have written the refund then failed before delete.)
  let existingPendingRefund = null;
  try {
    const refundsWs = ss.getSheetByName(TAB_REFUNDS);
    if (refundsWs && refundsWs.getLastRow() > 1) {
      const refRows = getAllRows(refundsWs);
      existingPendingRefund = refRows.find(rf => {
        const rfId = String(rf.Submission_ID || "").trim().toUpperCase();
        const rfStat = String(rf.Status || "").trim().toLowerCase();
        return rfId === targetId && rfStat === "pending";
      }) || null;
    }
  } catch (e) { /* non-fatal */ }

  if (existingPendingRefund) {
    // Recovery path: refund row already exists, just ensure order row is marked cancelled.
    try {
      const hIdxR = headerIndex(ws);
      const statusColR = hIdxR["Payment_Status"] || hIdxR["Payment Status"];
      if (statusColR && !_isOrderCancelled(r.Payment_Status)) {
        ws.getRange(r._row, statusColR).setValue("Cancelled \u2013 UPI Refund Pending");
      }
    } catch (e) { /* non-fatal */ }
    return {
      success: true,
      message: "Cancellation completed. Your refund was already in the queue and will be processed within 1-2 days."
    };
  }

  // GRACEFUL REFUND HANDLING with eligibility recalculation (Cases 1/2/3)
  const pStatStr = String(r.Payment_Status).toLowerCase();
  const isOnAccountOrder = pStatStr === "on account";
  let finalType = refundType; // Declare here so it is accessible at the end of the function for the soft-cancel remark.

  if (pStatStr === "paid" || pStatStr === "wallet paid" || isOnAccountOrder) {
    const custName = r.Customer_Name || "Customer";
    const ordersWs2 = ws; // same sheet
    const hIdx = headerIndex(ws); // needed for updating remaining rows
    const deleteDate = orderDateStr;
    const deleteMeal = String(r.Meal_Type).trim();

    // Get all orders for this phone+date (excluding the one being deleted)
    const sameDayRows = rows.filter(x =>
      String(x.Phone).trim() === String(phone).trim() &&
      (() => {
        const xd = x.Order_Date instanceof Date
          ? Utilities.formatDate(x.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
          : String(x.Order_Date).trim();
        return xd === deleteDate && String(x.Submission_ID) !== String(rowId);
      })()
    );

    // Calc remaining day subtotal (food only) after deletion
    const remainingDaySubtotal = sameDayRows.reduce((s, x) => s + (Number(x.Food_Subtotal) || 0), 0);
    // Calc old day subtotal (including deleted row)
    const oldDaySubtotal = remainingDaySubtotal + (Number(r.Food_Subtotal) || 0);

    // Over-discount claw-back: only claw back discounts that were ACTUALLY applied
    // to remaining orders (read from their Discount_Amount column), not a theoretical
    // volume tier. submitOrder only applies loyalty (6th-day) discounts, not volume tiers.
    let overDiscount = 0;
    {
      const discColIdx = hIdx["Discount_Amount"];
      const netColIdx  = hIdx["Net_Total"];

      // Sum of discounts actually applied to remaining rows
      const totalActualDiscount = sameDayRows.reduce((s, x) => s + (Number(x.Discount_Amount) || 0), 0);

      if (totalActualDiscount > 0) {
        // Re-compute what discount the remaining orders SHOULD get after deletion.
        // We use their total food subtotal and compare to what was actually given.
        // For now: if the deleted order was the "trigger" for the day's loyalty discount,
        // the remaining orders should have 0 discount (they didn't earn it alone).
        // Claw back = (actual given) − (what they deserve now).
        // Conservative: only claw back if none of the remaining rows have Loyalty_Discount=Yes.
        const remainingHasLoyalty = sameDayRows.some(x =>
          String(x.Loyalty_Discount || "").trim().toLowerCase() === "yes"
        );

        if (!remainingHasLoyalty) {
          // No loyalty day in remaining rows — the deleted row was the discount trigger.
          // Claw back all discounts from remaining rows.
          overDiscount = totalActualDiscount;

          // Update remaining rows: zero out their Discount_Amount and restore Net_Total
          if (overDiscount > 0 && discColIdx && netColIdx) {
            sameDayRows.forEach(x => {
              const xSub      = Number(x.Food_Subtotal)       || 0;
              const xSurcharge= Number(x.Inflation_Surcharge) || 0;
              const xDelivery = Number(x.Delivery_Charge)     || 0;
              const xSmallFee = Number(x.Small_Order_Fee)     || 0;
              const xReviewD  = Number(x.Review_Discount)     || 0;
              const newNetTotal = xSub + xDelivery + xSmallFee + xSurcharge - xReviewD; // discount = 0
              ws.getRange(x._row, discColIdx).setValue(0);
              ws.getRange(x._row, netColIdx) .setValue(newNetTotal);
            });
          }
        }
      }
    }

    // Delivery & Fee eligibility for remaining same-day orders
    // If combined day total drops below 150 after deletion → those remaining orders
    // should have been charged fees → customer saved them unjustly.
    const FREE_THR_D = 150;
    const freeAreaNames2 = getAreas().filter(a => a.free).map(a => a.name);
    const isNonFree = (area) => !freeAreaNames2.includes(area) && area !== "Self Pickup";

    let deliveryOwed = 0;
    let smallFeeOwed = 0;

    if (oldDaySubtotal >= FREE_THR_D && remainingDaySubtotal < FREE_THR_D) {
      // Day total drops below free-delivery threshold → remaining orders now owe fees.
      // We claw the amounts from THIS refund, AND update those rows in the sheet so that
      // if they are later cancelled themselves, the clawback doesn't fire a second time.
      const delivColIdx   = hIdx["Delivery_Charge"];
      const smallFeeColIdx = hIdx["Small_Order_Fee"];
      const netColIdx2    = hIdx["Net_Total"];

      sameDayRows.forEach(x => {
        const xArea = x.Area || "";
        const xSub  = Number(x.Food_Subtotal) || 0;
        let netDelta = 0;

        // 1. Delivery Clawback: order was in non-free area but charged ₹0 due to threshold
        if (xSub > 0 && isNonFree(xArea) && (Number(x.Delivery_Charge) || 0) === 0) {
          deliveryOwed += 10;
          netDelta += 10;
          if (delivColIdx) ws.getRange(x._row, delivColIdx).setValue(10);
        }

        // 2. Small Order Fee Clawback: Lunch/Dinner sub < ₹50 was waived due to threshold
        const xMeal = String(x.Meal_Type).trim();
        if ((xMeal === "Lunch" || xMeal === "Dinner") && xSub > 0 && xSub < 50
            && (Number(x.Small_Order_Fee) || 0) === 0) {
          smallFeeOwed += 10;
          netDelta += 10;
          if (smallFeeColIdx) ws.getRange(x._row, smallFeeColIdx).setValue(10);
        }

        // Update Net_Total on remaining row to reflect newly owed fees (prevents double-clawback)
        if (netDelta > 0 && netColIdx2) {
          ws.getRange(x._row, netColIdx2).setValue((Number(x.Net_Total) || 0) + netDelta);
        }
      });
    }

    // Loyalty Clawback Logic
    // If deleting an order breaks a streak that received a reward on a later date.
    // Admin cancellations are EXEMPT — streak is not penalised when kitchen cancels.
    let loyaltyClawback = 0;
    let loyaltyClawbackNote = "";
    const phoneStr = _normalizePhone(phone);

    if (!isAdminCall) {
      // Scan for a later streak-reward order that this cancellation would invalidate.
      const laterPayoffs = rows.filter(x => {
        if (String(x.Submission_ID) === String(rowId)) return false;
        if (_normalizePhone(x.Phone) !== phoneStr) return false;
        const xStat = String(x.Payment_Status || "").toLowerCase();
        if (xStat.includes("cancelled") || xStat.includes("deleted")) return false;
        if (String(x.Loyalty_Discount || "").trim().toLowerCase() !== "yes") return false;
        const xDate = x.Order_Date instanceof Date
          ? Utilities.formatDate(x.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
          : String(x.Order_Date).trim();
        return xDate >= orderDateStr; // payoff on or after the cancelled order's date
      });

      if (laterPayoffs.length > 0) {
        loyaltyClawback = Number(laterPayoffs[0].Discount_Amount) || 0;
        loyaltyClawbackNote = `Loyalty reward of ₹${loyaltyClawback} was applied on ${
          (() => { const d = laterPayoffs[0].Order_Date; return d instanceof Date ? Utilities.formatDate(d,"Asia/Kolkata","dd MMM") : String(d); })()
        } — cancelling this order breaks your streak, so that reward is reversed.`;
      }
    }

    // Refund = Net_Total − adjustment
    // Net_Total already correctly encodes: food + delivery + fees + surcharge − discount − mealCredit − reviewDiscount
    const adjustment = overDiscount + deliveryOwed + smallFeeOwed + loyaltyClawback;
    const rawRefund = Number(r.Net_Total) || 0;
    const netRefund = rawRefund - adjustment;           // may be negative
    const refundAmt = Math.max(0, netRefund);           // amount actually returned
    const cancellationCharge = Math.max(0, -netRefund); // deficit charged to wallet if order < clawback

    // ── HUMAN-READABLE REFUND BREAKDOWN ────────────────────────────────────
    function buildRefundBreakdown() {
      const lines = [];
      if (adjustment === 0) {
        lines.push(`Full refund of ₹${refundAmt}.`);
        return lines.join("\n");
      }
      lines.push(`Order total: ₹${rawRefund}`);
      lines.push(`Deductions (₹${adjustment} total):`);
      if (overDiscount > 0) {
        lines.push(`  • -₹${overDiscount} — discount reversal: a loyalty discount applied to your other order(s) on this day is reversed since it was earned as part of this streak order.`);
      }
      if (deliveryOwed > 0) {
        const numOrders = deliveryOwed / 10;
        lines.push(`  • -₹${deliveryOwed} — delivery fee: your remaining ${numOrders > 1 ? numOrders + " orders" : "order"} had free delivery because day total was ₹150+. It now drops below ₹150, so ₹10 delivery applies.`);
      }
      if (smallFeeOwed > 0) {
        lines.push(`  • -₹${smallFeeOwed} — small cart fee: a remaining order under ₹50 had its ₹10 small cart fee waived (day total was ₹150+). Now that drops below ₹150, the fee applies.`);
      }
      if (loyaltyClawback > 0) {
        lines.push(`  • -₹${loyaltyClawback} — loyalty reward reversal: ${loyaltyClawbackNote}`);
      }
      if (cancellationCharge > 0) {
        lines.push(`Refund: ₹${rawRefund} − ₹${adjustment} = -₹${cancellationCharge}`);
        lines.push(`Since the deduction (₹${adjustment}) exceeds your order amount (₹${rawRefund}), ₹${cancellationCharge} has been charged to your Svaadh Wallet. This will be deducted from your next order.`);
      } else {
        lines.push(`Refund: ₹${rawRefund} − ₹${adjustment} = ₹${refundAmt}`);
      }
      return lines.join("\n");
    }

    // ── DRY RUN: return breakdown without making any changes ─────────────────
    if (opts.dryRun) {
      return {
        success:            true,
        dryRun:             true,
        refundAmt:          refundAmt,
        adjustment:         adjustment,
        cancellationCharge: cancellationCharge,
        breakdownText:      buildRefundBreakdown()
      };
    }

    // Multi-Payment Logic: If any OTHER order for this meal/date is Wallet Paid,
    // force this refund to Wallet too (to keep the day's bookkeeping simple).
    const hasAnyOtherWalletPaid = sameDayRows.some(x => {
      const typeMatch = String(x.Meal_Type).trim() === deleteMeal;
      const statusMatch = String(x.Payment_Status).toLowerCase() === "wallet paid";
      return typeMatch && statusMatch;
    });

    finalType = refundType;
    let msgSuffix = "";

    // Auto-detect wallet refund if current was wallet paid, overriding passed type
    const currentWasWallet = (pStatStr === "wallet paid");
    const currentWasSplit  = (String(r.Payment_Method || "").trim().toLowerCase() === "split");
    if (isOnAccountOrder) {
      // On Account: no cash was collected — mark row as cancelled.
      // Remaining rows already updated above (discount/delivery recalculation).
      // On-account balance auto-corrects since it's derived from live sheet rows.
      msg = "Order removed from your On Account balance.";
      finalType = "__on_account_handled__"; // skip all refund payout logic
    } else if (currentWasWallet) {
      finalType = "wallet";
    } else if (currentWasSplit) {
      // Split orders: entire refund always goes to Wallet — wallet + UPI portions both back to wallet.
      if (refundAmt > 0) {
        _appendWalletTransaction(phone, custName, "Order Cancellation Refund", refundAmt, true, String(rowId));
      }
      msg = buildRefundBreakdown() + `\n\n₹${refundAmt} refunded to your Wallet.`;
      finalType = "__split_handled__"; // skip normal logic below
    } else if (hasAnyOtherWalletPaid && refundType === "manual_upi") {
      finalType = "wallet";
      msgSuffix = "\n(Consolidated to Wallet since other items in this meal were Wallet Paid.)";
    }

    // If cancellation charge > 0 (loyalty clawback exceeded the order value):
    // debit the deficit from the wallet — it'll show as a negative balance
    // that gets collected on the customer's next order.
    if (cancellationCharge > 0) {
      _appendWalletTransaction(phone, custName,
        `Cancellation Charge (streak reward reversal — ₹${loyaltyClawback} reward was applied to an order because of this order. Cancelling it breaks your streak, so ₹${cancellationCharge} is recovered here.)`,
        -cancellationCharge, true, String(rowId));
    }

    if (finalType === "wallet") {
      if (refundAmt > 0) {
        _appendWalletTransaction(phone, custName, "Order Cancellation Refund", refundAmt, true, String(rowId));
      }
      const walletLine = cancellationCharge > 0
        ? `₹0 refunded — ₹${cancellationCharge} charged to your Wallet (will be collected on your next order).`
        : `₹${refundAmt} refunded to your Wallet.${msgSuffix}`;
      msg = buildRefundBreakdown() + `\n\n` + walletLine;
    }
    else if (finalType === "manual_upi") {
      const REF_HEADERS = ["Submission_ID","Phone","Name","Amount","Meal","Date","Status","Timestamp","Adjustment_Note","Refund_Mode"];
      const refWs = getOrCreateTab(ss, TAB_REFUNDS, REF_HEADERS);
      const note = adjustment > 0
        ? `Adjusted -₹${adjustment} (overDiscount:${overDiscount}, deliveryOwed:${deliveryOwed}, smallFeeOwed:${smallFeeOwed}, loyaltyClawback:${loyaltyClawback})`
        : "";
      refWs.appendRow([rowId, phone, custName, refundAmt, r.Meal_Type, orderDateStr, "Pending", now, note, "upi"]);
      const upiLine = cancellationCharge > 0
        ? `₹0 refunded via UPI — ₹${cancellationCharge} charged to your Wallet (will be collected on your next order).`
        : `₹${refundAmt} refund request raised — we'll process it within 1-2 days.`;
      msg = buildRefundBreakdown() + `\n\n` + upiLine;
    }
  } 
  // ── SOFT CANCELLATION FOR UPI / SPLIT ──────────────────────────────────────
  // "Pending" means customer has ALREADY paid (UPI screenshot sent) but admin hasn't verified yet.
  // For Split orders, "Pending" = wallet was deducted AND UPI payment was sent — must soft-cancel just like UPI.
  // Admin will verify and then "Verify & Refund" triggers the split refund logic in markOrdersStatus.
  if (String(r.Payment_Status || "").toLowerCase().includes("pending") && (refundType === "wallet" || refundType === "manual_upi")) {
    let hIdx = headerIndex(ws);
    
    // Robust header detection (support both underscores and spaces)
    const statusCol = hIdx["Payment_Status"] || hIdx["Payment Status"];
    
    if (!hIdx["Refund_Preference"]) {
      const col = ws.getLastColumn() + 1;
      ws.getRange(1, col).setValue("Refund_Preference")
        .setFontWeight("bold").setBackground("#c0392b").setFontColor("white");
      hIdx = headerIndex(ws);
    }
    const prefCol = hIdx["Refund_Preference"];
    
    if (statusCol && prefCol) {
      ws.getRange(r._row, statusCol).setValue("Cancelled (Verify UPI)");
      // Split orders: refund preference is always wallet (full amount back to wallet)
      const isSoftSplit = String(r.Payment_Method || "").trim().toLowerCase() === "split";
      ws.getRange(r._row, prefCol).setValue(isSoftSplit ? "wallet" : refundType);
      console.info(`SUCCESS: Soft-cancelled row ${r._row} with preference ${isSoftSplit ? "wallet (split)" : refundType}`);

      // For split orders: wallet portion is already deducted — refund it immediately.
      // UPI portion will be added to wallet once admin verifies.
      let softCancelMsg = "Cancellation request received! Admin will verify your payment and process the refund (1-2 days). ✅";
      if (isSoftSplit) {
        const walletCredit = Number(r.Wallet_Credit) || 0;
        const upiDue = Math.max(0, (Number(r.Net_Total) || 0) - walletCredit);
        if (walletCredit > 0) {
          _appendWalletTransaction(phone, r.Customer_Name || "Customer", "Order Cancellation Refund (Wallet Part)", walletCredit, true, String(rowId));
        }
        softCancelMsg = upiDue > 0
          ? `₹${walletCredit} has been refunded to your Wallet instantly. ` +
            `Once Admin verifies your ₹${upiDue} UPI payment, it will also be added to your Wallet (1-2 days). ✅`
          : `₹${walletCredit} has been refunded to your Wallet. ✅`;
      }
      return { success: true, message: softCancelMsg };
    } else {
      console.error(`FAILED: Missing columns for soft-cancel. StatusCol:${statusCol}, PrefCol:${prefCol}`);
    }
  }

  // ─── SOFT-CANCEL THE ROW (mark status, never delete) ─────────────────
  // Orders are kept forever for audit trail. The Payment_Status remark
  // ensures the row is excluded from all prep/delivery counts via _isOrderCancelled().
  {
    const hIdxFinal = headerIndex(ws);
    const statusColFinal = hIdxFinal["Payment_Status"] || hIdxFinal["Payment Status"];
    if (statusColFinal) {
      let cancelRemark;
      if (finalType === "wallet" || finalType === "__split_handled__") {
        cancelRemark = "Cancelled \u2013 Refunded to Wallet";
      } else if (finalType === "manual_upi") {
        cancelRemark = "Cancelled \u2013 UPI Refund Pending";
      } else if (finalType === "__on_account_handled__") {
        cancelRemark = "Cancelled \u2013 On Account";
      } else {
        // Fallback for unknown type (e.g. zero-refund edge cases)
        cancelRemark = "Cancelled";
      }
      ws.getRange(r._row, statusColFinal).setValue(cancelRemark);
      console.info(`ORDER SOFT-CANCELLED: Row ${r._row} (${rowId}) marked as '${cancelRemark}'`);
    } else {
      console.error(`SOFT-CANCEL FAILED: Payment_Status column not found in header index.`);
    }
  }

  return {success: true, message: msg};
}
// Shared coord extractor for Apps Script (mirrors client-side regex)
// Priority: !3d/!4d (actual pinned location) > place/@ (share URL center) >
// ?q= / ?destination= / ?ll= > @ (camera center — last resort, can be far off)


// ── ORDER SUMMARY ────────────────────────────────────────────
function getOrderSummary(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date && !_isOrderCancelled(r.Payment_Status);
  });

  var meals = {};
  var totals = {orders: 0, customers: 0, revenue: 0, paid: 0, pending: 0};
  var customerSet = {};
  var LUNCH_DINNER_COLS = [
    "Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
    "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full",
    "Dal","Rice","Salad","Curd"
  ];

  dayRows.forEach(function(r) {
    var meal = String(r.Meal_Type || "");
    if (!meal) return;
    if (!meals[meal]) meals[meal] = {count:0, revenue:0, paid:0, pending:0, itemTotals:{}, customers:[]};
    var m = meals[meal];
    var net = Number(r.Net_Total) || 0;
    var payStatus = String(r.Payment_Status || "Pending");

    var items = {};
    if (meal === "Breakfast") {
      for (var n = 1; n <= 4; n++) {
        var item = String(r["BF_Item_"+n] || "").trim();
        var qty  = Number(r["BF_Qty_"+n]) || 0;
        if (item && qty > 0) {
          items[item] = (items[item] || 0) + qty;
          m.itemTotals[item] = (m.itemTotals[item] || 0) + qty;
        }
      }
      var curdBf = Number(r.Curd) || 0;
      if (curdBf > 0) { items["Curd"] = (items["Curd"] || 0) + curdBf; m.itemTotals["Curd"] = (m.itemTotals["Curd"] || 0) + curdBf; }
    } else {
      LUNCH_DINNER_COLS.forEach(function(col) {
        var q = Number(r[col]) || 0;
        if (q > 0) { items[col] = (items[col]||0)+q; m.itemTotals[col] = (m.itemTotals[col]||0)+q; }
      });
    }

    m.count++;
    m.revenue += net;
    if (payStatus === "Paid" || payStatus === "Wallet Paid" || payStatus === "Collected") m.paid += net; else m.pending += net;
    m.customers.push({
      id:        String(r.Submission_ID || ""),
      name:      String(r.Customer_Name || ""),
      phone:     String(r.Phone || ""),
      items:     items,
      area:      String(r.Area || ""),
      address:   String(r.Full_Address || r.Flat || ""),
      total:     net,
      payStatus: payStatus,
      notes:     String(r.Special_Notes_Kitchen || "")
    });

    totals.orders++;
    totals.revenue += net;
    if (payStatus === "Paid" || payStatus === "Wallet Paid" || payStatus === "Collected") totals.paid += net; else totals.pending += net;
    if (!customerSet[String(r.Phone)]) { customerSet[String(r.Phone)] = true; totals.customers++; }
  });

  return {date: date, meals: meals, totals: totals};
}
// ── GET ORDER HISTORY (date range) ────────────────────────────────────────────
function getOrderHistory(p) {
  var dateFrom = p.dateFrom, dateTo = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  // Archive-aware read so the History tab can show data from archived
  // quarters too. Falls back to live-only on any archive-read error.
  var rows;
  try {
    rows = getOrdersInRangeWithArchive(dateFrom, dateTo) || [];
  } catch (e) {
    console.warn("getOrderHistory: archive lookup failed, falling back to live: " + e.message);
    rows = getAllRows(ws);
  }

  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var filtered = rows.filter(function(r) {
    var d = fmtDate(r.Order_Date);
    return d >= dateFrom && d <= dateTo && !_isOrderCancelled(r.Payment_Status);
  });

  var orders = filtered.map(function(r) {
    var items = {};
    try { if (r.Items_JSON) items = JSON.parse(r.Items_JSON); } catch(e) {}
    return {
      id:             r.Submission_ID,
      date:           fmtDate(r.Order_Date),
      meal:           r.Meal_Type,
      name:           r.Customer_Name,
      phone:          r.Phone,
      area:           r.Area || "",
      wing:           r.Wing || "",
      flat:           r.Flat || "",
      total:          Number(r.Net_Total) || 0,
      gross:          Number(r.Gross_Total) || 0,
      status:         r.Payment_Status || "Pending",
      payment_method: r.Payment_Method || "UPI",
      notes:          r.Special_Notes || "",
      items:          items,
      delivery:       Number(r.Delivery_Charge) || 0,
      discount:       Number(r.Loyalty_Discount) || 0
    };
  });

  var totalRev     = orders.reduce(function(s,o){return s+o.total;},0);
  var totalPaid    = orders.filter(function(o){
    return String(o.status)==="Paid" || String(o.status)==="Wallet Paid" || String(o.status)==="Collected";
  }).reduce(function(s,o){return s+o.total;},0);
  var uniqueCusts  = Object.keys(orders.reduce(function(m,o){m[o.phone]=1;return m;},{})).length;

  return {
    success: true,
    orders:  orders,
    summary: {
      orderCount:      orders.length,
      uniqueCustomers: uniqueCusts,
      totalRevenue:    Math.round(totalRev),
      totalPaid:       Math.round(totalPaid),
      totalPending:    Math.round(totalRev - totalPaid)
    }
  };
}
// Audit Fix #13: Helper to cancel order from Admin Dashboard
function adminCancelOrder(body) {
  const pin = String(body.pin || "").trim();
  if (pin !== ADMIN_PIN) return {success:false, error:"STRICT ADMIN PIN REQUIRED"};
  
  const phone = String(body.phone || "").trim();
  const dateStr = String(body.date || "").trim();
  const meal = String(body.meal || "").trim();

  const ss = getSpreadsheet();
  const ws = ss.getSheetByName(TAB_ORDERS);
  const rows = getAllRows(ws);

  // Find ALL matching orders for this guest/meal/date
  const matches = rows.filter(r => {
    const rPhone = String(r.Phone || "").trim();
    const rMeal = String(r.Meal_Type || "").trim();
    const orderDate = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r.Order_Date).trim();
    const status = String(r.Payment_Status || "").toLowerCase();
    
    return rPhone === phone && rMeal === meal && orderDate === dateStr && status !== 'deleted' && !status.startsWith('cancelled');
  });

  if (!matches.length) return {success:false, error: "No matching orders found"};

  // 1. Determine Global Batch Refund Type
  // If ANY order in the batch is wallet-paid, or any OTHER order in the sheet for this guest/meal is wallet-paid
  let anyWallet = matches.some(m => String(m.Payment_Status).toLowerCase() === "wallet paid");
  if (!anyWallet) {
    // Also check if any order REMAINS that is wallet paid (should have been covered by filter but check rows list)
    anyWallet = rows.some(r => 
      String(r.Phone).trim() === phone && 
      String(r.Meal_Type).trim() === meal &&
      Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd') === dateStr &&
      String(r.Payment_Status).toLowerCase() === "wallet paid"
    );
  }

  // 2. Process each match one by one
  let totalRefund = 0;
  let batchMsg = "";
  let refundedToWallet = false;

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const pStat = String(m.Payment_Status).toLowerCase();
    
    // Decide refundType for this specific row in the batch context
    let rType = "none";
    if (pStat === "wallet paid") rType = "wallet";
    else if (pStat === "paid" || pStat.includes("pending")) {
      rType = anyWallet ? "wallet" : "manual_upi";
    }

    const result = deleteOrder(phone, m.Submission_ID, rType);
    if (result.success) {
      if (typeof result.message === "string" && result.message.includes("Wallet")) refundedToWallet = true;
    }
    // Force spreadsheet synchronization to avoid row-index/cache mismatch in the next iteration
    SpreadsheetApp.flush();
  }

  const noun = matches.length === 1 ? "order" : "orders";
  const mode = refundedToWallet ? "Wallet" : "Approvals";
  return {
    success: true, 
    message: `${matches.length} ${meal} ${noun} cancelled successfully (Refunded to ${mode})`
  };
}
function placeBulkOrders(body) {
  const pin = String(body.pin || "").trim();
  if (pin !== ADMIN_PIN) return {success:false, error:"STRICT ADMIN PIN REQUIRED"};
  
  const phone = body.phone;
  const name = body.name;
  const templates = body.templates; 
  const dates = body.dates;     
  
  let count = 0;
  dates.forEach(function(date) {
    templates.forEach(function(tpl) {
      const orderBody = {
        phone: phone,
        name: name,
        date: date,
        meal: tpl.meal,
        items: tpl.items,
        payment_method: "Wallet",
        payment_freq: "Prepaid Wallet",
        source: "Admin Bulk"
      };
      const res = submitOrder(orderBody);
      if (res.success) count++;
    });
  });
  return {success:true, count: count};
}
// ── SUBMIT MANUAL ORDER (Admin Feature) ────────────────────
function submitManualOrder(body) {
  const ss        = getSpreadsheet();
  const ordersWs  = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const custWs    = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);

  const phone     = String(body.phone    || "").trim();
  const name      = String(body.name     || "").trim();
  const amount    = Number(body.amount)  || 0;
  const date      = body.date || Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const mealType  = body.mealType  || "Other";
  const payMethod = body.paymentMethod || body.payMethod || "UPI";  // "UPI" | "On Account" | "Cash"
  const billingCycle = body.billingCycle || "Daily"; // only used when creating/updating for On Account

  if (!phone || amount <= 0) throw new Error("Invalid phone or amount");

  // ── 1. Look up existing customer ───────────────────────────
  const custRows = getAllRows(custWs);
  const pStr     = _normalizePhone(phone);
  const existing = custRows.find(r => _normalizePhone(r.Phone) === pStr);

  // ── 2. Update / create customer record ─────────────────────
  if (existing) {
    // Always safe to update name if provided, nothing else unless On Account
    const custHIdx = headerIndex(custWs);
    const updCell  = (col, val) => { if (custHIdx[col]) custWs.getRange(existing._row, custHIdx[col]).setValue(val); };

    if (name) updCell("Customer_Name", name);

    if (payMethod === "On Account") {
      // Mark On Account = Yes
      updCell("On_Account", "Yes");
      // Only change billing cycle if the existing one is NOT Monthly (never downgrade)
      const existingCycle = String(existing.Billing_Cycle || "").trim();
      if (existingCycle !== "Monthly") {
        updCell("Billing_Cycle", billingCycle);
      }
    }
    // For UPI/Cash: do NOT touch On_Account, Billing_Cycle, or any other field
    SpreadsheetApp.flush();

  } else {
    // New customer — create a minimal record; leave address/area/etc blank
    // so they can self-register later and fill in details naturally.
    const newCustProfile = {
      phone:    phone,
      name:     name || "",
      // Address fields intentionally omitted (blank)
    };
    if (payMethod === "On Account") {
      newCustProfile.onAccount    = "Yes";
      newCustProfile.billingCycle = billingCycle;
    }
    // For UPI/Cash: On_Account defaults to "No", Billing_Cycle to "Daily" (schema defaults)
    _upsertCustomer(ss, newCustProfile);
  }

  // ── 3. Pull customer address for order row (if they exist) ─
  const custRecord = existing || (() => {
    // Re-fetch after insert so we get the row
    const freshRows = getAllRows(custWs);
    return freshRows.find(r => _normalizePhone(r.Phone) === pStr);
  })();
  const custAddress = custRecord
    ? [custRecord.Wing && `Wing ${custRecord.Wing}`, custRecord.Flat && `Flat ${custRecord.Flat}`,
       custRecord.Floor && `${custRecord.Floor} Floor`, custRecord.Society, custRecord.Area]
       .filter(Boolean).join(", ")
    : "";
  const custArea    = custRecord ? (custRecord.Area    || "") : "";
  const custSociety = custRecord ? (custRecord.Society || "") : "";
  const custMaps    = custRecord ? (custRecord.Maps_Link || "") : "";

  // ── 4. Determine Payment_Method + Payment_Status for order row
  let orderPayMethod, orderPayStatus;
  if (payMethod === "On Account") {
    orderPayMethod = "On Account";
    orderPayStatus = "On Account";
  } else if (payMethod === "Cash") {
    orderPayMethod = "Cash";
    orderPayStatus = "Paid";
  } else {
    // UPI (default)
    orderPayMethod = "UPI";
    orderPayStatus = "Pending";
  }

  // ── 5. Append order row ────────────────────────────────────
  const hIdx = headerIndex(ordersWs);
  const row  = new Array(ORDERS_HEADERS.length).fill("");
  const set  = (colName, val) => { const i = hIdx[colName]; if (i) row[i - 1] = val; };

  // Generate a unique ID: SK-YYYYMMDD-M-XXXX (M = manual, easy to identify)
  const _midDate = Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyyMMdd");
  const _midRand = Math.floor(Math.random() * 9000) + 1000;
  const sid = `SK-${_midDate}-M-${_midRand}`;

  set("Submission_ID",  sid);
  set("Submitted_At",   getISTTimestamp());
  set("Order_Date",     date);
  set("Meal_Type",      mealType);
  set("Customer_Name",  name || (custRecord && custRecord.Customer_Name) || "");
  set("Phone",          _normalizePhone(phone));
  set("Food_Subtotal",  amount);
  set("Net_Total",      amount);
  set("Payment_Method", orderPayMethod);
  set("Payment_Status", orderPayStatus);
  set("Address",        custAddress);
  set("Area",           custArea);
  set("Society",        custSociety);
  set("Maps_Link",      custMaps);
  set("Source",         "Admin Manual Entry");

  ordersWs.appendRow(row);
  return { success: true, sid: sid };
}
// ============================================================
// ONE-TIME REPAIR: Fix Loyalty_Discount markers in SK_Orders
// ============================================================
/**
 * Run once from Apps Script editor to back-fill correct Loyalty_Discount
 * values for all customers. Safe to run multiple times (idempotent).
 *
 * What it does:
 *  1. Replays every customer's order history chronologically (same logic
 *     as submitOrder's virtual streak).
 *  2. Identifies which rows SHOULD have Loyalty_Discount = "Yes" based
 *     on the 6-consecutive-day rule.
 *  3. Writes "Yes" / "No" into the sheet only where the current value
 *     differs, so it doesn't thrash unchanged rows.
 *
 * Run from Editor: open Apps Script → select fixLoyaltyDiscountMarkers → Run
 */
function fixLoyaltyDiscountMarkers() {
  const ss  = getSpreadsheet();
  const ws  = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const data = ws.getDataRange().getValues();
  if (data.length < 2) { console.log("No data rows."); return; }

  const headers = data[0];
  const colIdx  = {};
  headers.forEach((h, i) => { colIdx[String(h).trim()] = i; });

  const COL_PHONE   = colIdx["Phone"];
  const COL_DATE    = colIdx["Order_Date"];
  const COL_STATUS  = colIdx["Payment_Status"];
  const COL_SUBTOT  = colIdx["Food_Subtotal"];
  const COL_LOYDISC = colIdx["Loyalty_Discount"];
  const COL_DISCAMT = colIdx["Discount_Amount"];

  if (COL_LOYDISC === undefined) {
    console.error("Loyalty_Discount column not found. Run initSchema() first.");
    return;
  }

  // ── 1. Group rows by phone ──────────────────────────────────
  const byPhone = {}; // phone → [{ rowIndex(1-based), date, subtotal, status, discAmt, current }]
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const phone  = _normalizePhone(String(row[COL_PHONE] || "").trim());
    if (!phone) continue;
    const stat   = String(row[COL_STATUS] || "").toLowerCase();
    const rawDate = row[COL_DATE];
    const dateStr = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Kolkata", "yyyy-MM-dd")
      : String(rawDate || "").trim();
    if (!dateStr) continue;

    if (!byPhone[phone]) byPhone[phone] = [];
    byPhone[phone].push({
      rowIndex: i + 1,          // 1-based sheet row
      date:     dateStr,
      subtotal: Number(row[COL_SUBTOT] || 0),
      discAmt:  Number(row[COL_DISCAMT] || 0),
      status:   stat,
      current:  String(row[COL_LOYDISC] || "").trim()
    });
  }

  // ── 2. For each customer, replay streak forward in time ─────
  const writes = []; // { rowIndex, value }
  let changed = 0, unchanged = 0;

  Object.entries(byPhone).forEach(([phone, rows]) => {
    // Sort chronologically, then by rowIndex (multiple meals same day)
    rows.sort((a, b) => a.date.localeCompare(b.date) || a.rowIndex - b.rowIndex);

    // Aggregate per day: group rows by date, pick only active orders
    const activeDayMap = {}; // date → [rows]
    rows.forEach(r => {
      const cancelled = r.status.includes("cancelled") || r.status.includes("deleted");
      if (cancelled) return;
      if (!activeDayMap[r.date]) activeDayMap[r.date] = [];
      activeDayMap[r.date].push(r);
    });

    // Get sorted unique active days
    const activeDates = Object.keys(activeDayMap).sort();

    let streakCount = 0; // consecutive active days going forward
    let prevDate    = null;

    activeDates.forEach(dateStr => {
      const d = new Date(dateStr + "T00:00:00+05:30");
      const dow = d.getDay(); // 0=Sun

      // Check continuity: is this date the "next expected" day after prevDate?
      let isContinuous = false;
      if (!prevDate) {
        isContinuous = true; // first ever day always starts streak
      } else {
        // Advance prevDate forward by 1+ days, skipping Sundays, to see if we land on dateStr
        const prev = new Date(prevDate + "T00:00:00+05:30");
        let nxt = new Date(prev); nxt.setDate(nxt.getDate() + 1);
        while (nxt.getDay() === 0) nxt.setDate(nxt.getDate() + 1); // skip Sundays
        const nxtISO = Utilities.formatDate(nxt, "Asia/Kolkata", "yyyy-MM-dd");
        isContinuous = (nxtISO === dateStr);
      }

      if (!isContinuous) {
        // Gap — reset streak
        streakCount = 0;
      }

      const is6thDay = (streakCount === 5); // 0-indexed: 0=1st, 5=6th

      // Mark all rows for this date
      activeDayMap[dateStr].forEach(r => {
        const expected = is6thDay ? "Yes" : "No";
        writes.push({ rowIndex: r.rowIndex, value: expected });
        if (r.current !== expected) changed++;
        else unchanged++;
      });

      if (is6thDay) {
        streakCount = 0; // reset after reward day
      } else {
        streakCount++;
      }

      prevDate = dateStr;
    });

    // Also mark cancelled rows on 6th-day dates as "No" (they didn't get reward)
    rows.forEach(r => {
      const cancelled = r.status.includes("cancelled") || r.status.includes("deleted");
      if (!cancelled) return;
      // Already handled above for active rows; just ensure cancelled don't have "Yes" erroneously
      if (r.current === "Yes") {
        writes.push({ rowIndex: r.rowIndex, value: "No" });
        changed++;
      }
    });
  });

  // ── 3. Batch-write all changes ──────────────────────────────
  writes.forEach(w => {
    ws.getRange(w.rowIndex, COL_LOYDISC + 1).setValue(w.value);
  });

  console.log(`fixLoyaltyDiscountMarkers complete.`);
  console.log(`  Rows updated : ${changed}`);
  console.log(`  Rows already correct: ${unchanged}`);
  console.log(`  Total writes : ${writes.length}`);
}
