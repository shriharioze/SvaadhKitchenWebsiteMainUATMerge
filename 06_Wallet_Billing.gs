// ============================================================
// 06_Wallet_Billing.gs
// Wallet ledger, manual UPI recharges (admin-verified), on-account
  billing, refunds, payment status management.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── WALLET HELPER ──────────────────────────────────────────
function _calculateWalletBalance(phone, preloadedRows) {
  if (!phone) return 0;
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const rows = Array.isArray(preloadedRows) ? preloadedRows : getAllRows(ws);

  let balance = 0;
  const pStr = _normalizePhone(phone);

  rows.forEach(w => {
    const rPhone = _normalizePhone(w.Phone);
    if (rPhone !== pStr) return;

    // Only count verified transactions
    const rVer = String(w.Verified || "").trim().toUpperCase();
    if (rVer !== "TRUE" && rVer !== "YES" && rVer !== "VERIFIED") return;

    const rAmt = _cleanNum(_get(w, "Amount"));
    // Also check legacy columns where Txn_Type may have been stored in a "Balance" column
    const rType = String(_get(w, "Txn_Type") || _get(w, "Balance") || "").trim().toLowerCase();

    if (rType.includes("recharge") || rType.includes("refund") || rType.includes("credit")
        || rType.includes("carry forward") || rType.includes("carry-forward")) {
      balance += rAmt;
    } else if (rType.includes("order") || rType.includes("deduct") || rType.includes("payment")) {
      balance -= rAmt;
    }
  });

  return Math.round(balance * 100) / 100;
}
function getWalletTransactions(phone) {
  if (!phone) return { transactions: [] };
  const ss   = getSpreadsheet();
  const ws   = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const rows = getAllRows(ws);
  const pStr = _normalizePhone(phone);

  // Filter to this customer's rows only, parse timestamps for sorting
  const mine = rows
    .filter(w => _normalizePhone(w.Phone) === pStr)
    .map(w => {
      const rType = String(w.Txn_Type || "").trim();
      const rAmt  = Number(w.Amount) || 0;
      const rVer  = String(w.Verified || "").trim().toUpperCase();
      const verified = (rVer === "TRUE" || rVer === "YES" || rVer === "VERIFIED");
      const typeLow  = rType.toLowerCase();
      const isCredit = typeLow.includes("recharge") || typeLow.includes("refund")
                    || typeLow.includes("credit") || typeLow.includes("carry forward")
                    || typeLow.includes("carry-forward");
      const rawTs  = w.Timestamp;
      const tsDate = rawTs instanceof Date ? rawTs : new Date(rawTs || 0);
      return {
        type:      rType || "Transaction",
        amount:    rAmt,
        direction: isCredit ? "credit" : "debit",
        verified,
        reference: String(w.Reference_ID || "").trim(),
        timestamp: rawTs instanceof Date
          ? Utilities.formatDate(rawTs, "Asia/Kolkata", "dd MMM yyyy, h:mm a")
          : String(rawTs || "").trim(),
        _ts: tsDate.getTime()
      };
    });

  // Sort newest first, take last 10
  mine.sort((a, b) => b._ts - a._ts);
  const top10 = mine.slice(0, 10).map(t => { delete t._ts; return t; });

  return { transactions: top10 };
}
// ── WALLET LOGIC ───────────────────────────────────────────────
/**
 * Append a transaction to SK_Wallet.
 * @param {string} phone      Customer phone number
 * @param {string} name       Customer name
 * @param {string} txnType    e.g. "Order Deduction", "Recharge", "Order Cancellation Refund"
 * @param {number} amount     Absolute transaction amount (always positive)
 * @param {boolean} isVerified TRUE = immediately counted in balance, FALSE = pending admin approval
 * @param {string} [refId]    Reference ID: Submission_ID for orders/refunds, or a recharge txn ref
 */
function _appendWalletTransaction(phone, name, txnType, amount, isVerified, refId) {
  // Serialize wallet writes. Apps Script LockService is re-entrant within the
  // same execution, so this also works when the caller (e.g. submitOrder) is
  // already holding the script lock.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { throw new Error("Wallet busy — please retry in a few seconds."); }
  try {
    const ss = getSpreadsheet();
    const ws = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
    const hIdx = headerIndex(ws);

    const totalCols = ws.getLastColumn();
    const row = new Array(totalCols).fill("");
    const set = (col, val) => { if (hIdx[col]) row[hIdx[col] - 1] = val; };

    set("Phone",         _normalizePhone(phone));
    set("Customer_Name", name);
    set("Txn_Type",      txnType);
    set("Amount",        amount);
    set("Verified",      isVerified ? "TRUE" : "FALSE");
    set("Reference_ID",  refId || "");
    set("Timestamp",     getISTTimestamp());

    ws.appendRow(row);
    SpreadsheetApp.flush();
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}
// ── REFUND MANAGEMENT (ADMIN) ────────────────────────────────
function getPendingRefunds() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_REFUNDS, ["Submission_ID","Phone","Name","Amount","Meal","Date","Status","Timestamp","Adjustment_Note","Refund_Mode"]);
  const rows = getAllRows(ws);
  return rows.filter(r => ["Pending", "Verification Required"].includes(String(r.Status)));
}
function markRefunded(submissionId) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_REFUNDS, []);
  const data = ws.getDataRange().getValues();
  const h = data[0];
  const idIdx = h.indexOf("Submission_ID");
  const statusIdx = h.indexOf("Status");
  const phoneIdx = h.indexOf("Phone");
  const nameIdx = h.indexOf("Name");
  const amtIdx = h.indexOf("Amount");
  const modeIdx = h.indexOf("Refund_Mode");

  if (idIdx === -1 || statusIdx === -1) return {success: false, error: "Sheet layout error"};

  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(submissionId)) {
      const row = data[i];
      const mode = modeIdx !== -1 ? String(row[modeIdx]).toLowerCase() : "upi";
      const phone = phoneIdx !== -1 ? String(row[phoneIdx]) : "";
      const name = nameIdx !== -1 ? String(row[nameIdx]) : "Customer";
      const amt = amtIdx !== -1 ? Number(row[amtIdx]) : 0;

      // Logic: If user chose Wallet refund, perform the ledger entry now
      if (mode === "wallet" && phone && amt > 0) {
        _appendWalletTransaction(phone, name, "Order Cancellation Refund", amt, true, String(submissionId));
      }

      // Mark the refund row as done
      ws.getRange(i + 1, statusIdx + 1).setValue("Refunded (" + now + ")");

      // ── Update the source order row remark to reflect completed refund ──
      // This closes the audit loop: the SK_Orders row was previously marked
      // "Cancelled – UPI Refund Pending" (or similar); now update it.
      try {
        const ordersWs = ss.getSheetByName(TAB_ORDERS);
        if (ordersWs) {
          const ordersData = ordersWs.getDataRange().getValues();
          const oHeaders = ordersData[0];
          const oIdIdx = oHeaders.indexOf("Submission_ID");
          const oStatusIdx = oHeaders.indexOf("Payment_Status");
          if (oIdIdx !== -1 && oStatusIdx !== -1) {
            for (var j = 1; j < ordersData.length; j++) {
              if (String(ordersData[j][oIdIdx]) === String(submissionId)) {
                const finalRemark = mode === "wallet"
                  ? "Cancelled \u2013 Refunded to Wallet (" + now + ")"
                  : "Cancelled \u2013 Refunded via UPI (" + now + ")";
                ordersWs.getRange(j + 1, oStatusIdx + 1).setValue(finalRemark);
                break;
              }
            }
          }
        }
      } catch(e) { /* non-fatal — refund row already updated */ }

      return {success: true};
    }
  }
  return {success: false, error: "Refund request not found"};
}
function markRefundRejected(submissionId) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_REFUNDS, []);
  const data = ws.getDataRange().getValues();
  const h = data[0];
  const idIdx = h.indexOf("Submission_ID");
  const statusIdx = h.indexOf("Status");

  if (idIdx === -1 || statusIdx === -1) return {success: false, error: "Sheet layout error"};

  const now = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(submissionId)) {
      ws.getRange(i + 1, statusIdx + 1).setValue("Rejected (" + now + ")");
      return {success: true};
    }
  }
  return {success: false, error: "Refund request not found"};
}
// ── GET UNPAID CUSTOMERS (reconciliation) ────────────────────────────────────
function getUnpaidCustomers(p) {
  const dateFrom = p.dateFrom;
  const dateTo   = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};

  const ss   = getSpreadsheet();
  const ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);

  // Collect all unpaid orders in the range
  const relevant = rows.filter(r =>
    String(r.Order_Date) >= dateFrom &&
    String(r.Order_Date) <= dateTo   &&
    (r.Payment_Status === "Pending" ||
     r.Payment_Status === "on account" ||
     !r.Payment_Status)
  );

  // Group by customer phone → sum net totals
  const map = {};
  relevant.forEach(r => {
    const key = String(r.Phone || "").trim();
    if (!key) return;
    if (!map[key]) map[key] = {phone:key, name:String(r.Customer_Name||"").trim(), total:0, orderCount:0};
    map[key].total      += Number(r.Net_Total) || 0;
    map[key].orderCount += 1;
  });

  // Load wallet rows ONCE — avoids N API calls inside the map loop
  const walletWsUC = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const walletRowsUC = getAllRows(walletWsUC);

  const customers = Object.values(map).map(c => {
    const wb = _calculateWalletBalance(c.phone, walletRowsUC);
    const net = Math.round(c.total - wb);
    return { ...c, total: net, walletBalance: Math.round(wb) };
  });
  const grandTotal = customers.reduce((s,c) => s + c.total, 0);
  return {success:true, customers, period:{from:dateFrom, to:dateTo}, grandTotal};
}
// ── MARK CUSTOMERS PAID (reconciliation) ─────────────────────────────────────
function markCustomersPaid(body) {
  const phones   = body.phones   || [];   // array of phone strings
  const dateFrom = body.dateFrom;
  const dateTo   = body.dateTo;
  if (!phones.length || !dateFrom || !dateTo)
    return {success:false, error:"phones, dateFrom, dateTo required"};

  const ss      = getSpreadsheet();
  const ws      = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const headers = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0];
  const hIdx    = {};
  headers.forEach((h,i) => { hIdx[h] = i+1; });

  const rows    = getAllRows(ws);
  let   updated = 0;
  rows.forEach(r => {
    if (phones.includes(String(r.Phone||"").trim()) &&
        String(r.Order_Date) >= dateFrom &&
        String(r.Order_Date) <= dateTo   &&
        (r.Payment_Status === "Pending" ||
         r.Payment_Status === "on account" ||
         !r.Payment_Status)) {
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue("Paid");
      updated++;
    }
  });
  return {success:true, updatedRows:updated, customersMarked:phones.length};
}
// ── GET DATE PAYMENTS ─────────────────────────────────────────────────────────
function getDatePayments(date) {
  if (!date) return {success:false, error:"date required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var rows = getAllRows(ws).filter(function(r){return fmtDate(r.Order_Date)===date && !_isOrderCancelled(r.Payment_Status);});

  var map = {};
  rows.forEach(function(r) {
    var phone = String(r.Phone||"").trim();
    if (!phone) return;
    if (!map[phone]) map[phone] = {phone:phone, name:String(r.Customer_Name||"").trim(),
      payFreq:String(r.Payment_Freq||"").trim(), meals:[], total:0, allPaid:true};
    map[phone].meals.push(r.Meal_Type);
    map[phone].total += Number(r.Net_Total)||0;
    if (!["Paid", "Wallet Paid", "Collected", "On Account"].includes(String(r.Payment_Status))) map[phone].allPaid = false;
  });

  var customers = Object.values(map).map(function(c) {
    return Object.assign({},c,{
      total:  Math.round(c.total),
      daily:  c.payFreq.toLowerCase().includes("daily"),
      status: c.allPaid ? "Paid" : "Pending"
    });
  }).sort(function(a,b){return a.name.localeCompare(b.name);});

  var grandTotal   = customers.reduce(function(s,c){return s+c.total;},0);
  var grandPaid    = customers.filter(function(c){return c.status==="Paid";}).reduce(function(s,c){return s+c.total;},0);

  return {success:true, date:date, customers:customers,
          grandTotal:grandTotal, grandPaid:grandPaid, grandPending:grandTotal-grandPaid};
}
// ── MARK ORDERS STATUS (by customer phone + date) ─────────────────────────────
function markOrdersStatus(body) {
  var date   = body.date;
  var phone  = body.phone;
  var sid    = body.sid; // Submission ID for precision
  var status = body.status || "Paid";
  if (!date || !phone) return {success:false, error:"date and phone required"};

  // Prevent race-condition double-processing (e.g. admin double-clicks Verify & Refund)
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch(e) { return {success:false, error:"Server busy — please retry"}; }
  try {

  var ss    = getSpreadsheet();
  var ws    = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var hIdx  = headerIndex(ws);
  var rows  = getAllRows(ws);
  
  var matches = rows.filter(function(r) {
    const d = r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date,"Asia/Kolkata","yyyy-MM-dd") : String(r.Order_Date).trim();
    return d === date && String(r.Phone||"").trim() === phone && (!sid || String(r.Submission_ID) === String(sid));
  });

  if (!matches.length) return {success: false, error: "No matching orders found"};

  var updated = 0;
  var now = getISTDate();
  
  // Sort descending by row index to allow safe deletion
  matches.sort((a,b) => b._row - a._row).forEach(function(r) {
    const currentStatus = String(r.Payment_Status || "").trim();
    if (currentStatus === "Cancelled (Verify UPI)" && status === "Paid") {
      // ── Process Refund Logic based on preference
      const pref = String(r.Refund_Preference || "upi").toLowerCase();
      const custName = r.Customer_Name || "Customer";

      // ── Recompute correct refund at verify time (same logic as hard-cancel) ──
      // Uses Net_Total as base so mealCredit baked into Net_Total is not over-refunded.
      const scOrderDate = r.Order_Date instanceof Date
        ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
        : String(r.Order_Date).trim();
      const scSameDayRows = rows.filter(x => {
        const xd = x.Order_Date instanceof Date
          ? Utilities.formatDate(x.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
          : String(x.Order_Date).trim();
        const xStat = String(x.Payment_Status || "").toLowerCase();
        return String(x.Phone).trim() === String(phone).trim() &&
               xd === scOrderDate &&
               String(x.Submission_ID) !== String(r.Submission_ID) &&
               !xStat.includes("deleted") && !xStat.includes("cancelled");
      });
      const scRemaining = scSameDayRows.reduce((s, x) => s + (Number(x.Food_Subtotal) || 0), 0);
      const scOldTotal  = scRemaining + (Number(r.Food_Subtotal) || 0);
      const scDiscRate  = (sub) => sub >= 450 ? 0.075 : sub >= 300 ? 0.05 : 0;
      const scOldRate   = scDiscRate(scOldTotal);
      const scNewRate   = scDiscRate(scRemaining);

      // Discount over-clawback on remaining rows
      let scOverDiscount = 0;
      if (scOldRate > scNewRate) {
        scOverDiscount = scSameDayRows.reduce((s, x) => {
          const xSub = Number(x.Food_Subtotal) || 0;
          return s + Math.round(xSub * scOldRate) - Math.round(xSub * scNewRate);
        }, 0);
        if (scOverDiscount > 0) {
          const scHIdx   = headerIndex(ws);
          const scDiscCol = scHIdx["Discount_Amount"];
          const scNetCol  = scHIdx["Net_Total"];
          scSameDayRows.forEach(x => {
            const xSub      = Number(x.Food_Subtotal)      || 0;
            const xSurcharge= Number(x.Inflation_Surcharge)|| 0;
            const xDelivery = Number(x.Delivery_Charge)    || 0;
            const xSmallFee = Number(x.Small_Order_Fee)    || 0;
            const xReviewD  = Number(x.Review_Discount)    || 0;
            const newD   = Math.round(xSub * scNewRate);
            const newNet = xSub + xDelivery + xSmallFee + xSurcharge - newD - xReviewD;
            if (scDiscCol) ws.getRange(x._row, scDiscCol).setValue(newD);
            if (scNetCol)  ws.getRange(x._row, scNetCol) .setValue(newNet);
          });
        }
      }

      // Delivery/small-fee clawback + row updates
      let scDeliveryOwed = 0;
      let scSmallFeeOwed = 0;
      const scFreeAreas  = getAreas().filter(a => a.free).map(a => a.name);
      const scIsNonFree  = (area) => !scFreeAreas.includes(area) && area !== "Self Pickup";
      const scFreeThr    = 150;
      if (scOldTotal >= scFreeThr && scRemaining < scFreeThr) {
        const scHIdx2    = headerIndex(ws);
        const scDelCol   = scHIdx2["Delivery_Charge"];
        const scSmallCol = scHIdx2["Small_Order_Fee"];
        const scNetCol2  = scHIdx2["Net_Total"];
        scSameDayRows.forEach(x => {
          const xSub  = Number(x.Food_Subtotal) || 0;
          const xMeal = String(x.Meal_Type).trim();
          let scNetDelta = 0;
          if (xSub > 0 && scIsNonFree(x.Area || "") && (Number(x.Delivery_Charge) || 0) === 0) {
            scDeliveryOwed += 10; scNetDelta += 10;
            if (scDelCol) ws.getRange(x._row, scDelCol).setValue(10);
          }
          if ((xMeal === "Lunch" || xMeal === "Dinner") && xSub > 0 && xSub < 50
              && (Number(x.Small_Order_Fee) || 0) === 0) {
            scSmallFeeOwed += 10; scNetDelta += 10;
            if (scSmallCol) ws.getRange(x._row, scSmallCol).setValue(10);
          }
          if (scNetDelta > 0 && scNetCol2) {
            ws.getRange(x._row, scNetCol2).setValue((Number(x.Net_Total) || 0) + scNetDelta);
          }
        });
      }

      const scAdj = scOverDiscount + scDeliveryOwed + scSmallFeeOwed;
      const amt   = Math.max(0, (Number(r.Net_Total) || 0) - scAdj);

      // ── Duplicate refund guard (shared for all paths below)
      const REF_HEADERS = ["Submission_ID","Phone","Name","Amount","Meal","Date","Status","Timestamp","Adjustment_Note","Refund_Mode"];
      const refWs = getOrCreateTab(ss, TAB_REFUNDS, REF_HEADERS);
      const existingRefunds = getAllRows(refWs);
      const alreadyExists = existingRefunds.some(rx => String(rx.Submission_ID) === String(r.Submission_ID));

      const isSplitOrder = String(r.Payment_Method || "").trim().toLowerCase() === "split";

      if (isSplitOrder) {
        // Split orders: entire refund always goes to Wallet — simple, no UPI queue.
        // Wallet portion was already deducted at order time, UPI portion was paid by customer.
        // Both come back to wallet in full.
        if (amt > 0) {
          _appendWalletTransaction(phone, custName, "Order Cancellation Refund", amt, true, String(r.Submission_ID));
        }
      } else if (pref === "wallet" && amt > 0) {
        _appendWalletTransaction(phone, custName, "Order Cancellation Refund", amt, true, String(r.Submission_ID));
      } else if (pref === "manual_upi" && amt > 0 && !alreadyExists) {
        refWs.appendRow([r.Submission_ID, phone, custName, amt, r.Meal_Type, date, "Pending", now, "Verified Soft Cancellation", "upi"]);
      }
      let finalCancelStatus = "Cancelled";
      if (isSplitOrder) {
        finalCancelStatus = "Cancelled \u2013 Refunded to Wallet";
      } else if (pref === "wallet") {
        finalCancelStatus = "Cancelled \u2013 Refunded to Wallet";
      } else if (pref === "manual_upi") {
        finalCancelStatus = "Cancelled \u2013 UPI Refund Pending";
      }
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue(finalCancelStatus);
    } else {
      // ── Standard Payment Approval or Rejection
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue(status);
    }
    updated++;
  });

  return {success:true, updatedRows:updated};
  } finally {
    lock.releaseLock();
  }
}
function rejectUPIPayment(body) {
  body.status = "Payment Rejected";
  return markOrdersStatus(body);
}
// ── ADMIN WALLET CREDIT ───────────────────────────────────────────────────────
function adminCreditWallet(body) {
  var phone  = String(body.phone || "").trim();
  var amount = Number(body.amount);
  if (!phone || phone.length < 10) return {success:false, error:"Valid phone required"};
  if (!amount || amount <= 0)      return {success:false, error:"Amount must be > 0"};

  // Look up customer name
  var ss      = getSpreadsheet();
  var profWs  = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var profRows = getAllRows(profWs);
  var profile  = profRows.find(function(r){ return String(r.Phone||"").trim() === phone; });
  var name     = profile ? (String(profile.Customer_Name||"").trim() || "Customer") : "Customer";

  _appendWalletTransaction(phone, name, "Admin Credit", amount, true, "ADMIN-" + Date.now());
  const settleRes = _autoSettlePendingOrders(phone);
  var newBalance = _calculateWalletBalance(phone);
  
  let msg = `₹${amount} credited to ${phone}. New balance: ₹${Math.round(newBalance)}`;
  if (settleRes.msg) {
    msg = settleRes.msg;
  }
  
  return {success:true, newBalance: Math.round(newBalance), msg: msg};
}
// ── WALLET TOPUP LOGIC ────────────────────────────────────────────────────────
function submitWalletRecharge(body) {
  var phone  = String(body.phone || "").trim();
  var name   = String(body.name || "").trim();
  var amount = Number(body.amount);
  if (!phone || isNaN(amount) || amount <= 0) return {success:false, error:"Invalid amount or phone"};

  // Unverified entry requiring admin to flip to TRUE
  const rechargeRef = "RCH-" + Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyyMMdd-HHmmss") + "-" + phone.slice(-4);
  _appendWalletTransaction(phone, name, "Recharge", amount, false, rechargeRef);
  
  return {success:true};
}
function getPendingRecharges() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const rawRows = getAllRows(ws);
  const pending = [];

  rawRows.forEach(w => {
    const rPhone = String(w.Phone || "").trim();
    const rName  = String(w.Customer_Name || "").trim();
    const rType  = String(w.Txn_Type || w.Balance || "").trim().toLowerCase();
    const rAmt   = Number(w.Amount) || 0;
    const rVer   = String(w.Verified || "").trim().toUpperCase();
    const rRef   = String(w.Reference_ID || "").trim();
    let   rTs    = w.Timestamp || "";
    if (rTs instanceof Date) rTs = Utilities.formatDate(rTs, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    else rTs = String(rTs).trim();

    if ((rVer === "FALSE" || !rVer) && rType.includes("recharge")) {
      pending.push({ 
        Phone: rPhone, 
        Customer_Name: rName, 
        Amount: rAmt, 
        Timestamp: rTs, 
        Reference_ID: rRef, 
        _row: w._row 
      });
    }
  });
  return pending;
}
function approveWalletRecharge(body) {
  const phone = String(body.phone || "").trim();
  const ts    = String(body.timestamp || "").trim();
  if (!phone || !ts) return {success:false, error:"Missing phone or timestamp"};

  const ss   = getSpreadsheet();
  const ws   = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const hIdx = headerIndex(ws);

  const vCol = hIdx["Verified"];
  const pCol = hIdx["Phone"];
  const tCol = hIdx["Timestamp"];
  if (!vCol || !pCol || !tCol) return {success:false, error:"Wallet sheet missing required columns"};

  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rPhone = String(rows[i][pCol-1] || "").trim();
    let   rTs    = rows[i][tCol-1];
    
    // Normalize timestamp for comparison
    if (rTs instanceof Date) {
      rTs = Utilities.formatDate(rTs, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    } else {
      rTs = String(rTs || "").trim();
    }
    
    const rVer = String(rows[i][vCol-1] || "").toUpperCase();

    // Check match
    if (rPhone === phone && rTs === ts) {
      if (rVer === "TRUE") return {success:true, msg:"Already verified"};
      ws.getRange(i+1, vCol).setValue("TRUE");
      const settleRes = _autoSettlePendingOrders(phone);
      return {success:true, msg: settleRes.msg || "Wallet Activated ✅"};
    }
  }
  
  Logger.log(`Activation Failed: No match for Phone: ${phone}, TS: ${ts}. Rows scanned: ${rows.length - 1}`);
  return {success:false, error:"Recharge request not found or already verified"};
}
function rejectWalletRecharge(body) {
  const phone = String(body.phone || "").trim();
  const ts    = String(body.timestamp || "").trim();
  if (!phone || !ts) return {success:false, error:"Missing phone or timestamp"};

  const ss   = getSpreadsheet();
  const ws   = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const hIdx = headerIndex(ws);

  const vCol = hIdx["Verified"];
  const pCol = hIdx["Phone"];
  const tCol = hIdx["Timestamp"];
  if (!vCol || !pCol || !tCol) return {success:false, error:"Wallet sheet missing required columns"};

  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rPhone = String(rows[i][pCol-1] || "").trim();
    let   rTs    = rows[i][tCol-1];
    
    // Normalize timestamp for comparison
    if (rTs instanceof Date) {
      rTs = Utilities.formatDate(rTs, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    } else {
      rTs = String(rTs || "").trim();
    }
    
    // Check match
    if (rPhone === phone && rTs === ts) {
      ws.getRange(i+1, vCol).setValue("REJECTED");
      return {success:true};
    }
  }
  
  return {success:false, error:"Recharge request not found"};
}
function batchProcessApprovals(body) {
  const { tab, action, items } = body;
  if (!tab || !action || !items || !items.length) return {success:false, error: "Invalid batch request"};
  
  const results = [];
  items.forEach(item => {
    let res;
    if (tab === 'refunds') {
      res = (action === 'approve') ? markRefunded(item.submissionId) : markRefundRejected(item.submissionId);
    } else if (tab === 'payments') {
      const payload = { ...item, status: (action === 'approve' ? 'Paid' : 'Payment Rejected') };
      res = markOrdersStatus(payload);
    } else if (tab === 'wallet') {
      res = (action === 'approve') ? approveWalletRecharge(item) : rejectWalletRecharge(item);
    }
    results.push(res);
  });
  
  const successCount = results.filter(r => r.success).length;
  return { success: true, total: items.length, successCount };
}
function getPendingUPIPayments() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  // Return Payment_Status == "Pending" OR "Cancelled (Verify UPI)"
  return rows.filter(r => {
    const s = String(r.Payment_Status).trim();
    const m = String(r.Payment_Method || "").trim();
    return s === "Pending" || s === "Cancelled (Verify UPI)" || s === "Pending Approval"
      || (m === "Split" && s === "Pending"); // Split orders awaiting UPI portion
  })
             .map(r => {
               const walletCredit = Number(r.Wallet_Credit) || 0;
               const isSplit = String(r.Payment_Method || "").trim() === "Split";
               const loyaltyDiscount = Number(r.Discount_Amount) || 0;
               const isLoyalty       = String(r.Loyalty_Discount || "").trim().toLowerCase() === "yes";
               return {
                 id: r.Submission_ID,
                 date: r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd") : r.Order_Date,
                 customer: r.Customer_Name,
                 phone: r.Phone,
                 amount: isSplit ? Math.max(0, (Number(r.Net_Total) || 0) - walletCredit) : r.Net_Total,
                 full_amount: r.Net_Total,
                 wallet_credit: walletCredit,
                 meal: r.Meal_Type,
                 timestamp: r.Submitted_At,
                 status: r.Payment_Status,
                 payment_method: String(r.Payment_Method || ""),
                 refund_preference: r.Refund_Preference || "",
                 loyalty_discount: isLoyalty ? loyaltyDiscount : 0,
                 is_loyalty: isLoyalty
               };
             });
}
// ═══════════════════════════════════════════════════════
// SENIOR BILLING — On Account Orders
// ═══════════════════════════════════════════════════════

function getBillingData(cycle, filterValue) {
  const ss = getSpreadsheet();
  const ordersWs  = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const custWs    = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const allOrders = getAllRows(ordersWs);
  const allCusts  = getAllRows(custWs);

  // Build customer map: phone → { billing_cycle, address, name }
  const custMap = {};
  allCusts.forEach(c => {
    const phone = String(c.Phone || '').trim();
    if (phone) custMap[phone] = {
      billing_cycle: String(c.Billing_Cycle || '').trim(),
      name:    c.Customer_Name || '',
      area:    c.Area || '',
      society: c.Society || '',
      wing:    c.Wing || '',
      flat:    c.Flat || '',
      floor:   c.Floor || ''
    };
  });

  // Compute date range based on cycle and filter (IST context)
  const now = getISTDate();
  let fromStr = '';
  let toStr   = '';

  // Helper to parse "YYYY-MM-DD" string reliably
  const parseYMD = (s) => {
    if (!s) return null;
    const p = s.split('-');
    return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  };

  if (cycle === 'Daily') {
    // Daily mode: show ALL pending On Account orders (no date restriction).
    // fromStr/toStr left blank — filtering below is skipped for Daily.
    fromStr = '';
    toStr   = '';
  } else if (cycle === 'Monthly') {
    const mIdx = (filterValue !== undefined && filterValue !== '') ? parseInt(filterValue) : now.getMonth();
    const first = new Date(now.getFullYear(), mIdx, 1);
    const last  = new Date(now.getFullYear(), mIdx + 1, 0);
    fromStr = Utilities.formatDate(first, 'Asia/Kolkata', 'yyyy-MM-dd');
    toStr   = Utilities.formatDate(last, 'Asia/Kolkata', 'yyyy-MM-dd');
  } else if (cycle === 'Weekly') {
    // filterValue is a date string YYYY-MM-DD
    let baseDate = parseYMD(filterValue) || now;
    
    // Find Monday of this week (Mon-Sat cycle)
    const day = baseDate.getDay(); // 0=Sun, 1=Mon...
    const diff = (day === 0 ? -6 : 1 - day); // Distance to Monday
    const mon = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + diff);
    
    // Saturday is Monday + 5 days
    const sat = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 5);
    
    fromStr = Utilities.formatDate(mon, 'Asia/Kolkata', 'yyyy-MM-dd');
    toStr   = Utilities.formatDate(sat, 'Asia/Kolkata', 'yyyy-MM-dd');
  } else {
    fromStr = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
    toStr   = fromStr;
  }

  // Filter On Account orders within cycle date range
  // For Daily: no date restriction — return ALL pending On Account orders.
  const onAccountOrders = allOrders.filter(r => {
    const status = String(r.Payment_Status || '').trim().toLowerCase();
    if (status !== 'on account') return false;
    if (cycle === 'Daily') return true; // all dates
    const dVal = r.Order_Date;
    const ds = dVal instanceof Date ? Utilities.formatDate(dVal, 'Asia/Kolkata', 'yyyy-MM-dd') : String(dVal).trim();
    return ds >= fromStr && ds <= toStr;
  });

  // Group by customer phone
  const grouped = {};
  onAccountOrders.forEach(r => {
    const phone = String(r.Phone || '').trim();
    const cust  = custMap[phone] || {};
    // Only include customers whose billing_cycle matches requested cycle
    if ((cust.billing_cycle || '').toLowerCase() !== cycle.toLowerCase()) return;

    if (!grouped[phone]) {
      grouped[phone] = {
        phone,
        name:    r.Customer_Name || cust.name || '',
        area:    r.Area || cust.area || '',
        society: r.Society || cust.society || '',
        wing:    r.Wing || cust.wing || '',
        flat:    r.Flat || cust.flat || '',
        floor:   r.Floor || cust.floor || '',
        billing_cycle: cust.billing_cycle || cycle,
        from: fromStr,
        to:   toStr,
        orders: [],
        total: 0
      };
    }

    const oDate = r.Order_Date;
    const ds    = oDate instanceof Date ? Utilities.formatDate(oDate, 'Asia/Kolkata', 'yyyy-MM-dd') : String(oDate).trim();

    grouped[phone].orders.push({
      sid:   String(r.Submission_ID || ''),
      date:  ds,
      meal:  String(r.Meal_Type || ''),
      items: String(r.Items_JSON || '{}'),
      net:   Number(r.Net_Total || 0),
      // Include for Daily flat-list display
      customer_name:  r.Customer_Name || cust.name || '',
      customer_phone: phone
    });
    grouped[phone].total += Number(r.Net_Total || 0);
  });

  // Default sorting: Total Amount Descending, then Name
  const customers = Object.values(grouped).sort((a,b) => {
    if (b.total !== a.total) return b.total - a.total;
    return a.name.localeCompare(b.name);
  });

  // For Daily: also build a flat, date-sorted list of all orders across all customers.
  // The frontend uses this to render the new flat-list view.
  let flat_orders = null;
  if (cycle === 'Daily') {
    flat_orders = [];
    customers.forEach(c => {
      c.orders.forEach(o => {
        flat_orders.push({
          sid:            o.sid,
          date:           o.date,
          meal:           o.meal,
          items:          o.items,
          net:            o.net,
          customer_name:  c.name,
          customer_phone: c.phone,
          area:           c.area,
          billing_cycle:  c.billing_cycle
        });
      });
    });
    // Sort by date ascending, then meal order (Breakfast→Lunch→Dinner)
    const mealOrder = { Breakfast: 0, Lunch: 1, Dinner: 2 };
    flat_orders.sort((a, b) => {
      const dateCmp = a.date.localeCompare(b.date);
      if (dateCmp !== 0) return dateCmp;
      return (mealOrder[a.meal] ?? 9) - (mealOrder[b.meal] ?? 9);
    });
  }

  return { success: true, cycle, from: fromStr, to: toStr, customers, flat_orders };
}
function markBillingCollected(submissionIds) {
  if (!submissionIds || !submissionIds.length) return { success: false, error: 'No submission IDs provided' };
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const hIdx = headerIndex(ws);
  const rows = getAllRows(ws);
  const statusCol = hIdx['Payment_Status'];
  if (!statusCol) return { success: false, error: 'Payment_Status column not found' };

  let count = 0;
  // Compare IDs exactly (trimmed) — do NOT strip non-digits; manual order IDs
  // contain mixed alphanumeric characters and digit-stripping causes false matches.
  const idSet = new Set(submissionIds.map(id => String(id).trim()));

  rows.forEach(r => {
    const cleanId = String(r.Submission_ID || '').trim();
    if (idSet.has(cleanId)) {
      ws.getRange(r._row, statusCol).setValue('Paid'); // Changed from 'Collected' to 'Paid' for uniformity
      count++;
    }
  });
  SpreadsheetApp.flush();
  return { success: true, count };
}
function undoMarkPaid(submissionIds) {
  if (!submissionIds || !submissionIds.length) return { success: false, error: 'No IDs' };
  const ss  = getSpreadsheet();
  const ws  = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const hIdx = headerIndex(ws);
  const rows = getAllRows(ws);
  const statusCol = hIdx['Payment_Status'];
  if (!statusCol) return { success: false, error: 'Column missing' };
  const idSet = new Set(submissionIds.map(id => String(id).trim()));
  let count = 0;
  rows.forEach(r => {
    if (idSet.has(String(r.Submission_ID || '').trim())) {
      ws.getRange(r._row, statusCol).setValue('Pending');
      count++;
    }
  });
  SpreadsheetApp.flush();
  return { success: true, count };
}
function toggleFeeExempt(phone, status) {
  const ss = getSpreadsheet();
  const custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const rows = getAllRows(custWs);
  const hIdx = headerIndex(custWs);
  const phoneStr = _normalizePhone(phone);
  const idx = rows.findIndex(r => _normalizePhone(r.Phone) === phoneStr);
  
  const val = (status === true || String(status).toLowerCase() === "yes") ? "Yes" : "No";
  
  if (idx !== -1) {
    custWs.getRange(idx + 2, hIdx["Fee_Exempt"]).setValue(val);
  } else {
    // Number not found - create a FUTURE whitelist entry (Predetermine)
    if (!phone || phone.length < 10) return { success: false, error: "Invalid phone number" };
    const row = new Array(CUSTOMERS_HEADERS.length).fill("");
    row[hIdx["Phone"] - 1] = phone;
    row[hIdx["Fee_Exempt"] - 1] = val;
    row[hIdx["Created_At"] - 1] = getISTTimestamp();
    row[hIdx["Customer_Name"] - 1] = "";
    custWs.appendRow(row);
  }
  return { success: true, status: val };
}
