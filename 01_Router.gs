// ============================================================
// 01_Router.gs
// doGet and doPost — entry points + action dispatching.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;
  const action = p.parameter ? p.action : (e.parameter.action || ""); // Fix for inconsistent parameter access
  const pin = p.pin || "";

  // Auth tiers resolved FIRST so every route below can safely reference them
  const isAdmin = _pinMatch(pin, ADMIN_PIN) && pin !== "";
  const isStaff = (_pinMatch(pin, KITCHEN_PIN) || _pinMatch(pin, ADMIN_PIN)) && pin !== "";

  // ── HDFC Return URL via GET ────────────────────────────────────
  // HDFC sometimes redirects the customer's browser via GET (not POST).
  // Detect by presence of order_id + status params with no _action.
  // Redirect browser to the order page URL with all params forwarded.
  if (p.order_id && p.status && !p.action && !p._action) {
    const params = Object.keys(p)
      .map(function(k) { return encodeURIComponent(k) + "=" + encodeURIComponent(p[k]); })
      .join("&");
    const redirectUrl = HDFC_ORDER_PAGE_URL + "?" + params;
    // Sandbox-aware full-viewport click target — auto-navigation is blocked
    // inside the Apps Script HtmlService iframe in many browsers, so we make
    // the entire page a tap/click target while attempting auto-redirect.
    return HtmlService.createHtmlOutput(_hdfcReturnRedirectHtml(redirectUrl));
  }
  // ─────────────────────────────────────────────────────────────

  try {
    if (action === "version") return jsonRes({version: CODE_VERSION, status:"ok"});
    if (action === "health") {
      // Lightweight liveness probe — reads one cell to confirm sheet connectivity.
      // Does NOT load orders or menu. Safe to call frequently from monitors.
      try {
        const ss = getSpreadsheet();
        const sheetCount = ss.getNumSheets();
        return jsonRes({ status: "ok", version: CODE_VERSION, sheets: sheetCount, ts: new Date().toISOString() });
      } catch(hErr) {
        return jsonRes({ status: "error", error: hErr.message, ts: new Date().toISOString() });
      }
    }
    if (action === "getConfig") return jsonRes({
      gateway_enabled: PAYMENT_GATEWAY_ENABLED,
      gateway_env: HDFC_ENV
    });
    if (action === "getAreas") return jsonRes(getAreas());
    if (action === "getCustomer") return jsonRes(getCustomer(p.phone));
    if (action === "verifyLogin") return jsonRes(verifyLogin(p.phone, p.pin));
    if (action === "setPin") {
      const profile = { phone: p.phone, pin: p.pin };
      _upsertCustomer(getSpreadsheet(), profile);
      return jsonRes({success:true});
    }
    if (action === "getWeeklyMenu") return jsonRes(getWeeklyMenu());
    if (action === "markOnAccount") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markOnAccount(p.phone, p.cycle, p.status));
    }
    
    // KITCHEN & DRIVER ACCESS (Staff PIN ONLY)
    if (action === "getKitchenSummary") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(getKitchenSummary(p.date));
    }
    if (action === "getDriverOrders") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(getDriverOrders(p.date));
    }
    if (action === "createDeliverySheet") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(createDeliverySheet(p.date, p.meal));
    }
    if (action === "getLabelOrders") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(getLabelOrders(p.date, p.meal));
    }

    // FULL ADMIN ACCESS (Admin PIN ONLY)
    if (action === "getAdminData") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getAdminData());
    }
    if (action === "getUnpaidCustomers") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getUnpaidCustomers(p));
    }
    if (action === "getOrderSummary") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getOrderSummary(p.date));
    }
    if (action === "getPackagingExpenses") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      if (p.from && p.to) return jsonRes(getPackagingExpensesRange(p.from, p.to));
      return jsonRes(getPackagingExpenses(p.date));
    }
    if (action === "getOrderHistory") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getOrderHistory(p));
    }
    if (action === "getCustomerList") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getCustomerList());
    }
    if (action === "getCustomerHistory") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getCustomerHistory(p.phone));
    }
    if (action === "getDatePayments") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getDatePayments(p.date));
    }
    if (action === "getAnalytics") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getAnalytics(p));
    }
    if (action === "getExpenses") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getExpenses(p));
    }
    if (action === "getExpenseAnalytics") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getExpenseAnalytics(p));
    }
    if (action === "getCustomExpenseCategories") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes({ success:true, categories: getCustomExpenseCategories() });
    }
    if (action === "getInventoryData") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getInventoryData(p));
    }
    if (action === "adminCreditWallet") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(adminCreditWallet(body));
    }
    if (action === "getChurnReport") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getChurnReport(p.sinceDate));
    }
    if (action === "getPendingRefunds") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getPendingRefunds());
    }
    if (action === "getPendingRecharges") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(getPendingRecharges());
    }
    if (action === "getPendingUPIPayments") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(getPendingUPIPayments());
    }
    if (action === "getPendingCounts") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes({
        refunds: getPendingRefunds().length,
        wallet: getPendingRecharges().length,
        payments: getPendingUPIPayments().length
      });
    }

    if (action === "syncGA4Data") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes({success: true, message: syncGA4Data()});
    }
    if (action === "setupAnalyticsTrigger") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes({success: true, message: setupAnalyticsTrigger()});
    }
    
    // Keep-alive ping — just wakes GAS, no sheet reads
    if (action === "ping") return jsonRes({ok: true, t: new Date().toISOString()});

    // Fallback menu / orders for customers (legacy)
    if (action === "getMenu") return jsonRes(getMenu(p.date));
    if (action === "getMenuBatch") return jsonRes(getMenuBatch(p.dates));
    if (action === "getWeeklyMenu") return jsonRes(getWeeklyMenu());
    if (action === "getCustomerOrders") return jsonRes(getCustomerOrders(p.phone));
    if (action === "getWalletValue") return jsonRes({wallet_balance: _calculateWalletBalance(p.phone)});
    if (action === "getWalletTransactions") return jsonRes(getWalletTransactions(p.phone));
    if (action === "getDayTotalsForDates") return jsonRes(getDayTotalsForDates(p.phone, p.dates));

    return jsonRes({error:"Unknown action or Access Denied"});
  } catch(err) {
    return jsonRes({error: err.message});
  }
}
function doPost(e) {
  try {
    // ── HDFC Return URL Handler ────────────────────────────────
    // Juspay POSTs payment result to our return URL (GitHub Pages can't accept POST → 405).
    // We use the Apps Script URL as the return URL instead.
    // When HDFC posts here with order_id + status (no _action), serve an HTML page
    // that immediately JS-redirects the browser to order.html with those params as GET params.
    const rawBody = e.postData ? e.postData.contents : "";
    let parsedForHdfc = {};
    try { parsedForHdfc = JSON.parse(rawBody); } catch(_) {}
    const isHdfcReturn = parsedForHdfc.order_id && parsedForHdfc.status && !parsedForHdfc._action;
    if (isHdfcReturn) {
      const params = Object.keys(parsedForHdfc)
        .map(k => encodeURIComponent(k) + "=" + encodeURIComponent(parsedForHdfc[k]))
        .join("&");
      const redirectUrl = HDFC_ORDER_PAGE_URL + "?" + params;
      return HtmlService.createHtmlOutput(_hdfcReturnRedirectHtml(redirectUrl));
    }
    // ── Also handle form-encoded POST (Juspay sometimes sends application/x-www-form-urlencoded)
    if (!parsedForHdfc.order_id && e.postData && e.postData.type === "application/x-www-form-urlencoded") {
      const formParams = e.parameter || {};
      if (formParams.order_id && formParams.status) {
        const params = Object.keys(formParams)
          .map(k => encodeURIComponent(k) + "=" + encodeURIComponent(formParams[k]))
          .join("&");
        const redirectUrl = HDFC_ORDER_PAGE_URL + "?" + params;
        return HtmlService.createHtmlOutput(
          `<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0;url=${redirectUrl}"></head>` +
          `<body><script>window.location.replace(${JSON.stringify(redirectUrl)});</script>` +
          `<p>Redirecting... <a href="${redirectUrl}">Click here if not redirected</a></p></body></html>`
        );
      }
    }
    // ── Normal API actions ─────────────────────────────────────
    const body = JSON.parse(rawBody);
    const action = body._action || "";
    const pin = body.pin || "";
    const isAdmin = _pinMatch(pin, ADMIN_PIN) && pin !== "";
    const isStaff = (_pinMatch(pin, KITCHEN_PIN) || _pinMatch(pin, ADMIN_PIN)) && pin !== "";

    // Customer self-service (phone-verified inside each function)
    if (action === "deleteOrder") return jsonRes(deleteOrder(body.phone, body.rowId, body.refundType, { isAdmin: isAdmin }));
    if (action === "previewCancellation") return jsonRes(_deleteOrderInternal(body.phone, body.rowId, body.refundType || "wallet", { dryRun: true }));
    if (action === "getCustomerOrders") return jsonRes(getCustomerOrders(body.phone));
    if (action === "verifyOrderPlaced") return jsonRes(verifyOrderPlaced(body));
    if (action === "updateProfile") {
      const profile = body.profile;
      if (!profile || !profile.phone) return jsonRes({error: "Phone required"});
      _upsertCustomer(getSpreadsheet(), profile);
      return jsonRes({success: true});
    }

    // Admin-only read: returns all customer profiles + wallet balances
    if (action === "getCustomerList") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getCustomerList());
    }
    
    // Delivery Actions (Staff PIN ONLY)
    if (action === "markDelivered") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(markDelivered(body));
    }
    if (action === "batchMarkEnRoute") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(batchMarkEnRoute(body));
    }
    if (action === "setStandardOrder") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(setStandardOrder(body.phone, body.items, body.templateName, body.meal));
    }
    if (action === "removeStandardOrder") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(removeStandardOrder(body.phone, body.templateName));
    }
    if (action === "placeBulkOrders") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(placeBulkOrders(body));
    }
    if (action === "markEnRoute") {
       if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
       return jsonRes(markEnRoute(body));
     }
     if (action === "markOrderPacked") {
       if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
       return jsonRes(markOrderPacked(body));
     }

    // Admin-only write actions
    if (action === "adminCancelOrder") {
      if (!isAdmin) return jsonRes({success:false, error: "STRICT ADMIN PIN REQUIRED"});
      return jsonRes(adminCancelOrder(body));
    }
    if (action === "markRefunded") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markRefunded(body.submissionId));
    }
    if (action === "toggleFeeExempt") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(toggleFeeExempt(body.phone, body.status));
    }
    if (action === "approveWalletRecharge") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(approveWalletRecharge(body));
    }
    if (action === "markReviewed") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markReviewed(body));
    }
    if (action === "deleteBreakfastItem") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteBreakfastItem(body.id));
    }
    if (action === "saveBreakfastItem") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveBreakfastItem(body));
    }
    if (action === "deleteSabjiItem") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteSabjiItem(body.id));
    }
    if (action === "seedTestData") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes({success:true, message: seedTestData()});
    }
    if (action === "saveSabjiItem") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveSabjiItem(body));
    }
    if (action === "saveLabels") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveLabels(body));
    }
    if (action === "markCustomersPaid") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(markCustomersPaid(body));
    }
    if (action === "markOrdersStatus") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(markOrdersStatus(body));
    }
    if (action === "markOnAccount") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markOnAccount(body.phone, body.cycle, body.status));
    }
    if (action === "getBillingData") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(getBillingData(body.cycle, body.filterValue));
    }
    if (action === "markBillingCollected") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markBillingCollected(body.submissionIds));
    }
    if (action === "undoMarkPaid") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(undoMarkPaid(body.submissionIds));
    }
    if (action === "saveArea") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveArea(body));
    }
    if (action === "deleteArea") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteArea(body));
    }
    if (action === "markCustomersPaid") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(markCustomersPaid(body));
    }
    if (action === "markOrdersStatus") {
      if (!isAdmin) return jsonRes({error:"Invalid PIN"});
      return jsonRes(markOrdersStatus(body));
    }
    if (action === "getReviews") return jsonRes(getReviews());
    if (action === "chat") return jsonRes(handleChat(body));
    if (action === "submitWalletRecharge") return jsonRes(submitWalletRecharge(body));
    if (action === "markRefundRejected") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(markRefundRejected(body.submissionId));
    }
    if (action === "rejectUPIPayment") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(rejectUPIPayment(body));
    }
    if (action === "adminCreditWallet") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(adminCreditWallet(body));
    }
    if (action === "rejectWalletRecharge") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(rejectWalletRecharge(body));
    }
    if (action === "batchProcessApprovals") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(batchProcessApprovals(body));
    }
    if (action === "saveInventoryEntry") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveInventoryEntry(body));
    }
    if (action === "deleteInventoryEntry") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteInventoryEntry(body));
    }
    if (action === "saveCustomExpenseCategory") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveCustomExpenseCategory(body));
    }
    if (action === "deleteCustomExpenseCategory") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteCustomExpenseCategory(body));
    }
    if (action === "saveExpense") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveExpense(body));
    }
    if (action === "deleteExpense") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(deleteExpense(body));
    }
    if (action === "triggerManualArchive") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(triggerManualArchive(body));
    }
    if (action === "setupQuarterlyArchiveTrigger") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      try { setupQuarterlyArchiveTrigger(); return jsonRes({success:true}); }
      catch(e) { return jsonRes({success:false, error:e.message}); }
    }

    if (action === "setPin") {
      const profile = { phone: body.phone, pin: body.pin };
      _upsertCustomer(getSpreadsheet(), profile);
      return jsonRes({success:true});
    }

    if (action === "saveMenu") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(saveMenu(body));
    }

    if (action === "upsertProfile") {
      // Capture PIN if provided during mid-flow profile upserts
      const profile = { ...body, pin: body.pin || "" };
      _upsertCustomer(getSpreadsheet(), profile);
      return jsonRes({success:true});
    }

    if (action === "submitManualOrder") {
      if (!isAdmin) return jsonRes({error:"STRICT ADMIN PIN REQUIRED"});
      return jsonRes(submitManualOrder(body));
    }

    // Client error logging (timeout / network failures reported by frontend)
    if (action === "logClientError") return jsonRes(logClientError(body));

    // ── HDFC PAYMENT GATEWAY ACTIONS ─────────────────────────
    // All gateway actions are gated by PAYMENT_GATEWAY_ENABLED.
    // The webhook action is the only one that uses its own auth (Basic Auth
    // from HDFC's server), not the customer or admin PIN.

    if (action === "hdfc_createSession") {
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_createSession(body));
    }
    if (action === "hdfc_savePendingOrder") return jsonRes(hdfc_savePendingOrder(body));
    if (action === "hdfc_getPendingOrder")  return jsonRes(hdfc_getPendingOrder(body));

    // Wallet top-up via HDFC SmartGateway (separate from order payment).
    if (action === "hdfc_createWalletRechargeSession") {
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_createWalletRechargeSession(body));
    }
    if (action === "hdfc_finalizeWalletRecharge") {
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_finalizeWalletRecharge(body.order_id));
    }

    if (action === "hdfc_webhook") {
      // HDFC posts to this URL with Basic Auth — verify credentials first
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_handleWebhook(body, e));
    }

    // ── HDFC Webhook auto-detect ───────────────────────────────
    // HDFC's server-side webhook POST will NOT contain _action.
    // Detect by presence of event_name (Juspay webhook signature field).
    if (!action && (body.event_name || (body.content && body.content.order))) {
      console.log("HDFC Webhook auto-detected, event:", body.event_name);
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_handleWebhook(body, e));
    }

    if (action === "hdfc_verifyReturn") {
      // Called by order.html when customer lands back after payment
      if (!PAYMENT_GATEWAY_ENABLED) return jsonRes({error:"Payment gateway not enabled."});
      return jsonRes(hdfc_verifyReturnPayload(body));
    }
    // ─────────────────────────────────────────────────────────

    // Regular order submission
    return jsonRes(submitOrder(body));
  } catch(err) {
    return jsonRes({error: err.message});
  }
}
