// ============================================================
// 11_Hdfc_Reconciler.gs
// Self-healing reconciliation for HDFC payments that charged
// successfully at the gateway but never wrote a row in SK_Orders
// (e.g. user closed the popup before the post-charge round-trip
// completed). A 5-minute time-based trigger sweeps the pending
// log, confirms each entry against the Status API / Webhook Log,
// and writes the SK_Orders row using the cached cart state.
// ============================================================
// Gated by PAYMENT_GATEWAY_ENABLED — safe to deploy on live.
// Sourced from SvaadhKitchenUAT v14.8.
// ============================================================

function setupReconcileTrigger() {
  // Remove any existing reconcile triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "reconcilePendingOrders") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("reconcilePendingOrders")
    .timeBased()
    .everyMinutes(5)
    .create();
  return "Reconcile trigger set — runs every 5 minutes.";
}

function reconcilePendingOrders() {
  if (!PAYMENT_GATEWAY_ENABLED) { Logger.log("reconcilePendingOrders: gateway disabled, skipping."); return; }

  const props   = PropertiesService.getScriptProperties();
  const raw     = props.getProperty("HDFC_PENDING_ORDERS") || "{}";
  var pending;
  try { pending = JSON.parse(raw); } catch (e) { Logger.log("reconcilePendingOrders: malformed JSON, aborting."); return; }

  const orderIds = Object.keys(pending);
  if (!orderIds.length) { Logger.log("reconcilePendingOrders: no pending entries, nothing to do."); return; }

  const now = Date.now();
  const summary = { checked: 0, skippedFresh: 0, skippedAlreadyDone: 0, skippedNotCharged: 0, reconciled: 0, errors: 0 };

  for (var i = 0; i < orderIds.length; i++) {
    const orderId = orderIds[i];
    const entry   = pending[orderId];
    const ageMs   = now - (entry.ts || 0);

    // Skip very-fresh entries — let the customer's browser finish first
    if (ageMs < 2 * 60 * 1000) { summary.skippedFresh++; continue; }
    summary.checked++;

    try {
      const result = _reconcileSingleEntry(orderId, entry);
      summary[result.outcome] = (summary[result.outcome] || 0) + 1;
      // If reconciled (or already done), remove from pending so we don't keep checking
      if (result.outcome === "reconciled" || result.outcome === "skippedAlreadyDone") {
        delete pending[orderId];
      }
    } catch (e) {
      summary.errors++;
      Logger.log("reconcilePendingOrders: error on " + orderId + " — " + e.message);
    }
  }

  // Persist any deletions
  try { props.setProperty("HDFC_PENDING_ORDERS", JSON.stringify(pending)); } catch(_) {}

  Logger.log("reconcilePendingOrders summary: " + JSON.stringify(summary));
  return summary;
}

function _reconcileSingleEntry(orderId, entry) {
  // Per-order lock so we never race with the customer's own verification
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15 * 1000); } catch (e) {
    return { outcome: "skippedFresh", reason: "lock-contention" };
  }

  try {
    // ── Wallet recharge? Route to recharge finalizer ────────────────────────
    if (/^SK\d{6}W/.test(orderId)) {
      const r = hdfc_finalizeWalletRecharge(orderId);
      if (r.success && !r.already_credited) return { outcome: "reconciled", kind: "recharge" };
      if (r.success && r.already_credited)  return { outcome: "skippedAlreadyDone", kind: "recharge" };
      return { outcome: "skippedNotCharged", kind: "recharge", reason: r.error || "not confirmed" };
    }

    // ── Regular / split order: check SK_Orders dedup first ───────────────────
    const ss     = getSpreadsheet();
    const ws     = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
    const data   = ws.getDataRange().getValues();
    const headers= data[0] || [];
    const gCol   = headers.indexOf("Gateway_Order_ID");
    if (gCol !== -1) {
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][gCol] || "").trim() === orderId) {
          return { outcome: "skippedAlreadyDone", row: r + 1 };
        }
      }
    }

    // ── Ask HDFC: is this actually CHARGED? ─────────────────────────────────
    // Primary: Status API call. Fallback: SK_Webhook_Log (HDFC's own
    // server-to-server ORDER_SUCCEEDED event, equally authoritative).
    // The fallback fires when the Status API can't be reached due to
    // urlfetch quota exhaustion or transient errors — without it, a
    // quota-exhausted day would block ALL stuck-order recovery until
    // midnight PST.
    var statusCheck;
    try { statusCheck = hdfc_getOrderStatus(orderId); }
    catch (e) { statusCheck = { confirmed: false, status: "FETCH_ERROR", amount: 0 }; }

    if (!statusCheck.confirmed) {
      const transient = (statusCheck.status === "FETCH_ERROR" ||
                         statusCheck.status === "API_ERROR" ||
                         statusCheck.status === "UNKNOWN" ||
                         statusCheck.status === "NEW");
      if (transient) {
        const webhookProof = _checkWebhookLogForCharge(orderId);
        if (webhookProof) {
          Logger.log("reconcile: " + orderId + " — Status API unavailable (" + statusCheck.status + "), but ORDER_SUCCEEDED webhook found in SK_Webhook_Log. Trusting webhook.");
          statusCheck = { confirmed: true, status: "CHARGED", amount: webhookProof.amount };
        }
      }
    }
    if (!statusCheck.confirmed) {
      return { outcome: "skippedNotCharged", status: statusCheck.status };
    }

    // ── Synthesize the submitOrder body from the pending entry ──────────────
    const body = _buildSubmitBodyFromPending(orderId, entry, statusCheck);
    if (!body || !body.orders || !body.orders.length) {
      return { outcome: "errors", reason: "empty orders in pending entry" };
    }

    // ── Submit ──────────────────────────────────────────────────────────────
    const subResult = submitOrder(body);
    if (subResult && (subResult.success || subResult.submission_id || subResult.submissionIds)) {
      Logger.log("reconcile: order " + orderId + " written to sheet. result=" + JSON.stringify(subResult));
      return { outcome: "reconciled", subResult: subResult };
    }
    return { outcome: "errors", reason: "submitOrder did not succeed: " + JSON.stringify(subResult) };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/**
 * Transform the pending-entry shape (S.orders from the frontend) into
 * the body shape that submitOrder() expects.
 */
function _buildSubmitBodyFromPending(orderId, entry, statusCheck) {
  const profile = entry.profile || {};
  const ordersByDate = entry.orders || {};
  const selectedDates = entry.selectedDates || Object.keys(ordersByDate);

  const orders = [];
  selectedDates.forEach(function(date) {
    const dayOrders = ordersByDate[date];
    if (!dayOrders) return;
    const meals = [];
    ["Breakfast", "Lunch", "Dinner"].forEach(function(meal) {
      const m = dayOrders[meal];
      if (!m || (Number(m.subtotal) || 0) <= 0) return;
      const itemsArr = Object.keys(m.items || {})
        .filter(function(k) { return Number(m.items[k]) > 0; })
        .map(function(k) { return { colKey: k, qty: Number(m.items[k]) }; });
      if (!itemsArr.length) return;

      const area    = String(m.area || profile.area || "").trim();
      const isPickup = area.toLowerCase().indexOf("pickup") !== -1;

      const buildAddr = function() {
        if (isPickup) return "Self Pickup (A 104, Shree laxmi vihar society)";
        const parts = [];
        if (m.wing)    parts.push("Wing " + m.wing);
        if (m.flat)    parts.push("Flat " + m.flat);
        if (m.floor)   parts.push(m.floor + " Floor");
        if (m.society) parts.push(m.society);
        if (area)      parts.push(area);
        return parts.join(", ");
      };

      meals.push({
        type:           meal,
        items:          itemsArr,
        notesKitchen:   m.notesKitchen  || m.notes || "",
        notesDelivery:  m.notesDelivery || "",
        subtotal:       Number(m.subtotal) || 0,
        address:        buildAddr(),
        area:           isPickup ? "Self Pickup" : area,
        wing:           isPickup ? "" : (m.wing || ""),
        flat:           isPickup ? "" : (m.flat || ""),
        floor:          isPickup ? "" : (m.floor || ""),
        society:        isPickup ? "" : (m.society || ""),
        delivery_point: m.delivery_point || "",
        maps:           isPickup ? "" : (m.maps || ""),
        landmark:       isPickup ? "" : (m.landmark || "")
      });
    });
    if (meals.length) orders.push({ date: date, meals: meals });
  });

  // Match the structure hdfc_submitVerifiedOrder builds on the frontend
  const isSplit       = String(entry.payment_choice || "") === "Split";
  const walletApplied = isSplit ? Number(entry.wallet_applied || 0) : 0;

  return {
    profile: {
      name:               profile.name    || "Customer",
      phone:              entry.phone     || profile.phone || "",
      address:            (function() {
        const parts = [];
        if (profile.wing)    parts.push("Wing " + profile.wing);
        if (profile.flat)    parts.push("Flat " + profile.flat);
        if (profile.floor)   parts.push(profile.floor + " Floor");
        if (profile.society) parts.push(profile.society);
        if (profile.area)    parts.push(profile.area);
        return parts.join(", ");
      })(),
      wing:               profile.wing    || "",
      flat:               profile.flat    || "",
      floor:              profile.floor   || "",
      society:            profile.society || "",
      area:               profile.area    || "",
      maps:               profile.maps    || "",
      landmark:           profile.landmark|| "",
      meal_addresses:     JSON.stringify(entry.mealAddrs || {}),
      payment_preference: profile.payment_preference || "Daily Payment",
      isFirstTime:        !!entry.isFirstTime
    },
    orders:           orders,
    payment_method:   isSplit ? "Split" : "Gateway (HDFC)",
    payment_status:   "Paid",
    wallet_credit:    walletApplied,
    gateway_order_id: orderId,
    gateway_status:   statusCheck.status || "CHARGED",
    gateway_paid:     true,
    settle_all:       false,
    // Tag the source so logs make it clear this row came from the reconciler
    placed_via:       "reconciler"
  };
}

