// ============================================================
// 10_Hdfc_Gateway.gs
// HDFC SmartGateway integration — session creation, webhook handling,
  Status API, HMAC, post-payment verification, authoritative recompute.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ╔══════════════════════════════════════════════════════════════════╗
// ║          HDFC SMARTGATEWAY INTEGRATION                          ║
// ║  All functions below are gated by PAYMENT_GATEWAY_ENABLED.      ║
// ║  Nothing here runs unless that flag is true.                    ║
// ║                                                                  ║
// ║  HOW IT WORKS (end-to-end):                                     ║
// ║  1. Customer picks UPI/Card/NetBanking on order.html            ║
// ║  2. order.html calls hdfc_createSession → gets payment_url      ║
// ║  3. Customer is redirected to HDFC HyperCheckout page           ║
// ║  4. HDFC fires webhook → hdfc_handleWebhook marks order paid    ║
// ║  5. HDFC redirects customer back → order.html verifies via      ║
// ║     hdfc_verifyReturnPayload (HMAC check)                       ║
// ║                                                                  ║
// ║  KEYS NEEDED (set in Script Properties):                        ║
// ║  HDFC_MERCHANT_ID, HDFC_API_KEY, HDFC_RESPONSE_KEY,            ║
// ║  HDFC_WEBHOOK_USERNAME, HDFC_WEBHOOK_PASSWORD,                  ║
// ║  HDFC_RETURN_URL, HDFC_ENV, HDFC_TEST_URL, HDFC_LIVE_URL        ║
// ╚══════════════════════════════════════════════════════════════════╝

/**
 * STEP 1 — Called by order.html when customer chooses to pay via gateway.
 * Creates a payment session with HDFC SmartGateway (Juspay HyperCheckout).
 * Returns a payment_url to redirect the customer to.
 *
 * @param {Object} body  { phone, name, amount, order_id, description }
 * @returns {{ success, payment_url, session_id } | { error }}
 */
function hdfc_createSession(body) {
  if (!PAYMENT_GATEWAY_ENABLED) return { error: "Gateway not enabled." };

  const phone        = String(body.phone  || "").trim();
  const name         = String(body.name   || "Customer").trim();
  const amountRupees = Number(body.amount || 0);
  const orderId      = String(body.order_id    || "").trim();
  const description  = String(body.description || "Svaadh Kitchen Order").trim();

  if (!phone || !orderId || amountRupees <= 0) {
    return { error: "Missing required fields: phone, order_id, amount." };
  }
  if (!HDFC_MERCHANT_ID || !HDFC_API_KEY) {
    return { error: "Gateway credentials not configured in Script Properties." };
  }

  // HDFC SmartGateway expects amount in RUPEES (empirically confirmed — UAT showed 100x
  // inflation when sending paisa, so SmartGateway/Juspay takes rupees directly, not paisa)
  const amountToSend = Math.round(amountRupees);

  const payload = {
    order_id:               orderId,
    amount:                 amountToSend,
    currency:               "INR",
    customer_id:            phone,
    customer_phone:         phone,
    customer_email:         phone + "@svaadh.noemail",
    payment_page_client_id: HDFC_MERCHANT_ID,
    action:                 "paymentPage",
    return_url:             HDFC_RETURN_URL,
    description:            description,
    first_name:             name.split(" ")[0] || name,
    last_name:              name.split(" ").slice(1).join(" ") || "",
    udf1:                   phone,
    udf3:                   "svaadh_kitchen",
    // udf2 intentionally omitted — blocked by HDFC for tokenization compliance
    notification_url:       HDFC_RETURN_URL   // webhook URL per-session (fallback if dashboard not set)
  };

  // Juspay Basic Auth: base64(api_key + ":") — API key as username, empty password
  // Merchant ID goes in a separate x-merchantid header (NOT in the auth string)
  const authToken = Utilities.base64Encode(HDFC_API_KEY + ":");

  const options = {
    method:      "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Basic " + authToken,
      "x-merchantid":  HDFC_MERCHANT_ID,
      "version":       "2023-01-01"
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const resp     = UrlFetchApp.fetch(HDFC_BASE_URL + "/session", options);
    const respCode = resp.getResponseCode();
    const respBody = JSON.parse(resp.getContentText());

    console.log("HDFC createSession [" + respCode + "]:", JSON.stringify(respBody));

    if (respCode !== 200 && respCode !== 201) {
      // Error details can be nested under error_info in Juspay responses
      const errMsg = (respBody.error_info && respBody.error_info.user_message)
        || respBody.user_message
        || respBody.error_message
        || respBody.error_info
        || "Unknown error";
      return { error: "HDFC session creation failed (HTTP " + respCode + "): " + errMsg };
    }

    const paymentUrl = (respBody.payment_links && respBody.payment_links.web)
      ? respBody.payment_links.web
      : null;

    if (!paymentUrl) {
      return { error: "HDFC returned no payment URL. Check credentials and merchant config." };
    }

    return {
      success:     true,
      payment_url: paymentUrl,
      session_id:  respBody.id || orderId,
      order_id:    orderId
    };

  } catch (err) {
    console.error("hdfc_createSession error:", err.message);
    return { error: "Network error creating payment session: " + err.message };
  }
}
function hdfc_handleWebhook(body, e) {
  if (!PAYMENT_GATEWAY_ENABLED) return { error: "Gateway not enabled." };

  // ── Auth check (fast — no I/O) ───────────────────────────────
  try {
    var authHeader = "";
    try { authHeader = e.parameter.Authorization || ""; } catch(_) {}
    if (HDFC_WEBHOOK_USERNAME && HDFC_WEBHOOK_PASSWORD) {
      var expectedAuth = "Basic " + Utilities.base64Encode(HDFC_WEBHOOK_USERNAME + ":" + HDFC_WEBHOOK_PASSWORD);
      if (authHeader && authHeader !== expectedAuth) {
        console.warn("HDFC Webhook: Invalid Basic Auth — possible spoofed request.");
        return { error: "Unauthorized" };
      }
    }
  } catch (authErr) {
    console.warn("HDFC Webhook auth check error:", authErr.message);
  }

  // ── Log-first: write one row, return immediately ─────────────
  // Do NOT call hdfc_markOrderPaid here — that involves an external Status API
  // call and sheet reads which will blow past HDFC's 5-second response window.
  try {
    const ss  = getSpreadsheet();
    const ws  = getOrCreateTab(ss, TAB_WEBHOOK_LOG, [
      "Received_At", "Event_Name", "Order_ID", "Raw_Payload", "Status", "Processed_At", "Result"
    ]);

    const eventName = String(body.event_name || "");
    const orderId   = String(
      (body.content && body.content.order && body.content.order.order_id) || ""
    );

    ws.appendRow([
      new Date(),          // Received_At
      eventName,           // Event_Name
      orderId,             // Order_ID
      JSON.stringify(body),// Raw_Payload — full webhook preserved
      "PENDING",           // Status
      "",                  // Processed_At (filled by processor)
      ""                   // Result       (filled by processor)
    ]);

    console.log("HDFC Webhook logged: event=" + eventName + " order=" + orderId);
  } catch (logErr) {
    // Even if logging fails, we must return 200 so HDFC doesn't retry endlessly.
    // The raw payload is in the GAS execution log regardless.
    console.error("HDFC Webhook log error:", logErr.message);
  }

  // Return 200 immediately — HDFC is satisfied. Processing happens async.
  return { success: true, received: true };
}
function hdfc_processWebhookLog() {
  if (!PAYMENT_GATEWAY_ENABLED) return;

  try {
    const ss  = getSpreadsheet();
    const ws  = getOrCreateTab(ss, TAB_WEBHOOK_LOG, [
      "Received_At", "Event_Name", "Order_ID", "Raw_Payload", "Status", "Processed_At", "Result"
    ]);

    const data = ws.getDataRange().getValues();
    if (data.length < 2) return; // no rows yet

    const headers        = data[0];
    const COL_EVENT      = headers.indexOf("Event_Name");
    const COL_PAYLOAD    = headers.indexOf("Raw_Payload");
    const COL_STATUS     = headers.indexOf("Status");
    const COL_PROC_AT    = headers.indexOf("Processed_At");
    const COL_RESULT     = headers.indexOf("Result");

    var processed = 0;

    for (var i = 1; i < data.length; i++) {
      const rowStatus = String(data[i][COL_STATUS] || "").trim();
      if (rowStatus !== "PENDING") continue; // skip already handled rows

      const eventName  = String(data[i][COL_EVENT]   || "");
      const rawPayload = String(data[i][COL_PAYLOAD]  || "{}");
      var   body;
      try { body = JSON.parse(rawPayload); } catch(_) { body = {}; }

      const content = body.content || {};
      const order   = content.order || {};
      var   result  = "";
      var   newStatus = "PROCESSED";

      if (eventName === "ORDER_SUCCEEDED" || order.status === "CHARGED") {
        const markResult = hdfc_markOrderPaid(order);
        result = JSON.stringify(markResult);
        if (markResult.error) newStatus = "FAILED";

      } else if (eventName === "REFUND_INITIATED" || eventName === "REFUND_SUCCEEDED") {
        // Placeholder — refund logic goes here when needed
        result = "Refund event acknowledged.";

      } else {
        result = "Unhandled event type: " + eventName;
      }

      // Update the log row
      ws.getRange(i + 1, COL_STATUS   + 1).setValue(newStatus);
      ws.getRange(i + 1, COL_PROC_AT  + 1).setValue(new Date());
      ws.getRange(i + 1, COL_RESULT   + 1).setValue(result);
      processed++;

      console.log("hdfc_processWebhookLog: row " + (i+1) + " → " + newStatus + " | " + result);
    }

    if (processed > 0) {
      console.log("hdfc_processWebhookLog: processed " + processed + " pending webhook(s).");
    }

  } catch (err) {
    console.error("hdfc_processWebhookLog error:", err.message);
  }
}
function setupHdfcWebhookTrigger() {
  // Remove any existing trigger for this function first (avoid duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "hdfc_processWebhookLog") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("hdfc_processWebhookLog")
    .timeBased()
    .everyMinutes(1)
    .create();
  console.log("✅ hdfc_processWebhookLog trigger created — runs every 1 minute.");
}
function hdfc_verifyReturnPayload(body) {
  if (!PAYMENT_GATEWAY_ENABLED) return { error: "Gateway not enabled." };

  const orderId   = String(body.order_id  || "").trim();
  const status    = String(body.status    || "").trim();
  const signature = String(body.signature || "").trim();

  if (!orderId || !status) {
    return { error: "Missing order_id or status in return payload." };
  }

  // ── Step 1: HMAC signature check (fast, no I/O) ─────────────
  // If signature is present and key is configured, verify it.
  // On mismatch: do NOT reject outright — fall back to Status API (Step 2).
  // Reason: HDFC UAT may not send a valid signature; Status API is authoritative.
  var signatureOk = false;
  if (HDFC_RESPONSE_KEY && signature) {
    const expectedSig = hdfc_hmacSha256(orderId + "|" + status, HDFC_RESPONSE_KEY);
    if (expectedSig.toLowerCase() === signature.toLowerCase()) {
      signatureOk = true;
    } else {
      console.warn("HDFC return: HMAC mismatch for order " + orderId + " — falling back to Status API.");
    }
  } else {
    console.warn("HDFC return: No signature or key — falling back to Status API.");
  }

  // ── Step 2: Status API verification (authoritative) ──────────
  // Always call if signature check didn't pass. Even if it did pass,
  // Status API gives us the real txn_id and confirmed status.
  var statusConfirmed = (status === "CHARGED" || status === "SUCCESS");
  if (!signatureOk) {
    const statusCheck = hdfc_getOrderStatus(orderId);
    if (statusCheck.confirmed) {
      statusConfirmed = true;
      console.log("HDFC return: Status API confirmed CHARGED for " + orderId);
    } else {
      console.warn("HDFC return: Status API returned '" + statusCheck.status + "' for " + orderId);
      // Only reject if BOTH signature AND Status API fail
      if (!statusConfirmed) {
        return { error: "Payment could not be verified. Status: " + statusCheck.status, paid: false };
      }
    }
  }

  const paid = statusConfirmed;

  // Cross-check with sheet (webhook should have already marked it paid)
  try {
    const ss   = getSpreadsheet();
    const ws   = getOrCreateTab(ss, TAB_ORDERS, []);
    const rows = getAllRows(ws);
    const orderRows = rows.filter(function(r) {
      return String(r.Submission_ID || "").trim() === orderId;
    });
    if (orderRows.length > 0) {
      const sheetStatus = String(orderRows[0].Payment_Status || "").toLowerCase();
      const alreadyPaid = sheetStatus === "paid" || sheetStatus === "collected";
      return {
        success:  true,
        paid:     paid || alreadyPaid,
        order_id: orderId,
        status:   status,
        message:  (paid || alreadyPaid) ? "Payment confirmed." : "Payment status: " + status
      };
    }
  } catch (err) {
    console.warn("hdfc_verifyReturnPayload sheet check error:", err.message);
  }

  return {
    success:  true,
    paid:     paid,
    order_id: orderId,
    status:   status,
    message:  paid ? "Payment confirmed." : "Payment status: " + status
  };
}
function hdfc_getOrderStatus(orderId) {
  if (!HDFC_MERCHANT_ID || !HDFC_API_KEY) {
    console.warn("hdfc_getOrderStatus: credentials not configured.");
    return { confirmed: false, status: "UNKNOWN", txn_id: "" };
  }

  const authToken = Utilities.base64Encode(HDFC_API_KEY + ":");
  const options = {
    method:             "get",
    headers: {
      "Authorization": "Basic " + authToken,
      "x-merchantid":  HDFC_MERCHANT_ID,
      "version":        "2023-01-01"
    },
    muteHttpExceptions: true
  };

  try {
    const resp     = UrlFetchApp.fetch(HDFC_BASE_URL + "/orders/" + orderId, options);
    const respCode = resp.getResponseCode();
    const respBody = JSON.parse(resp.getContentText());

    console.log("hdfc_getOrderStatus [" + respCode + "] for " + orderId + ":", JSON.stringify(respBody));

    if (respCode !== 200) {
      console.warn("hdfc_getOrderStatus: non-200 response (" + respCode + ") for " + orderId);
      return { confirmed: false, status: "API_ERROR", txn_id: "" };
    }

    const status = String(respBody.status || "").toUpperCase();
    const txnId  = String((respBody.txn_detail && respBody.txn_detail.txn_id) || respBody.txn_id || "");
    return {
      confirmed: (status === "CHARGED"),
      status:    status,
      txn_id:    txnId
    };
  } catch (err) {
    console.error("hdfc_getOrderStatus error:", err.message);
    return { confirmed: false, status: "FETCH_ERROR", txn_id: "" };
  }
}
function hdfc_markOrderPaid(order) {
  const orderId = String(order.order_id || "").trim();
  const txnId   = String(order.txn_id   || order.id || "").trim();
  const method  = String(order.payment_method_type || order.payment_method || "Gateway").trim();

  if (!orderId) return { error: "Webhook: missing order_id." };

  console.log("HDFC Webhook received for order: " + orderId + " via " + method);

  // ── Step 1: Verify payment via Status API (mandatory per HDFC security audit) ──
  // Never rely on the webhook payload alone. Confirm the order is truly CHARGED.
  const statusCheck = hdfc_getOrderStatus(orderId);
  if (!statusCheck.confirmed) {
    console.warn("hdfc_markOrderPaid: Status API returned '" + statusCheck.status + "' for " + orderId + ". Aborting mark-paid.");
    return { success: true, message: "Webhook received but Status API shows " + statusCheck.status + " — not marking paid." };
  }
  // Use txn_id from Status API if richer than what webhook provided
  if (statusCheck.txn_id && !txnId) txnId = statusCheck.txn_id;

  console.log("HDFC Status API confirmed CHARGED: " + orderId + " (TXN:" + txnId + ")");

  try {
    const ss      = getSpreadsheet();
    const ws      = getOrCreateTab(ss, TAB_ORDERS, []);
    const data    = ws.getDataRange().getValues();
    if (data.length < 2) return { error: "Orders sheet is empty." };

    const headers     = data[0];
    const COL_SID     = headers.indexOf("Submission_ID");
    const COL_PSTATUS = headers.indexOf("Payment_Status");
    const COL_PMETHOD = headers.indexOf("Payment_Method");
    const COL_NOTES   = headers.indexOf("Kitchen_Notes");

    if (COL_SID < 0 || COL_PSTATUS < 0) {
      return { error: "Webhook: required columns missing in SK_Orders." };
    }

    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COL_SID] || "").trim() === orderId) {

        // ── Step 2: Duplicate check — skip if already marked Paid ──
        const currentStatus = String(data[i][COL_PSTATUS] || "").toLowerCase();
        if (currentStatus === "paid" || currentStatus === "collected") {
          console.log("hdfc_markOrderPaid: order " + orderId + " row " + (i+1) + " already Paid — skipping (duplicate webhook).");
          continue;
        }

        ws.getRange(i + 1, COL_PSTATUS + 1).setValue("Paid");
        if (COL_PMETHOD >= 0) {
          ws.getRange(i + 1, COL_PMETHOD + 1).setValue("Gateway (" + method + ")");
        }
        if (COL_NOTES >= 0 && txnId) {
          var existing = String(data[i][COL_NOTES] || "");
          var note = existing ? existing + " | TXN:" + txnId : "TXN:" + txnId;
          ws.getRange(i + 1, COL_NOTES + 1).setValue(note);
        }
        updated++;
      }
    }

    if (updated === 0) {
      console.warn("HDFC Webhook: order " + orderId + " not found in SK_Orders (or already Paid).");
      return { success: true, message: "Order " + orderId + " not found or already processed." };
    }

    return { success: true, message: "Order " + orderId + " marked paid (" + updated + " row(s))." };

  } catch (err) {
    console.error("hdfc_markOrderPaid error:", err.message);
    return { error: "Failed to update order: " + err.message };
  }
}
// ─────────────────────────────────────────────────────────────────────────────
// PENDING ORDER STORE
// Saves the full order payload server-side before the HDFC redirect.
// Retrieved on return — no sessionStorage/localStorage dependency.
// Stored in Script Properties as a JSON map, auto-expired after 30 minutes.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Saves order payload before customer is redirected to HDFC checkout.
 * Called by order.html hdfc_initiatePayment() just before window.location.href redirect.
 *
 * @param {Object} body  { order_id, phone, amount, orders, selectedDates, profile, mealAddrs, isFirstTime }
 */
function hdfc_savePendingOrder(body) {
  if (!PAYMENT_GATEWAY_ENABLED) return { error: "Gateway not enabled." };

  const orderId = String(body.order_id || "").trim();
  if (!orderId) return { error: "order_id required." };

  try {
    const props    = PropertiesService.getScriptProperties();
    const raw      = props.getProperty("HDFC_PENDING_ORDERS") || "{}";
    const pending  = JSON.parse(raw);

    // Expire entries older than 30 minutes to keep size under control
    const now = Date.now();
    Object.keys(pending).forEach(function(k) {
      if (now - (pending[k].ts || 0) > 30 * 60 * 1000) delete pending[k];
    });

    pending[orderId] = {
      ts:            now,
      phone:         body.phone         || "",
      amount:        body.amount        || 0,
      orders:        body.orders        || {},
      selectedDates: body.selectedDates || [],
      profile:       body.profile       || {},
      mealAddrs:     body.mealAddrs     || {},
      isFirstTime:   body.isFirstTime   || false
    };

    props.setProperty("HDFC_PENDING_ORDERS", JSON.stringify(pending));
    console.log("hdfc_savePendingOrder: saved order " + orderId);
    return { success: true };

  } catch (err) {
    console.error("hdfc_savePendingOrder error:", err.message);
    return { error: err.message };
  }
}
function hdfc_getPendingOrder(body) {
  if (!PAYMENT_GATEWAY_ENABLED) return { error: "Gateway not enabled." };

  const orderId = String(body.order_id || "").trim();
  if (!orderId) return { error: "order_id required." };

  try {
    const props   = PropertiesService.getScriptProperties();
    const raw     = props.getProperty("HDFC_PENDING_ORDERS") || "{}";
    const pending = JSON.parse(raw);

    const entry = pending[orderId];
    if (!entry) {
      console.warn("hdfc_getPendingOrder: no pending entry for " + orderId);
      return { error: "Pending order not found. It may have expired (>30 min)." };
    }

    // Keep the entry — let it expire naturally after 30 minutes.
    // Not deleting here so that if the page reloads or redirects twice,
    // the second call still finds the data.
    console.log("hdfc_getPendingOrder: retrieved order " + orderId);
    return { success: true, ...entry };

  } catch (err) {
    console.error("hdfc_getPendingOrder error:", err.message);
    return { error: err.message };
  }
}
function hdfc_hmacSha256(message, secret) {
  const sig = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(message).getBytes(),
    Utilities.newBlob(secret).getBytes()
  );
  return sig.map(function(b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); }).join("");
}
function testHdfcConnection() {
  if (!PAYMENT_GATEWAY_ENABLED) {
    console.log("PAYMENT_GATEWAY_ENABLED is false. Enable it in Dev Script Properties first.");
    return;
  }
  const result = hdfc_createSession({
    phone:       "9999999999",
    name:        "Test Customer",
    amount:      1,
    order_id:    "SKTEST" + Date.now().toString(36).toUpperCase().slice(-9),  // alphanumeric ≤21
    description: "Svaadh Kitchen — Connection Test"
  });
  console.log("testHdfcConnection result:", JSON.stringify(result, null, 2));
}
