// ============================================================
// 13_Maintenance.gs
// Keep-alive pings, AI chatbot, system maintenance helpers.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── CHATBOT ──────────────────────────────────────────────────

function handleChat(body) {
  const userMessage = String(body.message || "").trim();
  const history     = body.history || [];   // [{role:"user"|"model", text:"..."}]
  if (!userMessage) return {reply: "Please send a message."};

  let extraMenu = "";
  try {
    // Basic date detection (e.g., "tomorrow", "15th", "15-04", "April 15")
    const msgLower = userMessage.toLowerCase();
    let targetDate = new Date();
    let foundDate = false;

    if (msgLower.includes("tomorrow")) {
      targetDate.setDate(targetDate.getDate() + 1);
      foundDate = true;
    } else if (msgLower.includes("today")) {
      foundDate = true;
    } else {
      // Look for day numbers (1st, 2nd, 3rd, 4th... 31st) or simple digits
      const dayMatch = msgLower.match(/(\d{1,2})(st|nd|rd|th)?/);
      if (dayMatch) {
        const day = parseInt(dayMatch[1]);
        if (day >= 1 && day <= 31) {
          targetDate.setDate(day);
          // If the detected day is in the past, assume next month
          if (targetDate < new Date()) targetDate.setMonth(targetDate.getMonth() + 1);
          foundDate = true;
        }
      }
    }

    if (foundDate) {
      const dateStr = Utilities.formatDate(targetDate, "Asia/Kolkata", "yyyy-MM-dd");
      const m = getMenu(dateStr);
      const bf = (m.breakfast || []).map(function(x) { return x.name + " ₹" + x.price; }).join(", ");
      extraMenu = "\nMenu for " + dateStr + " (" + Utilities.formatDate(targetDate, "Asia/Kolkata", "EEEE") + "): "
        + (Utilities.formatDate(targetDate, "Asia/Kolkata", "EEEE") === "Sunday" ? "CLOSED (Sunday)" :
          "BF: " + (bf || "TBD") + " | L: " + (m.lunch_dry || "") + (m.lunch_curry ? " & " + m.lunch_curry : "") +
          " | D: " + (m.dinner_dry || "") + (m.dinner_curry ? " & " + m.dinner_curry : ""));
    }
  } catch (e) {
    console.error("Date menu fetch failed:", e);
  }

  return {reply: callGemini(buildSystemPrompt(extraMenu), history, userMessage)};
}
function buildSystemPrompt(extraMenu) {
  const B = BUSINESS_CONTEXT;
  const breads = B.menu.breads.map(function(i){ return i.name+"₹"+i.price; }).join(", ");
  const sabji  = B.menu.sabji.map(function(i){ return i.name+"₹"+i.price; }).join(", ");
  const basics = B.menu.basics.map(function(i){ return i.name+"₹"+i.price; }).join(", ");

  var todayLine = "";
  try {
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
    var dayName = Utilities.formatDate(now, "Asia/Kolkata", "EEEE");
    var m = getMenu(todayStr);
    var bf = (m.breakfast||[]).map(function(x){ return x.name+"₹"+x.price; }).join(", ");
    todayLine = "Today is "+dayName+", "+todayStr+". "
      +(dayName==="Sunday" ? "Kitchen is CLOSED today (Sunday).\n" :
        "BF:"+(bf||"TBD")
      +"|L:"+(m.lunch_dry||"")+(m.lunch_curry?" & "+m.lunch_curry:"")
      +"|D:"+(m.dinner_dry||"")+(m.dinner_curry?" & "+m.dinner_curry:"")
      +((!m.lunch_dry&&!m.dinner_dry)?" (sabji TBD—send to WA group)":"")+"\n");
  } catch(e) { todayLine = "Today's menu: check WhatsApp group.\n"; }

  const prompt = "You are a helpful assistant for Svaadh Kitchen, a vegetarian cloud kitchen in Hadapsar, Pune."
    +" Closed Sundays. Over 2.5 years of service (since Aug 2023). Cutoffs: BF<7AM, Lunch<9AM, Dinner<4:30PM."
    +" AREAS: " + B.locations_served.join(", ") + ".\n"
    +" DELIVERY POLICY: FREE for Bhosale Nagar, Triveni Nagar, and Self Pickup. Other areas ₹10/meal if subtotal < ₹100. "
    + B.delivery.outside_policy + "\n"
    +" PRIVACY & SECURITY: DO NOT disclose user phone numbers, PINs, transaction IDs, UPI details, or specific refund info. If a user asks about their payment or refund, tell them to check their 'Svaadh Wallet' or 'View/Edit existing orders' dashboard, or message us on WhatsApp at " + B.contact.whatsapp + ".\n"
    + todayLine + (extraMenu || "")
    +"\nMEAL MODEL: Make Your Own Meal (not a fixed thali). Customers pick items individually.\n"
    +"Lunch/Dinner — Breads:"+breads+" | Sabji:"+sabji+" | Basics:"+basics+"\n"
    +"Breakfast: daily rotating ₹35–₹70. "+B.menu.breakfast_note+"\n"
    +"Self pickup also available (no delivery charge).\n"
    +"Uses Pure Ghee & Groundnut refined oil. Pure Veg kitchen.\n"
    +"Discounts(auto): 5% off≥₹300/day, 7.5% off≥₹450/day.\n"
    +"Payment: Wallet (Prepaid) or UPI("+B.payment.upi_id+"), prepaid cycle (requires wallet balance).\n"
    +"Order: "+B.ordering.order_url+" — no login needed, phone=identity, can book multiple days.\n"
    +"WhatsApp: "+B.contact.whatsapp+" | WA group: "+B.contact.whatsapp_group+"\n"
    +"Reply in customer's language (English/Hindi/Marathi). Be brief & warm. Match the language they use."
    +" For orders, always send to order URL. Don't invent info. Direct unknowns to WhatsApp.";
  
  return prompt;
}
function callGemini(systemPrompt, history, userMessage) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return "I'm having trouble connecting right now. Please WhatsApp us at +91 99307 48908 for help!";
  }

  // Build contents array: last 6 history messages + current message (caps token usage)
  const contents = [];
  const recentHistory = (history || []).slice(-6);
  recentHistory.forEach(function(msg) {
    if (msg.role === "user" || msg.role === "model") {
      contents.push({role: msg.role, parts: [{text: String(msg.text || "")}]});
    }
  });
  contents.push({role: "user", parts: [{text: userMessage}]});

  const payload = {
    system_instruction: {parts: [{text: systemPrompt}]},
    contents: contents,
    generationConfig: {maxOutputTokens: 512, temperature: 0.7}
  };

  try {
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const raw = response.getContentText();
    const data = JSON.parse(raw);
    if (data.candidates && data.candidates[0] && data.candidates[0].content) {
      return data.candidates[0].content.parts[0].text;
    }
    return "I'm not sure how to answer that. Please WhatsApp us at +91 99307 48908!";
  } catch(e) {
    return "I'm having trouble right now. Please call or WhatsApp us at +91 99307 48908.";
  }
}
// ── KEEP-ALIVE ────────────────────────────────────────────────────────────────
// Keeps the GAS instance warm so customers never hit a cold-start timeout.
// Set up once: Apps Script editor → Triggers → Add Trigger:
//   Function: keepAlive | Event: Time-based | Type: Minutes timer | Every: 10 minutes
function keepAlive() {
  // Intentionally empty — just waking the instance is enough.
  // GAS logs will show "keepAlive" executions confirming it's running.
}
// Run this once from Apps Script editor to register the trigger automatically.
// After that it runs forever — no manual intervention needed.
function setupKeepAliveTrigger() {
  // Remove any existing keepAlive trigger first (avoid duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "keepAlive") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("keepAlive")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("keepAlive trigger registered — fires every 10 minutes.");
}

// ── QUARTERLY ARCHIVE ─────────────────────────────────────────────────────────
/*
  Archives SK_Orders and SK_Wallet for a given month into a new Google
  Spreadsheet, writes Balance Carry Forward snapshots so wallet balances are
  preserved, then deletes the archived rows from the main sheet.

  Runs on the 10th of every month — archives the previous calendar month.
  e.g. May 10 → archives April data.
