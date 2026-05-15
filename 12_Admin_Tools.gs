// ============================================================
// 12_Admin_Tools.gs
// Admin remediation tools (voidOrderRow), seed data, ledger automation.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

// ── CUSTOMER LEDGER ──────────────────────────────────────────
function _getLedgerFolder(year) {
  const parentName = LEDGER_FOLDER;
  const yearStr    = String(year);
  const parents    = DriveApp.getFoldersByName(parentName);
  let parent;
  if (parents.hasNext()) { parent = parents.next(); }
  else { parent = DriveApp.createFolder(parentName); }

  const children = parent.getFoldersByName(yearStr);
  if (children.hasNext()) return children.next();
  return parent.createFolder(yearStr);
}
function _getOrCreateCustomerLedger(ss, phone, name, year) {
  // Check if ledger ID already stored
  const custWs  = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const custRows = getAllRows(custWs);
  const cust = custRows.find(r => String(r.Phone).trim() === String(phone).trim());

  if (cust && cust.Ledger_Sheet_ID) {
    try {
      return SpreadsheetApp.openById(cust.Ledger_Sheet_ID);
    } catch(e) { /* file may have been deleted */ }
  }

  const folder  = _getLedgerFolder(year);
  const ledger  = SpreadsheetApp.create(`Svaadh — ${name} (${phone})`);
  DriveApp.getFileById(ledger.getId()).moveTo(folder);

  // Store ID in customers sheet
  if (cust) {
    const hIdx = headerIndex(custWs);
    if (hIdx["Ledger_Sheet_ID"]) {
      custWs.getRange(cust._row, hIdx["Ledger_Sheet_ID"]).setValue(ledger.getId());
    }
  }
  return ledger;
}
function _ensureMonthTab(ledgerSs, year, monthIdx) {
  const MONTHS = ["January","February","March","April","May","June",
                  "July","August","September","October","November","December"];
  const tabName = `${MONTHS[monthIdx]} ${year}`;
  let ws = ledgerSs.getSheetByName(tabName);
  if (ws) return ws;

  ws = ledgerSs.insertSheet(tabName);
  const headers = ["Submission_ID","Date","Meal","Items Ordered","Subtotal (₹)","Delivery (₹)","Discount (₹)","Net Total (₹)"];
  const periods = ["1–10","11–20","21–end"];

  let row = 1;
  periods.forEach(p => {
    ws.getRange(row, 1).setValue(`● Period ${p}`).setFontWeight("bold").setBackground("#f5f0eb");
    ws.getRange(row, 1, 1, headers.length).merge();
    row++;
    ws.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#c0392b").setFontColor("white");
    row++;
    // 3 blank data rows (grow dynamically)
    row += 3;
    ws.getRange(row, 1).setValue("Period Total").setFontWeight("bold");
    ws.getRange(row, 7).setFormula(`=SUMIF(G${row-3}:G${row-1},"<>",G${row-3}:G${row-1})`);
    row += 2;
  });
  return ws;
}
function _updateLedger(ss, profile, orders) {
  const ist  = new Date(new Date().getTime() + 5.5 * 3600 * 1000);
  const year = ist.getFullYear();
  const monthIdx = ist.getMonth();

  const ledger = _getOrCreateCustomerLedger(ss, profile.phone, profile.name, year);
  const ws     = _ensureMonthTab(ledger, year, monthIdx);

  for (const order of orders) {
    for (const meal of order.meals) {
      const sid = meal._sid || order.submissionId || "";
      const summary = (meal.items || [])
        .filter(function(it){ return it.qty > 0; })
        .map(function(it){ return it.qty + "×" + it.colKey; })
        .join(", ") || "—";
      const delCharge = meal.deliveryCharge || 0;
      const discAmt   = meal.discountAmount  || 0;
      const netTotal  = meal.subtotal + delCharge - discAmt;
      ws.appendRow([sid, order.date, meal.type, summary, meal.subtotal, delCharge, discAmt, netTotal]);
    }
  }
}
function getOrCreateFolderPath(pathParts) {
  var folder = DriveApp.getRootFolder();
  pathParts.forEach(function(name) {
    var iter = folder.getFoldersByName(name);
    folder = iter.hasNext() ? iter.next() : folder.createFolder(name);
  });
  return folder;
}
function seedTestData() {
  try {
    Logger.log("Starting seedTestData...");
    const ss = getSpreadsheet();
    if (!ss) throw new Error("Could not open spreadsheet. Check SHEET_ID in Script Properties.");
    
    const ws = ss.getSheetByName(TAB_ORDERS);
    if (!ws) throw new Error("Sheet tab '" + TAB_ORDERS + "' not found.");

    const today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate() + 1);

    const getDayStr = (d) => Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
    const tStr = getDayStr(today);
    const mStr = getDayStr(tomorrow);

    Logger.log("Generating data for " + tStr + " and " + mStr);

    const testData = [
      { date: tStr, meal: "Breakfast", name: "Rahul Deshpande", area: "Magarpatta", society: "Pentagon 1", wing: "B", flat: "402", items: {"Kanda Poha": 2}, total: 70, notes: "Less spicy please" },
      { date: tStr, meal: "Breakfast", name: "Anjali Singh", area: "Amanora", society: "Tower 13", wing: "C", flat: "1805", items: {"Ghee Upma": 1, "Thalipeeth": 1}, total: 90, notes: "Extra chutney" },
      { date: tStr, meal: "Lunch", name: "Amit Kulkarni", area: "Bhosale Nagar", society: "Laxmi Vihar", wing: "A", flat: "104", items: {"Chapati": 3, "Dry_Sabji_Mini": 1, "Dal": 1}, total: 71, notes: "Deliver at gate" },
      { date: tStr, meal: "Lunch", name: "Sneha Patil", area: "Magarpatta", society: "Cosmos", wing: "E", flat: "P-5", items: {"Phulka": 2, "Curry_Sabji_Full": 1, "Rice": 1}, total: 77, notes: "" },
      { date: tStr, meal: "Dinner", name: "Mayur Joshi", area: "DP Road", society: "Riverview", wing: "F", flat: "901", items: {"Jowar_Bhakri": 2, "Curry_Sabji_Mini": 1}, total: 62, notes: "Ring bell and leave" },
      { date: mStr, meal: "Breakfast", name: "Priya Rao", area: "Magarpatta", society: "Pentagon 3", wing: "A", flat: "610", items: {"Sabudana Khichdi": 1}, total: 40, notes: "" },
      { date: mStr, meal: "Lunch", name: "Vikram Shah", area: "Amanora", society: "Adreno", wing: "1", flat: "1502", items: {"Ghee_Phulka": 4, "Dry_Sabji_Full": 1, "Salad": 1}, total: 100, notes: "Call on arrival" },
      { date: mStr, meal: "Dinner", name: "Svaadh Test", area: "Bhosale Nagar", society: "Self Pickup", wing: "-", flat: "-", items: {"Chapati": 2, "Dry_Sabji_Mini": 1, "Dal": 1}, total: 62, notes: "I will pick up" }
    ];

    testData.forEach((d, idx) => {
      Logger.log("Processing row " + (idx + 1) + ": " + d.name);
      const row = new Array(ORDERS_HEADERS.length).fill(""); 
      
      row[0] = "TEST-" + Math.floor(Math.random() * 100000); 
      row[1] = new Date(); 
      row[2] = d.date; 
      row[3] = d.meal; 
      row[4] = d.name; 
      row[5] = "9999999999"; 
      row[6] = d.area; 
      row[7] = d.wing; 
      row[8] = d.flat; 
      row[9] = "1"; 
      row[10] = d.society; 
      row[11] = d.wing + "-" + d.flat + ", " + d.society; 
      row[14] = JSON.stringify(d.items); 
      
      Object.keys(d.items).forEach(itemName => {
         const colIdx = ORDERS_HEADERS.indexOf(itemName); 
         if (colIdx >= 0) row[colIdx] = d.items[itemName];
         else {
           const colKey = itemName.replace(/ /g,"_");
           const altIdx = ORDERS_HEADERS.indexOf(colKey);
           if (altIdx >= 0) row[altIdx] = d.items[itemName];
         }
      });

      if (d.meal === "Breakfast") {
        let bIdx = 0;
        for (const [key, val] of Object.entries(d.items)) {
           if (bIdx === 0) { row[29] = key; row[30] = val; }
           if (bIdx === 1) { row[31] = key; row[32] = val; }
           if (bIdx === 2) { row[33] = key; row[34] = val; }
           if (bIdx === 3) { row[35] = key; row[36] = val; }
           bIdx++;
        }
      }

      const fieldMap = {
        "Special_Notes_Kitchen": d.notes,
        "Food_Subtotal": d.total,
        "Net_Total": d.total,
        "Payment_Method": "UPI",
        "Payment_Status": "Paid",
        "Payment_Freq": "Daily Payment"
      };

      Object.keys(fieldMap).forEach(key => {
        const i = ORDERS_HEADERS.indexOf(key);
        if (i >= 0) row[i] = fieldMap[key];
      });

      ws.appendRow(row);
    });

    Logger.log("Seed successful.");
    return "Success: 8 test orders added to sheet.";
  } catch(err) {
    Logger.log("ERROR in seedTestData: " + err.message);
    return "Error: " + err.message;
  }
}

// ============================================================
// voidOrderRow — soft-cancel an SK_Orders row
// ============================================================
// Stamps Payment_Status with a void marker plus reason+timestamp
// so the order is excluded from reports without being deleted.
// Used to clean up the SK-20260514-2601 false-Paid row during the
// HDFC UAT remediation (see CODE_VERSION 14.7 notes).
// ============================================================

// ── ADMIN: VOID AN SK_ORDERS ROW ─────────────────────────────────────────────
// Marks an order as void (e.g. duplicate, or marked Paid in error when
// HDFC actually says AUTHORIZATION_FAILED). Doesn't delete — keeps the row
// for audit. Sets Payment_Status to "Voided" and appends the reason to
// Special_Notes_Kitchen for visibility.
function voidOrderRow(submissionId, reason) {
  if (!submissionId) return { success: false, error: "submissionId required" };
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const data = ws.getDataRange().getValues();
  const headers = data[0] || [];
  const sidCol     = headers.indexOf("Submission_ID");
  const psCol      = headers.indexOf("Payment_Status");
  const notesCol   = headers.indexOf("Special_Notes_Kitchen");
  if (sidCol === -1) return { success: false, error: "Submission_ID column missing" };

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][sidCol] || "").trim() === String(submissionId).trim()) {
      const stamp  = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm");
      const note   = "[VOIDED " + stamp + "] " + (reason || "Admin void");
      const existingNote = String(data[r][notesCol] || "").trim();
      const newNote = existingNote ? existingNote + " | " + note : note;

      if (psCol !== -1)    ws.getRange(r + 1, psCol + 1).setValue("Voided");
      if (notesCol !== -1) ws.getRange(r + 1, notesCol + 1).setValue(newNote);
      SpreadsheetApp.flush();
      console.log("voidOrderRow: " + submissionId + " voided. Reason: " + reason);
      return { success: true, submissionId: submissionId, reason: reason };
    }
  }
  return { success: false, error: "Submission_ID not found: " + submissionId };
}

// ============================================================
// Bank reconciliation tools — ported from main (post-merge sync).
// Used by the admin dashboard to reconcile a bank-statement export
// against unpaid SK_Orders rows for a date range and bulk-mark them
// Paid. Routes wired in 01_Router.gs.
// ============================================================
function getUnpaidOrdersData(p) {
  const dateFrom = p.dateFrom;
  const dateTo = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};

  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);

  const relevant = rows.filter(r => {
    const d = r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date,"Asia/Kolkata","yyyy-MM-dd") : String(r.Order_Date).trim();
    return d >= dateFrom && d <= dateTo &&
    (r.Payment_Status === "Pending" || r.Payment_Status === "on account" || !r.Payment_Status);
  });

  const orders = relevant.map(r => ({
    id: r.Submission_ID,
    date: r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date,"Asia/Kolkata","yyyy-MM-dd") : String(r.Order_Date).trim(),
    phone: r.Phone,
    name: r.Customer_Name,
    total: Math.round(Number(r.Net_Total) || 0),
    status: r.Payment_Status || "Pending",
    meal: r.Meal_Type
  }));

  return {success:true, orders};
}

/**
 * Reconcile a bank-statement export against unpaid SK_Orders.
 *
 * Matching is two-tiered. A transaction matches ONLY when there is a single
 * unambiguous mapping from tx to one or more orders of the SAME customer.
 *
 * Customer-Name gate: every significant token (length > 2) of the
 * customer's name must appear in the transaction narration (case/space
 * insensitive). Anything failing this gate is skipped immediately.
 *
 * Tier-1 — strict same-date match (transaction date == order date):
 *   * (1a) single order's Net_Total == tx amount, OR
 *   * (1b) sum of THAT day's B+L+D for the customer == tx amount.
 *
 * Tier-2 — contiguous date-sorted range (no tx-date constraint):
 *   Customer's unpaid orders are sorted by date; we look for any
 *   contiguous subrange whose Net_Total sum == tx amount. This covers
 *   the realistic case where a customer transfers (e.g.) Rs.500 to
 *   clear their last three days' B+L+D in one go. Tier-2 only fires if
 *   Tier-1 didn't already find a hit, AND only when EXACTLY ONE
 *   contiguous range matches — multiple matches => ambiguous => manual.
 *
 * On a clean tier-1 or tier-2 hit, those orders are auto-marked Paid in
 * the same call. Everything else returns "Pending Manual Review" so the
 * admin handles it by hand. No fuzzy / partial / name-only matching —
 * those produced false positives in earlier UAT testing.
 */
function reconcileTransactions(body) {
  const transactions = body.transactions; // [{date, description, amount}]
  const dateFrom     = body.dateFrom;
  const dateTo       = body.dateTo;
  const autoMark     = body.autoMark !== false;  // default true; UI can pass false to preview only

  const unpaid = getUnpaidOrdersData({dateFrom, dateTo}).orders;

  // Group unpaid by phone for fast name-lookup
  const groupedByPhone = {};
  unpaid.forEach(o => {
    const key = String(o.phone || "").trim();
    if (!key) return;
    if (!groupedByPhone[key]) groupedByPhone[key] = { name: o.name, orders: [] };
    groupedByPhone[key].orders.push(o);
  });

  const sidsToMarkPaid = [];

  const results = transactions.map(tx => {
    const txAmount = Math.round(Number(tx.amount) || 0);
    const txDesc   = String(tx.description || "").toLowerCase();
    const txDate   = String(tx.date || "").trim();   // YYYY-MM-DD expected

    let bestMatch = {
      status: "Pending Manual Review",
      reason: "No exact Name + Amount match found",
      matchedOrders: []
    };

    // Walk every customer; STOP on first perfect hit. Tier priority:
    //   Tier-1: tx_date == order_date  (single order OR same-day B+L+D sum)
    //   Tier-2: any contiguous date-sorted range of THIS customer's
    //           unpaid orders summing to tx amount (no tx_date constraint;
    //           covers the "customer paid for last 3 days in one transfer"
    //           pattern). Must be exactly ONE such range or it's ambiguous.
    for (const phone in groupedByPhone) {
      const data = groupedByPhone[phone];
      const nameParts = data.name.toLowerCase().replace(/[^a-z ]/g, "").split(" ").filter(x => x.length > 2);
      const cleanDesc = txDesc.replace(/[^a-z]/g, "");
      const nameMatch = nameParts.length > 0 && nameParts.every(part => cleanDesc.includes(part));
      if (!nameMatch) continue;

      // ─── Tier 1: same-day strict match ─────────────────────────────────
      const sameDayOrders = data.orders.filter(o => o.date === txDate);
      if (sameDayOrders.length > 0) {
        // (1a) Single-order amount match on same date
        const singleMatch = sameDayOrders.find(o => Math.round(o.total) === txAmount);
        if (singleMatch) {
          bestMatch = {
            status: "Match",
            matchType: "Date + Name + Amount (single order)",
            phone: phone,
            name: data.name,
            date: txDate,
            matchedOrders: [singleMatch],
            total: singleMatch.total
          };
          sidsToMarkPaid.push(singleMatch.id);
          break;
        }
        // (1b) Sum of customer's orders on the SAME DATE matches tx amount
        const dailySum = Math.round(sameDayOrders.reduce((s, o) => s + o.total, 0));
        if (dailySum === txAmount) {
          bestMatch = {
            status: "Match",
            matchType: "Date + Name + Amount (B+L+D daily sum)",
            phone: phone,
            name: data.name,
            date: txDate,
            matchedOrders: sameDayOrders,
            total: dailySum
          };
          sameDayOrders.forEach(o => sidsToMarkPaid.push(o.id));
          break;
        }
      }

      // ─── Tier 2: contiguous date-sorted range sum ──────────────────────
      // Covers the case: customer paid Rs.500 for unpaid orders across
      // May 12-14 (different meals across different days). Tx date may
      // be any date, often a few days after the last order.
      const sorted = data.orders
        .slice()
        .sort((a, b) =>
          String(a.date || "").localeCompare(String(b.date || "")) ||
          String(a.id || "").localeCompare(String(b.id || ""))
        );
      const ranges = [];
      for (let i = 0; i < sorted.length; i++) {
        let sum = 0;
        for (let j = i; j < sorted.length; j++) {
          sum += Math.round(sorted[j].total);
          if (sum === txAmount) {
            ranges.push(sorted.slice(i, j + 1));
            break;            // exact hit — no point extending further from i
          }
          if (sum > txAmount) break;   // overshot — try a different start
        }
        // Safety cap: a single customer with hundreds of unpaid orders is
        // unrealistic, but bound the work so a pathological case can't hang.
        if (ranges.length > 5) break;
      }

      if (ranges.length === 1) {
        const range = ranges[0];
        const dates = Array.from(new Set(range.map(o => o.date)));
        const matchType = range.length === 1
          ? "Name + Amount (single order — different tx date)"
          : (dates.length === 1
              ? "Name + Amount (same-date B+L+D)"
              : "Name + Amount (multi-date contiguous, " + dates.length + " dates)");
        bestMatch = {
          status: "Match",
          matchType: matchType,
          phone: phone,
          name: data.name,
          dateRange: dates.length > 1 ? (dates[0] + " to " + dates[dates.length - 1]) : dates[0],
          matchedOrders: range,
          total: Math.round(range.reduce((s, o) => s + o.total, 0))
        };
        range.forEach(o => sidsToMarkPaid.push(o.id));
        break;
      }

      if (ranges.length > 1) {
        // Ambiguous — multiple contiguous subranges hit the same total.
        // Don't auto-mark; expose candidates so the admin can pick one.
        bestMatch = {
          status: "Pending Manual Review",
          reason: "Multiple possible matches (" + ranges.length + " candidate ranges) — pick the right one manually",
          phone: phone,
          name: data.name,
          candidates: ranges.map(r => ({
            orders: r,
            dates:  Array.from(new Set(r.map(o => o.date))),
            total:  Math.round(r.reduce((s, o) => s + o.total, 0))
          }))
        };
        break;
      }
      // No range matched for this customer — keep scanning others.
    }

    return Object.assign({ transaction: tx }, bestMatch);
  });

  // Auto-mark all perfectly-matched orders Paid in one shot.
  let markedCount = 0;
  if (autoMark && sidsToMarkPaid.length) {
    try {
      const r = markOrdersPaidBulk({ submissionIds: [...new Set(sidsToMarkPaid)] });
      markedCount = (r && r.updated) || 0;
    } catch (e) {
      console.warn("reconcileTransactions: auto-mark failed:", e.message);
    }
  }

  return {
    success: true,
    results: results,
    matched_count: results.filter(r => r.status === "Match").length,
    pending_count: results.filter(r => r.status === "Pending Manual Review").length,
    auto_marked:   markedCount
  };
}

function markOrdersPaidBulk(body) {
  const sids = body.submissionIds;
  if (!sids || !sids.length) return {success:false, error:"submissionIds required"};

  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const headers = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0];
  const hIdx = {};
  headers.forEach((h,i) => { hIdx[h] = i+1; });

  const rows = getAllRows(ws);
  let updated = 0;

  rows.forEach(r => {
    if (sids.includes(String(r.Submission_ID)) &&
        (r.Payment_Status === "Pending" || r.Payment_Status === "on account" || !r.Payment_Status)) {
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue("Paid");
      updated++;
    }
  });

  return {success:true, updated};
}
