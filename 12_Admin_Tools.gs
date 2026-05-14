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
