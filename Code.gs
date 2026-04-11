// ============================================================
// SVAADH KITCHEN — Code.gs (New System)
// One Google Sheet, clean schema, no Tally dependency
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
// SECRETS & KEYS: Setup these in Google Apps Script Project Settings > Script Properties
const SP             = PropertiesService.getScriptProperties();
const SHEET_ID       = SP.getProperty("SHEET_ID");
const ADMIN_PIN      = SP.getProperty("ADMIN_PIN") || "7532";
const KITCHEN_PIN    = SP.getProperty("KITCHEN_PIN") || "7284";
const PLACE_ID       = SP.getProperty("PLACE_ID") || "";
const GOOGLE_PLACES_API_KEY = SP.getProperty("GOOGLE_PLACES_API_KEY") || "";

const CODE_VERSION   = 14.1; // Prepaid Standard (Legacy Removal)
const LEDGER_FOLDER  = "Svaadh Customer Ledgers";
// ─────────────────────────────────────────────────────────────

// Sheet tab names
const TAB_ORDERS     = "SK_Orders";
const TAB_CUSTOMERS  = "SK_Customers";
const TAB_MENU       = "SK_Daily_Menu";
const TAB_BF_MASTER  = "SK_Master_Breakfast";
const TAB_SABJI      = "SK_Master_Sabjis";
const TAB_AREAS      = "SK_Areas";
const TAB_WALLET     = "SK_Wallet"; // Holds prepaid balances
const TAB_REFUNDS    = "SK_Refunds"; // Manual refund requests

// Canonical SK_Wallet column schema — NEVER reorder these
const WALLET_HEADERS = ["Phone", "Customer_Name", "Txn_Type", "Amount", "Verified", "Reference_ID", "Timestamp"];

// ── COLUMN LAYOUT — SK_Orders ────────────────────────────────
// A   Submission_ID
// B   Submitted_At
// C   Order_Date
// D   Meal_Type
// E   Customer_Name
// F   Phone
// G   Area
// H   Wing
// I   Flat
// J   Floor
// K   Society
// L   Full_Address
// M   Maps_Link
// N   Landmark
// O   Items_JSON
// P   Chapati
// Q   Without_Oil_Chapati
// R   Phulka
// S   Ghee_Phulka
// T   Jowar_Bhakri
// U   Bajra_Bhakri
// V   Dry_Sabji_Mini
// W   Dry_Sabji_Full
// X   Curry_Sabji_Mini
// Y   Curry_Sabji_Full
// Z   Dal
// AA  Rice
// AB  Salad
// AC  Curd
// AD  BF_Item_1
// AE  BF_Qty_1
// AF  BF_Item_2
// AG  BF_Qty_2
// AH  BF_Item_3
// AI  BF_Qty_3
// AJ  BF_Item_4
// AK  BF_Qty_4
// AL  Special_Notes
// AM  Food_Subtotal
// AN  Delivery_Charge
// AO  Discount_Amount
// AP  Net_Total
// AQ  Payment_Method
// AR  Payment_Status
// AS  Payment_Freq
// AT  First_Time
// AU  Source

const CUSTOMERS_HEADERS = [
  "Phone","Customer_Name","Area","Wing","Flat","Floor","Society","Full_Address",
  "Maps_Link","Landmark","Payment_Freq","Created_At","Ledger_Sheet_ID","PIN","Meal_Addresses",
  "Review_Promo_Count", "Review_Reward_Claimed"
];

const ORDERS_HEADERS = [
  "Submission_ID","Submitted_At","Order_Date","Meal_Type",
  "Customer_Name","Phone","Area","Wing","Flat","Floor","Society","Full_Address","Maps_Link","Landmark",
  "Items_JSON",
  "Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
  "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full",
  "Dal","Rice","Salad","Curd",
  "BF_Item_1","BF_Qty_1","BF_Item_2","BF_Qty_2","BF_Item_3","BF_Qty_3","BF_Item_4","BF_Qty_4",
  "Special_Notes_Kitchen","Special_Notes_Delivery",
  "Food_Subtotal","Delivery_Charge","Discount_Amount","Review_Discount","Net_Total",
  "Payment_Method","Payment_Status","Payment_Freq","First_Time","Source","Refund_Preference", "Packed"
];

// Item colKey → Orders column name mapping (for quick lookup)
const ITEM_COL_MAP = {
  "Chapati":"Chapati","Without_Oil_Chapati":"Without_Oil_Chapati",
  "Phulka":"Phulka","Ghee_Phulka":"Ghee_Phulka",
  "Jowar_Bhakri":"Jowar_Bhakri","Bajra_Bhakri":"Bajra_Bhakri",
  "Dry_Sabji_Mini":"Dry_Sabji_Mini","Dry_Sabji_Full":"Dry_Sabji_Full",
  "Curry_Sabji_Mini":"Curry_Sabji_Mini","Curry_Sabji_Full":"Curry_Sabji_Full",
  "Dal":"Dal","Rice":"Rice","Salad":"Salad","Curd":"Curd",
  // Legacy colKey aliases from order.html
  "L_CHAPATI":"Chapati","L_WO_CHAPATI":"Without_Oil_Chapati","L_PHULKA":"Phulka","L_GHEE_PHULKA":"Ghee_Phulka",
  "L_JOWAR":"Jowar_Bhakri","L_BAJRA":"Bajra_Bhakri",
  "L_DRY_MINI":"Dry_Sabji_Mini","L_DRY_FULL":"Dry_Sabji_Full",
  "L_CURRY_MINI":"Curry_Sabji_Mini","L_CURRY_FULL":"Curry_Sabji_Full",
  "L_DAL":"Dal","L_RICE":"Rice","L_SALAD":"Salad","L_CURD":"Curd",
  "D_CHAPATI":"Chapati","D_WO_CHAPATI":"Without_Oil_Chapati","D_PHULKA":"Phulka","D_GHEE_PHULKA":"Ghee_Phulka",
  "D_JOWAR":"Jowar_Bhakri","D_BAJRA":"Bajra_Bhakri",
  "D_DRY_MINI":"Dry_Sabji_Mini","D_DRY_FULL":"Dry_Sabji_Full",
  "D_CURRY_MINI":"Curry_Sabji_Mini","D_CURRY_FULL":"Curry_Sabji_Full",
  "D_DAL":"Dal","D_RICE":"Rice","D_SALAD":"Salad","D_CURD":"Curd",
  "B_CURD":"Curd"
};

// ── BUSINESS CONTEXT (for chatbot) ──────────────────────────
const BUSINESS_CONTEXT = {
  name: "Svaadh Kitchen",
  type: "Cloud Kitchen",
  tagline: "Wholesome homemade vegetarian meals, straight from our kitchen to your plate.",
  about: "Svaadh Kitchen is a home-based vegetarian cloud kitchen in Hadapsar, Pune, serving fresh and wholesome homemade meals since August 2023 (over 2.5 years). We specialize in homemade vegetarian food, offering breakfast, lunch, and dinner with a changing daily sabji menu. We deliver to Bhosale Garden, Magarpatta, Amanora, DP Road, Triveni Nagar, and 10 other nearby areas. Customers can also opt for Self Pickup to waive all fees.",
  vision: "To make homemade vegetarian meals easily accessible and affordable for everyone, while maintaining taste, quality, and consistency.",
  locations_served: ["Bhosale Garden", "Magarpatta", "Amanora", "DP Road", "Triveni Nagar", "Malwadi", "SadeSatraNali", "Kirtane Baug", "Tupe Patil Road", "BG Shirke Road", "Vaiduwadi", "Pune-Solapur Road", "Vihar Chowk", "Mandai", "Gadital", "Self Pickup"],
  order_cutoffs: { breakfast: "before 7:00 AM", lunch: "before 9:30 AM", dinner: "before 5:00 PM", closed_on: "Sunday" },
  delivery: {
    free_area: "Bhosale Garden, Triveni Nagar, Self Pickup",
    charge: "₹10 per meal for others if subtotal is below ₹100. Free for Bhosale Garden, Triveni Nagar and Self Pickup always.",
    per_meal_address: "Each meal can go to a different address — breakfast at home, lunch at office, dinner back home."
  },
  menu: {
    note: "Today's sabji (dry and curry) changes daily — shown in the order form. Breakfast items also rotate daily.",
    breads: [
      {name:"Chapati", price:9, unit:"per piece"},
      {name:"Without Oil Chapati", price:8, unit:"per piece"},
      {name:"Phulka", price:7, unit:"per piece"},
      {name:"Ghee Phulka", price:10, unit:"per piece"},
      {name:"Jowar Bhakri", price:20, unit:"per piece"},
      {name:"Bajra Bhakri", price:20, unit:"per piece"}
    ],
    sabji: [
      {name:"Dry Sabji Mini (100ml)", price:22},
      {name:"Dry Sabji Full (250ml)", price:45},
      {name:"Curry Sabji Mini (100ml)", price:22},
      {name:"Curry Sabji Full (250ml)", price:45}
    ],
    basics: [
      {name:"Dal (200ml)", price:22},
      {name:"Rice (100g)", price:12},
      {name:"Salad (40g)", price:6},
      {name:"Curd (50g)", price:12}
    ],
    breakfast: "Rotating daily (₹35–₹70). Items include Kanda Poha ₹35, Ghee Upma ₹40, Sabudana Khichdi ₹40, Tikhi Pudi ₹45, Idli Chutney ₹45, Masala Dosa ₹45, Aloo Paratha ₹50, Veg Sandwich ₹50, Thalipeeth ₹50, Ghee Sheera ₹50, Paneer Paratha ₹70. Curd available extra ₹12. Check the order form for today's options.",
    breakfast_note: "Curd (50g ₹12) is available as an add-on for breakfast — not included by default. Pure Ghee is used to make breakfast items."
  },
  discounts: {
    tier1: "5% off when the day total is ₹300 or more",
    tier2: "7.5% off when the day total is ₹450 or more",
    note: "Discounts are applied automatically per day's total when placing an order."
  },
  payment: {
    options: ["Svaadh Wallet (Prepaid)", "UPI", "Prepaid Wallet Billing"],
    upi_id: "9819969682@hdfc",
    prepaid_wallet: "Prepaid Wallet Billing operates as a prepaid wallet. Customers must maintain a top-up balance, and orders are deducted immediately."
  },
  ordering: {
    order_url: "https://www.svaadhkitchen.in/order.html",
    process: "Open the order form → enter phone number → fill address → pick dates → choose meals → review bill → pay via Wallet or UPI.",
    advance: "Select multiple dates on the calendar to order for the full week in one go.",
    edit_cancel: "Use 'View/Edit existing orders' on the order form home screen to edit or cancel before the cutoff.",
    no_login: "No login needed — phone number is your identity. Details are saved automatically."
  },
  contact: {
    phone_primary: "9930748908",
    phone_alt: "9819969682",
    whatsapp: "+91 93222 46765",
    whatsapp_link: "https://wa.me/919322246765",
    whatsapp_group: "https://chat.whatsapp.com/EpLv7mtYipm61ScKjbOiuk",
    email: "svaadh.kitchen@gmail.com",
    google_page: "https://share.google/UnZM2xcLOF2QVO9cj"
  }
};

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;
  const action = p.parameter ? p.action : (e.parameter.action || ""); // Fix for inconsistent parameter access
  const pin = p.pin || "";
  
  try {
    if (action === "version") return jsonRes({version: CODE_VERSION, status:"ok"});
    if (action === "getAreas") return jsonRes(getAreas());
    if (action === "getCustomer") return jsonRes(getCustomer(p.phone));
    if (action === "verifyLogin") return jsonRes(verifyLogin(p.phone, p.pin));
    if (action === "setPin") {
      const profile = { phone: p.phone, pin: p.pin };
      _upsertCustomer(getSpreadsheet(), profile);
      return jsonRes({success:true});
    }
    if (action === "getWeeklyMenu") return jsonRes(getWeeklyMenu());
    
    // Auth Tiers (STRICTLY ISOLATED)
    const isAdmin = (pin === ADMIN_PIN && pin !== "");
    const isStaff = (pin === KITCHEN_PIN || pin === ADMIN_PIN) && pin !== "";

    // KITCHEN & DRIVER ACCESS (Staff PIN ONLY)
    if (action === "getKitchenSummary") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(getKitchenSummary(p.date));
    }
    if (action === "getDriverOrders") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(getDriverOrders(p.date));
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
    
    // Fallback menu / orders for customers (legacy)
    if (action === "getMenu") return jsonRes(getMenu(p.date));
    if (action === "getWeeklyMenu") return jsonRes(getWeeklyMenu());
    if (action === "getCustomerOrders") return jsonRes(getCustomerOrders(p.phone));
    if (action === "getWalletValue") return jsonRes({wallet_balance: _calculateWalletBalance(p.phone)});
    if (action === "getDayTotalsForDates") return jsonRes(getDayTotalsForDates(p.phone, p.dates));

    return jsonRes({error:"Unknown action or Access Denied"});
  } catch(err) {
    return jsonRes({error: err.message});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body._action || "";
    const pin = body.pin || "";
    const isAdmin = (pin === ADMIN_PIN && pin !== "");
    const isStaff = (pin === KITCHEN_PIN || pin === ADMIN_PIN) && pin !== "";

    // Customer actions (pinned via their own phone/PIN handled inside functions)
    if (action === "deleteOrder") return jsonRes(deleteOrder(body.phone, body.rowId, body.refundType));
    
    // Delivery Actions (Staff PIN ONLY)
    if (action === "markDelivered") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(markDelivered(body));
    }
    if (action === "batchMarkEnRoute") {
      if (!isStaff) return jsonRes({error:"STRICT STAFF PIN REQUIRED"});
      return jsonRes(batchMarkEnRoute(body));
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
    if (action === "payAllPendingWithWallet") return jsonRes(payAllPendingWithWallet(body));

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

    // Regular order submission
    return jsonRes(submitOrder(body));
  } catch(err) {
    return jsonRes({error: err.message});
  }
}

function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ──────────────────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getOrCreateTab(ss, name, headers) {
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
  }
  
  if (headers && headers.length > 0) {
    const lastCol = ws.getLastColumn();
    const currentHeaders = lastCol > 0 ? ws.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||"").trim()) : [];
    
    // Force header row synchronization by explicitly setting range if any mismatch
    headers.forEach((h, i) => {
      if (currentHeaders[i] !== h) {
        ws.getRange(1, i + 1).setValue(h)
          .setFontWeight("bold")
          .setBackground("#c0392b")
          .setFontColor("white");
        if (i === 0) ws.setFrozenRows(1);
        // Force certain columns to stay as Plain Text to preserve leading zeros
        if (h === "Phone" || h === "PIN") {
          ws.getRange(1, i + 1, ws.getMaxRows(), 1).setNumberFormat("@");
        }
      }
    });

    // CRITICAL: If headers were provided, ensure No Extra Columns exist beyond them
    // This prevents "Timestamp" duplicates if things drifted in legacy versions
    if (headers.length > 0 && ws.getLastColumn() > headers.length) {
      const extra = ws.getLastColumn() - headers.length;
      ws.deleteColumns(headers.length + 1, extra);
    }
  }
  return ws;
}

function getISTDate() {
  const now = new Date();
  // Cross-environment IST Date object
  return new Date(now.getTime() + (now.getTimezoneOffset() + 330) * 60000);
}

function getISTTimestamp() {
  return Utilities.formatDate(getISTDate(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}

function generateSubmissionID() {
  const ist = getISTDate();
  const dateStr = Utilities.formatDate(ist, "Asia/Kolkata", "yyyyMMdd");
  const rand = Math.floor(Math.random() * 9000) + 1000;
  return `SK-${dateStr}-${rand}`;
}

function headerIndex(ws) {
  // Returns {colName: 1-based-index} for the given sheet
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i + 1; });
  return idx;
}

function getAllRows(ws) {
  const last = ws.getLastRow();
  if (last < 2) return [];
  const data = ws.getRange(2, 1, last - 1, ws.getLastColumn()).getValues();
  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  return data.map((row, ri) => {
    const obj = {_row: ri + 2};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── SCHEMA INIT ──────────────────────────────────────────────
function initSchema() {
  const ss = getSpreadsheet();
  getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner"
  ]);
  getOrCreateTab(ss, TAB_BF_MASTER, ["ID","Name","Price","Active"]);
  getOrCreateTab(ss, TAB_SABJI,     ["ID","Name","Type","Active"]);
  getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  return {success: true, message: "Schema initialised"};
}

/**
 * Normalizes phone numbers for reliable comparison across Google Sheets.
 * Handles scientific notation (e.g., 9.87E+9) and trailing decimals (.0).
 */
function _normalizePhone(phone) {
  let p = String(phone || "").trim();
  if (p.includes(".")) p = p.split(".")[0];
  if (p.toUpperCase().includes("E+") && !isNaN(Number(p))) {
    p = String(Math.round(Number(p)));
  }
  return p;
}

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
    promoCount:         Number(r.Review_Promo_Count) || 0,
    wallet_balance:     _calculateWalletBalance(phone)
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
      promoCount:         Number(r.Review_Promo_Count) || 0,
      wallet_balance:     _calculateWalletBalance(phone)
    }
  };
}

// ── WALLET HELPER ──────────────────────────────────────────
function _calculateWalletBalance(phone) {
  if (!phone) return 0;
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const rows = getAllRows(ws);

  let balance = 0;
  const pStr = _normalizePhone(phone);

  rows.forEach(w => {
    const rPhone = _normalizePhone(w.Phone);
    if (rPhone !== pStr) return;

    // Only count verified transactions
    const rVer = String(w.Verified || "").trim().toUpperCase();
    if (rVer !== "TRUE" && rVer !== "YES" && rVer !== "VERIFIED") return;

    const rAmt  = Number(w.Amount) || 0;
    // Also check legacy columns where Txn_Type may have been stored in a "Balance" column
    const rType = String(w.Txn_Type || w.Balance || w.Txn_Type || "").trim().toLowerCase();

    if (rType.includes("recharge") || rType.includes("refund") || rType.includes("credit")) {
      balance += rAmt;
    } else if (rType.includes("order") || rType.includes("deduct") || rType.includes("payment")) {
      balance -= rAmt;
    }
  });

  return Math.round(balance * 100) / 100;
}

function saveMenu(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner"
  ]);
  const rows = ws.getDataRange().getValues();
  const headers = rows[0];
  const dIdx = headers.indexOf("Date");
  
  if (dIdx === -1) return {success:false, error:"Date column missing in SK_Daily_Menu"};

  const dateStr = body.date;
  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    let rDate = rows[i][dIdx];
    if (rDate instanceof Date) rDate = Utilities.formatDate(rDate, "Asia/Kolkata", "yyyy-MM-dd");
    if (String(rDate).trim() === dateStr) {
      existingRow = i + 1;
      break;
    }
  }

  const newRow = new Array(headers.length).fill("");
  const setVal = (h, val) => {
    const idx = headers.indexOf(h);
    if (idx >= 0) newRow[idx] = val;
  };

  setVal("Date", dateStr);
  setVal("Breakfast_JSON", JSON.stringify(body.breakfast || []));
  setVal("Lunch_Dry", body.lunch_dry || "");
  setVal("Lunch_Curry", body.lunch_curry || "");
  setVal("Dinner_Dry", body.dinner_dry || "");
  setVal("Dinner_Curry", body.dinner_curry || "");
  setVal("Cutoff_Breakfast", body.cutoff_breakfast || "");
  setVal("Cutoff_Lunch", body.cutoff_lunch || "");
  setVal("Cutoff_Dinner", body.cutoff_dinner || "");

  if (existingRow >= 0) {
    ws.getRange(existingRow, 1, 1, newRow.length).setValues([newRow]);
  } else {
    ws.appendRow(newRow);
  }
  return {success:true};
}


// ── GET MENU ─────────────────────────────────────────────────
function getMenu(dateStr) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => {
    const d = x.Date instanceof Date
      ? Utilities.formatDate(x.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(x.Date).trim();
    return d === dateStr;
  });

  // Breakfast master items
  const bfWs = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const bfRows = getAllRows(bfWs).filter(x => String(x.Active).toLowerCase() !== "false");
  const breakfast = bfRows.map(x => ({name: String(x.Name), price: Number(x.Price)}));

  if (!r) return {
    breakfast, lunch_dry:"", lunch_curry:"", dinner_dry:"", dinner_curry:"",
    cutoff_overrides:{}
  };

  const co = {};
  if (r && r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
  if (r && r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
  if (r && r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);

  // MERGE LOGIC: Start with master active items, then merge daily overrides
  const masterActive = breakfast;
  let dailyBf = [];
  try { if (r && r.Breakfast_JSON) dailyBf = JSON.parse(r.Breakfast_JSON); } catch(e) {}

  // Prioritize Daily selections (where specific prices or choices were made)
  // but ensure Master Active items are always present.
  const finalBreakfast = [...dailyBf];
  masterActive.forEach(m => {
    if (!finalBreakfast.some(d => d.name === m.name)) {
      finalBreakfast.push(m);
    }
  });

  return {
    breakfast:    finalBreakfast,
    lunch_dry:    r ? (r.Lunch_Dry || "") : "",
    lunch_curry:  r ? (r.Lunch_Curry || "") : "",
    dinner_dry:   r ? (r.Dinner_Dry || "") : "",
    dinner_curry: r ? (r.Dinner_Curry || "") : "",
    cutoff_overrides: co
  };
}

// ── GET WEEKLY MENU (next 7 days) ────────────────────────────
function getWeeklyMenu() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, []);
  const rows = getAllRows(ws);

  // Breakfast master items
  const bfWs = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const bfRows = getAllRows(bfWs).filter(x => String(x.Active).toLowerCase() !== "false");
  const defaultBreakfast = bfRows.map(x => ({name: String(x.Name), price: Number(x.Price)}));

  // Build a map: dateStr → row for quick lookup
  const menuMap = {};
  rows.forEach(x => {
    const d = x.Date instanceof Date
      ? Utilities.formatDate(x.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(x.Date).trim();
    menuMap[d] = x;
  });

  // Generate next 7 days
  const today = new Date();
  const days = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(today);
    d.setDate(d.getDate() + i);
    const dateStr = Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    const dayName = Utilities.formatDate(d, "Asia/Kolkata", "EEEE");
    const displayDate = Utilities.formatDate(d, "Asia/Kolkata", "dd MMM");

    const r = menuMap[dateStr];
    let bfDaily = [];
    try {
      if (r && r.Breakfast_JSON) bfDaily = JSON.parse(r.Breakfast_JSON);
    } catch(e) {}

    // Merge Master + Daily
    const finalBf = [...bfDaily];
    defaultBreakfast.forEach(m => {
      if (!finalBf.some(d => d.name === m.name)) finalBf.push(m);
    });

    days.push({
      date: dateStr,
      dayName: dayName,
      displayDate: displayDate,
      breakfast: finalBf,
      lunch_dry: r ? (r.Lunch_Dry || "") : "",
      lunch_curry: r ? (r.Lunch_Curry || "") : "",
      dinner_dry: r ? (r.Dinner_Dry || "") : "",
      dinner_curry: r ? (r.Dinner_Curry || "") : "",
      menuSet: !!r
    });
  }

  return { success: true, days: days };
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
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  const hIdx = headerIndex(ws);

  const totalCols = ws.getLastColumn();
  const row = new Array(totalCols).fill("");
  const set = (col, val) => { if (hIdx[col]) row[hIdx[col] - 1] = val; };

  set("Phone",         phone);
  set("Customer_Name", name);
  set("Txn_Type",      txnType);
  set("Amount",        amount);
  set("Verified",      isVerified ? "TRUE" : "FALSE");
  set("Reference_ID",  refId || "");
  set("Timestamp",     getISTTimestamp());

  ws.appendRow(row);
}

// ── SUBMIT ORDER ─────────────────────────────────────────────
function submitOrder(body) {
  const ss = getSpreadsheet();
  const ordersWs = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const profile   = body.profile || {};
  const orders    = body.orders  || [];   // [{date, meals:[{type,items,notes,subtotal,area}]}]

  const submittedAt  = getISTTimestamp();
  const payMethod    = body.payment_method  || "UPI";
  const payStatus    = body.payment_status  || "Pending";
  const firstTime    = profile.isFirstTime ? "Yes" : "No";
  const payFreq      = profile.payment_preference || "Daily Payment";

  // Build the header→index map once
  const hIdx = headerIndex(ordersWs);

  // Fetch free areas dynamically (replaces hardcoded FREE_AREA = "Bhosale Garden")
  const freeAreaNames = getAreas().filter(function(a){ return a.free; }).map(function(a){ return a.name; });
  const DELIVERY  = 10;
  const FREE_THR  = 150;

  const submissionIds = [];
  
  // Fetch existing orders once for all dates in this submission to calculate combined-day fees/discounts
  const submissionDates = orders.map(o => o.date);
  const existingDayTotals = getDayTotalsForDates(profile.phone, submissionDates.join(',')).dayTotals || {};

  // Fetch current promo state
  const custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  const cIdx   = headerIndex(custWs);
  const cRows  = getAllRows(custWs);
  const phoneStr = _normalizePhone(profile.phone);
  const cRowIdx = cRows.findIndex(r => _normalizePhone(r.Phone) === phoneStr);
  let promoCount = 0;
  if (cRowIdx !== -1) {
    promoCount = Number(cRows[cRowIdx].Review_Promo_Count) || 0;
  }

  for (const order of orders) {
    const orderDate = order.date;
    const existingDateInfo = (existingDayTotals[orderDate] || {});
    
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
    const getDisc = (sub) => submissionDayFoodTotal > 0 ? Math.round(submissionDateDiscAmt * (sub / submissionDayFoodTotal)) : 0;

    for (const meal of order.meals) {
      const sid = generateSubmissionID();
      submissionIds.push(sid);
      meal._sid = sid; // carry sid for ledger
      
      const mealType = meal.type;
      const sub = Number(meal.subtotal) || 0;
      const mealArea = meal.area || profile.area || "";
      const items  = meal.items || [];   // [{colKey, qty}]
      const nKitchen = meal.notesKitchen || "";
      const nDelivery = meal.notesDelivery || "";
      
      // Get combined totals for THIS specific meal type (prev + current)
      const prevMealSub = (existingDateInfo[mealType] || {}).subtotal || 0;
      const combinedMealSub = sub + prevMealSub;
      
      // Delivery & Fee logic (matches frontend)
      const isPickup  = (mealArea === "Self Pickup");
      const isDayFree = (combinedDayTotal >= FREE_THR);
      const isFreeArea = freeAreaNames.includes(mealArea);

      let delCharge = 0;
      if (!isDayFree && !isPickup && !isFreeArea && sub > 0) {
        delCharge = DELIVERY;
      }

      let smallOrderFee = 0;
      if (!isDayFree && !isPickup && (mealType === "Lunch" || mealType === "Dinner") && sub > 0 && combinedMealSub < 50) {
        smallOrderFee = 10;
      }

      // Calculation of credits for previously paid fees on the same day (Retroactive waiver)
      let dateDeliveryCredit = 0;
      let dateSmallFeeCredit = 0;
      if (isDayFree) {
        Object.keys(existingDateInfo).forEach(mType => {
          dateDeliveryCredit += (Number(existingDateInfo[mType].delivery_charged) || 0);
          dateSmallFeeCredit += (Number(existingDateInfo[mType].small_fee_charged) || 0);
        });
      }
      const totalDateCredit = dateDeliveryCredit + dateSmallFeeCredit;
      const mealCredit = submissionDayFoodTotal > 0 ? (totalDateCredit * (sub / submissionDayFoodTotal)) : 0;
      
      const discAmt = getDisc(sub);
      const inflationSurcharge = Math.ceil(sub / 10);
      
      // Google Review Promo Logic (10% OFF per meal)
      let reviewDiscount = 0;
      if (promoCount > 0 && sub > 0) {
        reviewDiscount = Math.round(sub * 0.10);
        promoCount--;
      }
      
      const netTotal  = sub + delCharge + smallOrderFee + inflationSurcharge - discAmt - mealCredit - reviewDiscount;
      meal._reviewDiscount = reviewDiscount; // carry for set() below


      // Build items JSON
      const itemsObj = {};
      items.forEach(({colKey, qty}) => {
        const canonical = ITEM_COL_MAP[colKey] || colKey;
        itemsObj[canonical] = qty;
      });

      // Address fields handling (Sanitized for Pickup)
      const wing    = isPickup ? "" : (meal.wing    || profile.wing    || "");
      const flat    = isPickup ? "" : (meal.flat    || profile.flat    || "");
      const floor   = isPickup ? "" : (meal.floor   || profile.floor   || "");
      const society = isPickup ? "" : (meal.society || profile.society || "");
      const area    = mealArea;

      const fullAddr = (area === "Self Pickup")
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
      set("Phone",               profile.phone    || "");
      set("Area",                area);
      set("Wing",                wing);
      set("Flat",                flat);
      set("Floor",               floor);
      set("Society",             society);
      set("Full_Address",        fullAddr);
      set("Maps_Link",           mapsLink);
      set("Landmark",            landmark);
      if (!hIdx["Small_Order_Fee"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Small_Order_Fee");
        hIdx["Small_Order_Fee"] = ordersWs.getLastColumn();
      }
      if (!hIdx["Inflation_Surcharge"]) {
        ordersWs.getRange(1, ordersWs.getLastColumn() + 1).setValue("Inflation_Surcharge");
        hIdx["Inflation_Surcharge"] = ordersWs.getLastColumn();
      }
      set("Items_JSON",          JSON.stringify(itemsObj));
      set("Special_Notes_Kitchen",  nKitchen);
      set("Special_Notes_Delivery", nDelivery);
      set("Food_Subtotal",       sub);
      set("Small_Order_Fee",     smallOrderFee);
      set("Inflation_Surcharge", inflationSurcharge);
      set("Delivery_Charge",     delCharge);
      set("Discount_Amount",     discAmt);
      if (hIdx["Review_Discount"]) {
        set("Review_Discount",   meal._reviewDiscount || 0);
      }
      set("Net_Total",           netTotal);
      
      let pStat = payStatus;
      // ════ WALLET DEDUCTION LOGIC ════
      if (payMethod === "Wallet") {
        let currentBalance = _calculateWalletBalance(profile.phone);
        
        if (currentBalance >= netTotal) {
          _appendWalletTransaction(profile.phone || "", profile.name || "Customer", "Order Deduction", netTotal, true, sid);
          pStat = "Wallet Paid";
        } else {
          pStat = "Pending"; // Wallet failed, fallback to pending
        }
      }
      
      set("Payment_Method",      payMethod);
      set("Payment_Status",      pStat);
      set("Payment_Freq",        payFreq);
      set("First_Time",          firstTime);
      set("Source",              "WebApp");

      // Fill individual item columns
      if (mealType === "Breakfast") {
        // Breakfast: dynamic items go to BF_Item_N/BF_Qty_N
        let bfSlot = 1;
        items.forEach(({colKey, qty}) => {
          if (bfSlot > 4) return;
          const displayName = colKey === "B_CURD" ? "Curd" : colKey;
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
          set(canonical, qty);
        });
      }

      ordersWs.appendRow(row);
    }
  }

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
    const realRow = cRowIdx + 2;
    custWs.getRange(realRow, cIdx["Review_Promo_Count"]).setValue(promoCount);
  }

  return {success: true, submissionId: submissionIds[0] || ""};
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
      if (hIdx[col] && val !== undefined && val !== null) {
        ws.getRange(rowNum, hIdx[col]).setValue(val);
      }
    };
    update("Customer_Name", profile.name || "");
    update("Area",          profile.area || "");
    update("Wing",          profile.wing || "");
    update("Flat",          profile.flat || "");
    update("Floor",         profile.floor || "");
    update("Society",       profile.society || "");
    update("Full_Address",  fullAddr);
    update("Maps_Link",     profile.maps || "");
    update("Landmark",      profile.landmark || "");
    update("Payment_Freq",  profile.payment_preference || "Daily Payment");
    // Only update PIN if provided (non-empty)
    if (profile.pin) update("PIN", profile.pin);
    if (profile.meal_addresses) update("Meal_Addresses", profile.meal_addresses);
  } else {
    // For new records, construct a clean Row Array mapping directly to our schema
    const newRow = CUSTOMERS_HEADERS.map(h => {
      let val = "";
      switch(h) {
        case "Phone":           val = profile.phone || ""; break;
        case "Customer_Name":   val = profile.name || ""; break;
        case "Area":            val = profile.area || ""; break;
        case "Wing":            val = profile.wing || ""; break;
        case "Flat":            val = profile.flat || ""; break;
        case "Floor":           val = profile.floor || ""; break;
        case "Society":         val = profile.society || ""; break;
        case "Full_Address":    val = fullAddr; break;
        case "Maps_Link":       val = profile.maps || ""; break;
        case "Landmark":        val = profile.landmark || ""; break;
        case "Payment_Freq":    val = profile.payment_preference || "Daily Payment"; break;
        case "Created_At":      val = getISTTimestamp(); break;
        case "PIN":             val = profile.pin || ""; break;
        case "Meal_Addresses":  val = profile.meal_addresses || ""; break;
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

// ── GET CUSTOMER ORDERS ──────────────────────────────────────
// ── GET DAY TOTALS FOR DATES (used to compute combined-day fees) ─
// Returns existing meal subtotals per date for the given phone,
// excluding the current cart being built (which is not yet placed).
function getDayTotalsForDates(phone, datesParam) {
  if (!phone || !datesParam) return { dayTotals: {} };
  const dates = String(datesParam).split(',').map(d => d.trim()).filter(Boolean);
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);

  const result = {};
  dates.forEach(d => { result[d] = {}; });

  rows.filter(r => {
    const rPhone = String(r.Phone || '').trim();
    const pStat  = String(r.Payment_Status || '').toLowerCase();
    if (rPhone !== String(phone).trim()) return false;
    if (pStat.includes('deleted') || pStat.includes('cancelled')) return false;
    const rDate = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r.Order_Date).trim();
    return dates.includes(rDate);
  }).forEach(r => {
    const rDate = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r.Order_Date).trim();
    const meal = String(r.Meal_Type).trim();
    if (!result[rDate][meal]) result[rDate][meal] = { subtotal: 0, delivery_charged: 0, discount_applied: 0, small_fee_charged: 0, count: 0 };
    result[rDate][meal].subtotal       += Number(r.Food_Subtotal    || 0);
    result[rDate][meal].delivery_charged += Number(r.Delivery_Charge || 0);
    result[rDate][meal].discount_applied += Number(r.Discount_Amount || 0);
    result[rDate][meal].small_fee_charged += Number(r.Small_Order_Fee || 0);
    result[rDate][meal].count++;
  });

  return { dayTotals: result };
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
      return {
        rowId:              r.Submission_ID,
        date:               fmtD(r),
        meal:               r.Meal_Type,
        summary:            _buildSummary(r),
        total:              r.Net_Total,
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
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
      return {
        rowId:              r.Submission_ID,
        date:               fmtD(r),
        meal:               r.Meal_Type,
        summary:            _buildSummary(r),
        total:              r.Net_Total,
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
        deliveredAt:        delTracker.deliveredAt,
        enRouteAt:          delTracker.enRouteAt
      };
    });

  return {orders: upcoming, past_orders: past, wallet_balance: _calculateWalletBalance(phone)};
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

// ── DELETE ORDER (with Refund Logic) ─────────────────────────
function deleteOrder(phone, rowId, refundType) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const now = getISTDate();
  let msg = "Order deleted successfully";
  const today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
  const hourIST = now.getHours() + now.getMinutes() / 60;
  const CUTOFFS = { Breakfast: 7, Lunch: 9.5, Dinner: 17 };

  const r = rows.find(x => {
    // Deep Normalization: Keep only digits to handle commas, decimals (123.0), or scientific notation
    const cleanSheetId = String(x.Submission_ID || "").replace(/\D/g, "");
    const cleanTargetId = String(rowId || "").replace(/\D/g, "");
    const sheetPhone = String(x.Phone || "").trim();
    const targetPhone = String(phone || "").trim();
    return cleanSheetId === cleanTargetId && sheetPhone === targetPhone;
  });
  if (!r) {
    console.error(`CANCELLATION FAILED: Submission ID ${rowId} (Clean: ${String(rowId).replace(/\D/g,"")}) not found for phone ${phone}.`);
    return {success: false, error: "Order record not found in system."};
  }
  const orderDateStr = r.Order_Date instanceof Date
    ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
    : String(r.Order_Date).trim();
  if (orderDateStr < today) return {success: false, error: "Cannot delete past orders"};

  // Block deletion if cutoff has passed for today's orders
  if (orderDateStr === today) {
    const cutoffHour = CUTOFFS[r.Meal_Type];
    if (cutoffHour !== undefined && hourIST >= cutoffHour) {
      return {success: false, error: `Cutoff for ${r.Meal_Type} has already passed`};
    }
  }

  // GRACEFUL REFUND HANDLING with eligibility recalculation (Cases 1/2/3)
  const pStatStr = String(r.Payment_Status).toLowerCase();
  if (pStatStr === "paid" || pStatStr === "wallet paid") {
    const custName = r.Customer_Name || "Customer";
    const ordersWs2 = ws; // same sheet
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

    // Discount eligibility helper
    const discRate = (sub) => sub >= 450 ? 0.075 : sub >= 300 ? 0.05 : 0;
    const oldRate = discRate(oldDaySubtotal);
    const newRate = discRate(remainingDaySubtotal);

    // Calc over-discount on remaining orders: they received discount at oldRate
    // but now only deserve newRate
    let overDiscount = 0;
    if (oldRate > newRate) {
      const overOnRemaining = sameDayRows.reduce((s, x) => {
        const xSub = Number(x.Food_Subtotal) || 0;
        const oldD = Math.round(xSub * oldRate);
        const newD = Math.round(xSub * newRate);
        return s + (oldD - newD);
      }, 0);
      overDiscount = overOnRemaining;
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
      // 1. Delivery Clawback: sum up delivery on all remaining rows that were waived due to the 150 rule
      sameDayRows.forEach(x => {
        const xArea = x.Area || "";
        const xSub = Number(x.Food_Subtotal) || 0;
        if (xSub > 0 && isNonFree(xArea)) {
          // They should have paid delivery. Check if they were charged 0.
          if ((Number(x.Delivery_Charge) || 0) === 0) {
             deliveryOwed += 10;
          }
        }
        
        // 2. Small Order Fee Clawback: if Lunch/Dinner row < 50 and was waived
        const xMeal = String(x.Meal_Type).trim();
        if ((xMeal === "Lunch" || xMeal === "Dinner") && xSub > 0 && xSub < 50) {
          if ((Number(x.Small_Order_Fee) || 0) === 0) {
            smallFeeOwed += 10;
          }
        }
      });
    } else {
      // Partial drop handling: If they were always below 150, but a specific meal sub dropped below 100/50?
      // Actually, the user wants the 150 rule to be the primary toggle now.
      // But we still need to handle the case where day total was < 150, and they delete an order,
      // and a specific meal total drops.
      // Wait, the user said "remove the breakfast delivery free if lunch/dinner ordered, instead put this, 
      // day's total over ₹150 then free delivery across all."
      
      // Let's stick to the 150 rule as the only common waiver.
    }

    // Actual refund = what was charged on deleted row minus any amount now owed back
    const adjustment = overDiscount + deliveryOwed + smallFeeOwed;
    const rawRefund = Number(r.Net_Total) || 0;
    const refundAmt = Math.max(0, rawRefund - adjustment);

    // Multi-Payment Logic: If any OTHER order for this meal/date is Wallet Paid, 
    // force this refund to Wallet too (to keep the day's bookkeeping simple).
    const hasAnyOtherWalletPaid = sameDayRows.some(x => {
      const typeMatch = String(x.Meal_Type).trim() === deleteMeal;
      const statusMatch = String(x.Payment_Status).toLowerCase() === "wallet paid";
      return typeMatch && statusMatch;
    });

    let finalType = refundType;
    let msgSuffix = "";
    
    // Auto-detect wallet refund if current was wallet paid, overriding passed type
    const currentWasWallet = (pStatStr === "wallet paid");
    if (currentWasWallet) {
      finalType = "wallet";
    } else if (hasAnyOtherWalletPaid && refundType === "manual_upi") {
      finalType = "wallet";
      msgSuffix = " (Consolidated to Wallet since other items in this meal were Wallet Paid)";
    }

    if (finalType === "wallet") {
      _appendWalletTransaction(phone, custName, "Order Cancellation Refund", refundAmt, true, String(rowId));
      msg = `₹${refundAmt} refunded to Wallet${msgSuffix}`;
    }
    else if (finalType === "manual_upi") {
      const REF_HEADERS = ["Submission_ID","Phone","Name","Amount","Meal","Date","Status","Timestamp","Adjustment_Note","Refund_Mode"];
      const refWs = getOrCreateTab(ss, TAB_REFUNDS, REF_HEADERS);
      const note = adjustment > 0
        ? `Adjusted -₹${adjustment} (overDiscount:${overDiscount}, deliveryOwed:${deliveryOwed}, smallFeeOwed:${smallFeeOwed}, bfDeliveryOwed:${bfDeliveryOwed})`
        : "";
      refWs.appendRow([rowId, phone, custName, refundAmt, r.Meal_Type, orderDateStr, "Pending", now, note, "upi"]);
      msg = `₹${refundAmt} refund request raised in Approvals`;
    }
  } 
  // ── SOFT CANCELLATION FOR UPI ── (Turn 47 feature)
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
      ws.getRange(r._row, prefCol).setValue(refundType);
      console.info(`SUCCESS: Soft-cancelled row ${r._row} with preference ${refundType}`);
      return {
        success: true, 
        message: "Cancellation request received! Admin will verify your payment and process the refund (1-2 days). ✅"
      };
    } else {
      console.error(`FAILED: Missing columns for soft-cancel. StatusCol:${statusCol}, PrefCol:${prefCol}`);
    }
  }

  ws.deleteRow(r._row);

  // Also remove from customer ledger if it exists (10-day billing)
  try {
    const custWs = ss.getSheetByName("SK_Customers");
    if (custWs) {
      const custRows = getAllRows(custWs);
      const cust = custRows.find(c => String(c.Phone).trim() === String(phone).trim());
      if (cust && cust.Ledger_Sheet_ID) {
        const ledger = SpreadsheetApp.openById(cust.Ledger_Sheet_ID);
        ledger.getSheets().forEach(function(tab) {
          const data = tab.getDataRange().getValues();
          for (var i = data.length - 1; i >= 0; i--) {
            if (String(data[i][0]) === String(rowId)) { tab.deleteRow(i + 1); break; }
          }
        });
      }
    }
  } catch(e) { /* ledger cleanup is non-fatal */ }

  return {success: true, message: msg};
}



// ── ADMIN: GET ALL DATA ──────────────────────────────────────
function getAdminData() {
  const ss = getSpreadsheet();

  const bfWs   = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const sabjiWs= getOrCreateTab(ss, TAB_SABJI,     []);
  const menuWs = getOrCreateTab(ss, TAB_MENU,       []);

  const bfRows    = getAllRows(bfWs);
  const sabjiRows = getAllRows(sabjiWs);
  const menuRows  = getAllRows(menuWs);

  const breakfastMaster = bfRows.map(r => ({
    id: String(r.ID), name: String(r.Name), price: Number(r.Price),
    default_on: r.Active === true || String(r.Active).toUpperCase() === "TRUE"
  }));

  const sabjiMaster = sabjiRows.map(r => ({
    id: String(r.ID), name: String(r.Name), type: String(r.Type),
    active: String(r.Active).toLowerCase() !== "false"
  }));

  const menuEntries = menuRows.map(r => {
    const d = r.Date instanceof Date
      ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Date).trim();
    const co = {};
    if (r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
    if (r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
    if (r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);
    let breakfast = [];
    try { if (r.Breakfast_JSON) breakfast = JSON.parse(r.Breakfast_JSON); } catch(e) {}
    return {
      date:             d,
      breakfast:        breakfast,
      lunch_dry:        r.Lunch_Dry    || "",
      lunch_curry:      r.Lunch_Curry  || "",
      dinner_dry:       r.Dinner_Dry   || "",
      dinner_curry:     r.Dinner_Curry || "",
      cutoff_overrides: co,
    };
  });

  return {breakfastMaster, sabjiMaster, menuEntries};
}

// ── ADMIN: SAVE MENU ─────────────────────────────────────────
function saveMenu(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_MENU, []);
  const rows = getAllRows(ws);
  const hIdx = headerIndex(ws);

  const dateStr     = body.date;
  const existing    = rows.find(r => {
    const d = r.Date instanceof Date
      ? Utilities.formatDate(r.Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Date).trim();
    return d === dateStr;
  });

  // breakfast comes as array from admin, serialise to JSON string for storage
  const bfJson = body.breakfast
    ? JSON.stringify(body.breakfast)
    : (body.breakfastJson || "");

  const newRow = [
    dateStr,
    bfJson,
    body.lunch_dry        || body.lunchDry    || "",
    body.lunch_curry      || body.lunchCurry  || "",
    body.dinner_dry       || body.dinnerDry   || "",
    body.dinner_curry     || body.dinnerCurry || "",
    body.cutoff_breakfast || body.cutoffBf    || "",
    body.cutoff_lunch     || body.cutoffL     || "",
    body.cutoff_dinner    || body.cutoffD     || "",
  ];

  if (existing) {
    ws.getRange(existing._row, 1, 1, newRow.length).setValues([newRow]);
  } else {
    ws.appendRow(newRow);
  }
  return {success: true, action: existing ? "updated" : "saved"};
}

// ── ADMIN: BREAKFAST MASTER CRUD ─────────────────────────────
function saveBreakfastItem(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const rows = getAllRows(ws);

  // Admin sends "default_on" (true/false); map to Active column
  const isActive = body.default_on !== false && body.default_on !== "false";

  if (body.id) {
    const r = rows.find(x => String(x.ID) === String(body.id));
    if (r) {
      const hIdx = headerIndex(ws);
      ws.getRange(r._row, hIdx["Name"]).setValue(body.name);
      ws.getRange(r._row, hIdx["Price"]).setValue(body.price);
      ws.getRange(r._row, hIdx["Active"]).setValue(isActive ? "true" : "false");
      return {success: true};
    }
  }
  const newId = "BF-" + new Date().getTime();
  ws.appendRow([newId, body.name, body.price, isActive ? "true" : "false"]);
  return {success: true, id: newId};
}

function deleteBreakfastItem(id) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_BF_MASTER, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => String(x.ID) === String(id));
  if (!r) return {success: false, error: "Not found"};
  ws.deleteRow(r._row);
  return {success: true};
}

// ── ADMIN: SABJI MASTER CRUD ──────────────────────────────────
function saveSabjiItem(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_SABJI, []);
  const rows = getAllRows(ws);

  if (body.id) {
    const r = rows.find(x => String(x.ID) === String(body.id));
    if (r) {
      const hIdx = headerIndex(ws);
      ws.getRange(r._row, hIdx["Name"]).setValue(body.name);
      ws.getRange(r._row, hIdx["Type"]).setValue(body.type);
      ws.getRange(r._row, hIdx["Active"]).setValue(body.active !== false ? "true" : "false");
      return {success: true};
    }
  }
  const newId = "SB-" + new Date().getTime();
  ws.appendRow([newId, body.name, body.type || "Dry", "true"]);
  return {success: true, id: newId};
}

function deleteSabjiItem(id) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_SABJI, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => String(x.ID) === String(id));
  if (!r) return {success: false, error: "Not found"};
  ws.deleteRow(r._row);
  return {success: true};
}

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

// ── AREAS ────────────────────────────────────────────────────

const AREAS_HEADERS = ["Area_Name", "Area_Label", "Free_Delivery"];

const DEFAULT_AREAS = [
  ["Bhosale Garden",  "🏠 Bhosale Garden (Free Delivery)",            "TRUE"],
  ["Triveni Nagar",   "📍 Triveni Nagar (Free Delivery)",             "TRUE"],
  ["Magarpatta",      "🏙️ Magarpatta",                               "FALSE"],
  ["Amanora",         "🏢 Amanora Town",                              "FALSE"],
  ["DP Road",         "🛣️ DP Road",                                   "FALSE"],
  ["Malwadi",         "📍 Malwadi",                                   "FALSE"],
  ["SadeSatraNali",   "📍 SadeSatraNali",                             "FALSE"],
  ["Kirtane Baug",    "📍 Kirtane Baug",                              "FALSE"],
  ["Tupe Patil Road", "🛣️ Tupe Patil Road",                          "FALSE"],
  ["BG Shirke Road",  "🏢 BG Shirke Road",                            "FALSE"],
  ["Vaiduwadi",       "📍 Vaiduwadi (Till Yash Honda Only)",          "FALSE"],
  ["Solapur Road",    "🛣️ Pune-Solapur Road (Till Gadital Only)",     "FALSE"],
  ["Vihar Chowk",     "📍 Vihar Chowk",                               "FALSE"],
  ["Mandai",          "📍 Hadapsar Mandai",                           "FALSE"],
  ["Gadital",         "📍 Gadital",                                   "FALSE"],
  ["Pickup",          "📦 Self Pickup (Waives all fees)",             "TRUE"]
];

function getAreas() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
  const rows = getAllRows(ws);
  // Seed defaults on first run
  if (rows.length === 0) {
    DEFAULT_AREAS.forEach(function(r) { ws.appendRow(r); });
    return DEFAULT_AREAS.map(function(r) { return {name:r[0], label:r[1], free:true}; });
  }
  return rows.map(function(r) {
    return {name: r.Area_Name, label: r.Area_Label, free: r.Free_Delivery === true || String(r.Free_Delivery).toUpperCase() === "TRUE"};
  });
}

function saveArea(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
  const data = ws.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf("Area_Name");
  const labelIdx = headers.indexOf("Area_Label");
  const freeIdx = headers.indexOf("Free_Delivery");
  // Update if exists
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]).toLowerCase() === String(body.name).toLowerCase()) {
      data[i][labelIdx] = body.label;
      data[i][freeIdx]  = body.free ? "TRUE" : "FALSE";
      ws.getDataRange().setValues(data);
      return {success: true};
    }
  }
  // Add new
  ws.appendRow([body.name, body.label, body.free ? "TRUE" : "FALSE"]);
  return {success: true};
}

function deleteArea(body) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_AREAS, AREAS_HEADERS);
  const data = ws.getDataRange().getValues();
  const nameIdx = data[0].indexOf("Area_Name");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]) === String(body.name)) {
      ws.deleteRow(i + 1);
      return {success: true};
    }
  }
  return {success: false, error: "Area not found"};
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

      ws.getRange(i + 1, statusIdx + 1).setValue("Refunded (" + now + ")");
      return {success: true};
    }
  }
  return {success: false, error: "Refund request not found"};
}

// ── ROTI PACKING UTILITY ──────────────────────────────────────
function calculatePackets(total, max) {
  if (total <= 0) return [];
  if (total <= max) return [total];
  var numPacks = Math.ceil(total / max);
  var baseSize = Math.floor(total / numPacks);
  var remainder = total % numPacks;
  var packs = [];
  for (var i = 0; i < numPacks; i++) {
    packs.push(i < remainder ? baseSize + 1 : baseSize);
  }
  return packs;
}

// ── KITCHEN SUMMARY ──────────────────────────────────────────
function getKitchenSummary(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date;
  });

  var meals = {};
  var ROTI_COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri"];
  var ROTI_LIMITS = {
    "Chapati":6, "Without_Oil_Chapati":6,
    "Phulka":12, "Ghee_Phulka":12,
    "Jowar_Bhakri":2, "Bajra_Bhakri":2
  };
  var menu = getMenu(date);

  var orders = [];

  dayRows.forEach(function(r) {
    var meal = String(r.Meal_Type || "");
    if (!meal) return;
    if (!meals[meal]) meals[meal] = {count: 0};
    var m = meals[meal];
    m.count++;

    // For Labels Tab
    var summaryParts = [];
    if (meal === "Breakfast") {
      if (!m.items) m.items = {};
      for (var n = 1; n <= 4; n++) {
        var item = String(r["BF_Item_"+n] || "").trim();
        var qty  = Number(r["BF_Qty_"+n]) || 0;
        if (item && qty > 0) {
          m.items[item] = (m.items[item] || 0) + qty;
          summaryParts.push(qty + " " + item);
        }
      }
      var curdBf = Number(r.Curd) || 0;
      if (curdBf > 0) {
        m.items["Curd"] = (m.items["Curd"] || 0) + curdBf;
        summaryParts.push(curdBf + " Curd");
      }
    } else {
      if (!m.rotis) m.rotis = {};
      if (!m.rotiMatrix) {
        m.rotiMatrix = {};
        ROTI_COLS.forEach(function(c) { m.rotiMatrix[c] = {}; });
      }
      ROTI_COLS.forEach(function(c) {
        var q = Number(r[c]) || 0;
        if (q > 0) {
          m.rotis[c] = (m.rotis[c] || 0) + q;
          summaryParts.push(q + " " + c.replace(/_/g, " "));
          var packs = calculatePackets(q, ROTI_LIMITS[c]);
          packs.forEach(function(p) {
            m.rotiMatrix[c][p] = (m.rotiMatrix[c][p] || 0) + 1;
          });
        }
      });
      if (!m.sabji) {
        m.sabji = {
          dry_kg: 0, curry_kg: 0,
          dry_name: (meal === "Lunch" ? menu.lunch_dry : menu.dinner_dry) || "Sabji (Dry)",
          curry_name: (meal === "Lunch" ? menu.lunch_curry : menu.dinner_curry) || "Sabji (Curry)",
          dry_mini: 0, dry_full: 0, curry_mini: 0, curry_full: 0
        };
      }
      var dMini = Number(r.Dry_Sabji_Mini)||0;
      var dFull = Number(r.Dry_Sabji_Full)||0;
      var cMini = Number(r.Curry_Sabji_Mini)||0;
      var cFull = Number(r.Curry_Sabji_Full)||0;
      m.sabji.dry_mini  += dMini;
      m.sabji.dry_full  += dFull;
      m.sabji.curry_mini += cMini;
      m.sabji.curry_full += cFull;
      
      if (dMini > 0) summaryParts.push(dMini + " Mini Dry");
      if (dFull > 0) summaryParts.push(dFull + " Full Dry");
      if (cMini > 0) summaryParts.push(cMini + " Mini Curry");
      if (cFull > 0) summaryParts.push(cFull + " Full Curry");

      if (!m.other) m.other = {Dal:{kg:0, count:0}, Rice:{count:0}, Salad:{count:0}, Curd:{count:0}};
      var dalQ = Number(r.Dal)   || 0;
      var riceQ = Number(r.Rice)  || 0;
      var saladQ = Number(r.Salad) || 0;
      var curdQ = Number(r.Curd)  || 0;
      
      m.other.Dal.kg      += dalQ * 1.33;
      m.other.Dal.count   += dalQ;
      m.other.Rice.count  += riceQ;
      m.other.Salad.count += saladQ;
      m.other.Curd.count  += curdQ;

      if (dalQ > 0) summaryParts.push(dalQ + " Dal");
      if (riceQ > 0) summaryParts.push(riceQ + " Rice");
      if (saladQ > 0) summaryParts.push(saladQ + " Salad");
      if (curdQ > 0) summaryParts.push(curdQ + " Curd");
    }

    orders.push({
      Submission_ID: String(r.Submission_ID || ""),
      Customer_Name: String(r.Customer_Name || ""),
      Meal_Type: meal,
      summary: summaryParts.join(", "),
      Special_Notes_Kitchen: String(r.Special_Notes_Kitchen || ""),
      marathiNotes: String(r.marathiNotes || ""),
      Packed: r.Packed === true || String(r.Packed).toLowerCase() === "true"
    });
  });

  ["Lunch","Dinner"].forEach(function(meal) {
    if (!meals[meal]) return;
    var m = meals[meal];
    if (m.other && m.other.Dal) m.other.Dal.kg = Math.round(m.other.Dal.kg * 100) / 100;
  });

  return {date: date, meals: meals, orders: orders};
}

// ── DRIVER ORDERS ─────────────────────────────────────────────
function getDriverOrders(date) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);
  var meals = {Breakfast: [], Lunch: [], Dinner: []};

  // Load delivery status from SK_Deliveries tab
  var delMap = {};
  var delWs  = ss.getSheetByName("SK_Deliveries");
  if (delWs) {
    getAllRows(delWs).forEach(function(r) {
      var sid = String(r.Submission_ID || "").trim();
      if (sid) delMap[sid] = String(r.Delivered_At || "");
    });
  }

  rows.forEach(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    if (d !== date) return;
    if (String(r.Payment_Status) === "Cancelled") return;
    var area = String(r.Area || "").trim();
    if (area.toLowerCase() === "pickup") return;
    var meal = String(r.Meal_Type || "");
    if (!meals[meal]) return;
    var sid = String(r.Submission_ID || "");
    meals[meal].push({
      submissionId:  sid,
      name:          String(r.Customer_Name || ""),
      phone:         String(r.Phone || ""),
      area:          area,
      address:       String(r.Full_Address || ""),
      landmark:      String(r.Landmark || ""),
      maps:          String(r.Maps_Link || ""),
      notes:         String(r.Special_Notes_Delivery || ""),
      deliveredAt:   delMap[sid] || "",
      amount:        Number(r.Net_Total || r.Food_Subtotal || 0),
      paymentStatus: String(r.Payment_Status || "")
    });
  });

  // ── Resolve shortened Maps URLs server-side (goo.gl, maps.app.goo.gl) ────────
  // Collect orders whose URL has no extractable coords
  var allOrders = [].concat(meals.Breakfast, meals.Lunch, meals.Dinner);
  var toResolve = [];
  allOrders.forEach(function(o) {
    if (o.maps && !_extractCoordsGS(o.maps) &&
        (o.maps.indexOf('goo.gl') > -1 || o.maps.indexOf('maps.app') > -1)) {
      toResolve.push(o);
    }
  });
  if (toResolve.length > 0) {
    try {
      var reqs = toResolve.map(function(o) {
        return { url: o.maps, followRedirects: false, muteHttpExceptions: true };
      });
      var resps = UrlFetchApp.fetchAll(reqs);
      resps.forEach(function(resp, i) {
        var code = resp.getResponseCode();
        if (code >= 300 && code < 400) {
          var hdrs = resp.getHeaders();
          var loc  = hdrs['Location'] || hdrs['location'] || '';
          if (loc) {
            var c = _extractCoordsGS(loc);
            if (c) { toResolve[i].lat = c.lat; toResolve[i].lng = c.lng; }
          }
        }
      });
    } catch(e) { /* silently fail — route optimisation just skips those stops */ }
  }

  return {date: date, meals: meals};
}

// Shared coord extractor for Apps Script (mirrors client-side regex)
function _extractCoordsGS(url) {
  if (!url) return null;
  var m = url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: +m[1], lng: +m[2] };
  m = url.match(/[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: +m[1], lng: +m[2] };
  return null;
}

// ── ORDER SUMMARY ────────────────────────────────────────────
function getOrderSummary(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date && String(r.Payment_Status) !== "Cancelled";
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
    if (payStatus === "Paid" || payStatus === "Wallet Paid") m.paid += net; else m.pending += net;
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
    if (payStatus === "Paid" || payStatus === "Wallet Paid") totals.paid += net; else totals.pending += net;
    if (!customerSet[String(r.Phone)]) { customerSet[String(r.Phone)] = true; totals.customers++; }
  });

  return {date: date, meals: meals, totals: totals};
}

// ── LABEL ORDERS ──────────────────────────────────────────────
function getLabelOrders(date, meal) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);
  var COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
              "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full","Dal","Rice","Salad"];

  var orders = rows
    .filter(function(r) {
      var d = r.Order_Date instanceof Date
        ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
        : String(r.Order_Date).trim();
      return d === date && String(r.Meal_Type) === meal;
    })
    .map(function(r) {
      var obj = {
        name:  String(r.Customer_Name || ""),
        area:  String(r.Area || ""),
        notes: String(r.Special_Notes || ""),
        Curd:  Number(r.Curd) || 0
      };
      if (meal === "Breakfast") {
        for (var n = 1; n <= 4; n++) {
          obj["BF_Item_"+n] = String(r["BF_Item_"+n] || "");
          obj["BF_Qty_"+n]  = Number(r["BF_Qty_"+n])  || 0;
        }
      } else {
        COLS.forEach(function(col) { obj[col] = Number(r[col]) || 0; });
      }
      return obj;
    });

  return {orders: orders};
}

// ── PACKAGING EXPENSES ────────────────────────────────────────
// Edit unit costs below to match your actual supplier prices
var PKG_UNIT_COSTS = {
  "Breakfast Box":           3.00,
  "Delivery Bag":            1.00,
  "Label / Sticker":         0.25,
  "Bread Packet":            0.70,
  "Sabji Container (Mini)":  2.00,
  "Sabji Container (Full)":  2.50,
  "Dal Container":           2.00,
  "Rice Container":          1.50,
  "Salad Container":         1.00,
  "Curd Container":          1.00
};

function getPackagingExpenses(date) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var dayRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d === date;
  });

  if (dayRows.length === 0) return {date: date, orderCount: 0, meals: {}, items: [], total: 0};

  var counts = {};
  var mealCounts = {Breakfast:0, Lunch:0, Dinner:0};
  function add(key, qty) { if (qty > 0) counts[key] = (counts[key]||0) + qty; }

  dayRows.forEach(function(r) {
    var meal = String(r.Meal_Type || "");
    if (mealCounts[meal] !== undefined) mealCounts[meal]++;

    add("Label / Sticker", 1);
    if (meal === "Breakfast") {
      add("Breakfast Box", 1);
      add("Curd Container", Number(r.Curd) || 0);
    } else {
      add("Delivery Bag", 1);
      var breadCols = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri"];
      var hasBread = breadCols.some(function(c) { return (Number(r[c])||0) > 0; });
      if (hasBread) add("Bread Packet", 1);
      add("Sabji Container (Mini)", (Number(r.Dry_Sabji_Mini)||0) + (Number(r.Curry_Sabji_Mini)||0));
      add("Sabji Container (Full)", (Number(r.Dry_Sabji_Full)||0) + (Number(r.Curry_Sabji_Full)||0));
      add("Dal Container",          Number(r.Dal)   || 0);
      add("Rice Container",         Number(r.Rice)  || 0);
      add("Salad Container",        Number(r.Salad) || 0);
      add("Curd Container",         Number(r.Curd)  || 0);
    }
  });

  var itemOrder = ["Breakfast Box","Delivery Bag","Label / Sticker","Bread Packet",
                   "Sabji Container (Mini)","Sabji Container (Full)",
                   "Dal Container","Rice Container","Salad Container","Curd Container"];
  var items = [];
  var total = 0;
  itemOrder.forEach(function(key) {
    var qty = counts[key] || 0;
    if (!qty) return;
    var unitCost = PKG_UNIT_COSTS[key] || 0;
    var t = qty * unitCost;
    items.push({name: key, qty: qty, unitCost: unitCost, total: t});
    total += t;
  });

  var mealsOut = {};
  Object.keys(mealCounts).forEach(function(m) { if (mealCounts[m] > 0) mealsOut[m] = mealCounts[m]; });

  return {date: date, orderCount: dayRows.length, meals: mealsOut, items: items, total: total};
}

// ── LABEL DRIVE SAVE ─────────────────────────────────────────
function saveLabels(body) {
  var date   = body.date;  // "2026-03-18"
  var meal   = body.meal;  // "Lunch"
  var pdfB64 = body.pdf;   // base64-encoded PDF bytes

  var parts      = date.split("-");
  var year       = parts[0];                                        // "2026"
  var monthNum   = parseInt(parts[1], 10);                          // 3
  var monthNames = ["January","February","March","April","May","June",
                    "July","August","September","October","November","December"];
  var monthName  = monthNames[monthNum - 1];                        // "March"
  var mealAbbrev = meal.toLowerCase().substring(0, 7);              // "breakfa" / "lunch" / "dinner"
  var langCode   = (body.lang === "Devanagari") ? "mar" : "eng";    // "eng" / "mar"
  var filename   = "labels_" + mealAbbrev + "_" + langCode + "_" + date + "_58x25.pdf";
  // e.g. "labels_breakfa_eng_2026-03-05_58x25.pdf"

  var folder = getOrCreateFolderPath([
    "Svaadh Kitchen", "Accounting", "Tally Form Daily Sheets",
    "Processed_Orders", "Labels", year, monthName
  ]);

  // Replace existing file to avoid duplicates
  var existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var pdfBlob = Utilities.newBlob(Utilities.base64Decode(pdfB64), "application/pdf", filename);
  var file = folder.createFile(pdfBlob);
  return {url: file.getUrl(), name: filename};
}

function getOrCreateFolderPath(pathParts) {
  var folder = DriveApp.getRootFolder();
  pathParts.forEach(function(name) {
    var iter = folder.getFoldersByName(name);
    folder = iter.hasNext() ? iter.next() : folder.createFolder(name);
  });
  return folder;
}

// ── CHATBOT ──────────────────────────────────────────────────

function handleChat(body) {
  const userMessage = String(body.message || "").trim();
  const history     = body.history || [];   // [{role:"user"|"model", text:"..."}]
  if (!userMessage) return {reply: "Please send a message."};
  return {reply: callGemini(buildSystemPrompt(), history, userMessage)};
}

function buildSystemPrompt() {
  const B = BUSINESS_CONTEXT;
  const breads = B.menu.breads.map(function(i){ return i.name+"₹"+i.price; }).join(", ");
  const sabji  = B.menu.sabji.map(function(i){ return i.name+"₹"+i.price; }).join(", ");
  const basics = B.menu.basics.map(function(i){ return i.name+"₹"+i.price; }).join(", ");

  var todayLine = "";
  try {
    // Use new Date() directly — Apps Script respects the project timezone (Asia/Kolkata)
    // DO NOT manually add 5.5 hours — Utilities.formatDate already applies the timezone
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

  return "You are a helpful assistant for Svaadh Kitchen, a vegetarian cloud kitchen in Hadapsar, Pune."
    +" Closed Sundays. Over 2.5 years of service (since Aug 2023). Cutoffs: BF<7AM, Lunch<9:30AM, Dinner<5PM."
    +" Areas: Bhosale Garden/Triveni Nagar(free), others(₹10/meal if subtotal<₹100), Self Pickup available.\n"
    + todayLine
    +"MEAL MODEL: Make Your Own Meal (not a fixed thali). Customers pick items individually.\n"
    +"Lunch/Dinner — Breads:"+breads+" | Sabji:"+sabji+" | Basics:"+basics+"\n"
    +"Breakfast: daily rotating ₹35–₹70. "+B.menu.breakfast_note+"\n"
    +"Self pickup also available (no delivery charge).\n"
    +"Uses Pure Ghee & Groundnut refined oil.\n"
    +"Discounts(auto): 5% off≥₹300/day, 7.5% off≥₹450/day.\n"
    +"Payment: Wallet (Prepaid) or UPI("+B.payment.upi_id+"), prepaid cycle (requires wallet balance).\n"
    +"Order: "+B.ordering.order_url+" — no login needed, phone=identity, can book multiple days.\n"
    +"WhatsApp: "+B.contact.whatsapp+" | WA group: "+B.contact.whatsapp_group+"\n"
    +"Reply in customer's language(English/Hindi/Marathi). Be brief & warm."
    +" For orders send to order URL. Don't invent info. Direct unknowns to WhatsApp.";
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
     r.Payment_Status === "Pending"         ||
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

  const customers = Object.values(map).map(c => ({...c, total: Math.round(c.total)}));
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
         r.Payment_Status === "Pending"         ||
         !r.Payment_Status)) {
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue("Paid");
      updated++;
    }
  });
  return {success:true, updatedRows:updated, customersMarked:phones.length};
}

// ── GET ORDER HISTORY (date range) ────────────────────────────────────────────
function getOrderHistory(p) {
  var dateFrom = p.dateFrom, dateTo = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var rows = getAllRows(ws);

  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var filtered = rows.filter(function(r) {
    var d = fmtDate(r.Order_Date);
    return d >= dateFrom && d <= dateTo && String(r.Payment_Status) !== "Cancelled";
  });

  var orders = filtered.map(function(r) {
    return {
      id:     r.Submission_ID,
      date:   fmtDate(r.Order_Date),
      meal:   r.Meal_Type,
      name:   r.Customer_Name,
      phone:  r.Phone,
      area:   r.Area,
      total:  Number(r.Net_Total) || 0,
      status: r.Payment_Status || "Pending",
      notes:  r.Special_Notes || ""
    };
  });

  var totalRev     = orders.reduce(function(s,o){return s+o.total;},0);
  var totalPaid    = orders.filter(function(o){return String(o.status)==="Paid" || String(o.status)==="Wallet Paid";}).reduce(function(s,o){return s+o.total;},0);
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

// ── GET CUSTOMER LIST ─────────────────────────────────────────────────────────
function getCustomerList() {
  var ss   = getSpreadsheet();
  var ordersWs = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var ordRows = getAllRows(ordersWs);

  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var custRows = getAllRows(custWs);
  var cMap = {};
  custRows.forEach(function(c) {
    var p = _normalizePhone(c.Phone);
    if (p) {
      cMap[p] = {
        count: Number(c.Review_Promo_Count) || 0,
        claimed: (String(c.Review_Reward_Claimed) === "TRUE" || String(c.Review_Reward_Claimed) === "true")
      };
    }
  });

  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var map = {};
  ordRows.forEach(function(r) {
    if (String(r.Payment_Status) === "Cancelled") return;
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
        reviewClaimed: cMap[normP] ? cMap[normP].claimed : false
      };
    }
    map[phone].orderCount++;
    map[phone].totalSpent += Number(r.Net_Total)||0;
    if (String(r.Payment_Status) !== "Paid" && String(r.Payment_Status) !== "Wallet Paid") map[phone].pendingAmt += Number(r.Net_Total)||0;
    if (d > map[phone].lastDate) {
      map[phone].lastDate = d;
      map[phone].name = String(r.Customer_Name||map[phone].name).trim();
    }
  });

  var customers = Object.values(map)
    .map(function(c){return Object.assign({},c,{totalSpent:Math.round(c.totalSpent),pendingAmt:Math.round(c.pendingAmt)});})
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

  var rows = getAllRows(ws).filter(function(r){return String(r.Phone||"").trim()===phone;});

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
  var activeOrders = orders.filter(function(o){return String(o.status)!=="Cancelled";});
  var totalSpent = Math.round(activeOrders.reduce(function(s,o){return s+o.total;},0));
  var pending    = Math.round(activeOrders.filter(function(o){return String(o.status)!=="Paid" && String(o.status)!=="Wallet Paid";}).reduce(function(s,o){return s+o.total;},0));

  return {success:true, phone:phone, name:name, area:area, payFreq:payFreq,
          orders:orders, totalSpent:totalSpent, pending:pending, orderCount:orders.length};
}

// ── GET DATE PAYMENTS ─────────────────────────────────────────────────────────
function getDatePayments(date) {
  if (!date) return {success:false, error:"date required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var rows = getAllRows(ws).filter(function(r){return fmtDate(r.Order_Date)===date && String(r.Payment_Status)!=="Cancelled";});

  var map = {};
  rows.forEach(function(r) {
    var phone = String(r.Phone||"").trim();
    if (!phone) return;
    if (!map[phone]) map[phone] = {phone:phone, name:String(r.Customer_Name||"").trim(),
      payFreq:String(r.Payment_Freq||"").trim(), meals:[], total:0, allPaid:true};
    map[phone].meals.push(r.Meal_Type);
    map[phone].total += Number(r.Net_Total)||0;
    if (String(r.Payment_Status)!=="Paid" && String(r.Payment_Status)!=="Wallet Paid") map[phone].allPaid = false;
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
    if (currentStatus === "Cancelled (Verify UPI)") {
      // ── Process Refund Logic based on preference
      const pref = String(r.Refund_Preference || "upi").toLowerCase();
      const amt = Number(r.Net_Total) || 0;
      const custName = r.Customer_Name || "Customer";
      
      if (pref === "wallet" && amt > 0) {
        _appendWalletTransaction(phone, custName, "Order Cancellation Refund", amt, true, String(r.Submission_ID));
      } else if (pref === "manual_upi" && amt > 0) {
        const refWs = getOrCreateTab(ss, TAB_REFUNDS, ["Submission_ID","Phone","Name","Amount","Meal","Date","Status","Timestamp","Adjustment_Note","Refund_Mode"]);
        refWs.appendRow([r.Submission_ID, phone, custName, amt, r.Meal_Type, date, "Pending", now, "Verified Soft Cancellation", "upi"]);
      }
      ws.deleteRow(r._row); // Final delete after verification
    } else {
      // ── Standard Payment Approval
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue(status);
    }
    updated++;
  });

  return {success:true, updatedRows:updated};
}

// ── DELETED OBSOLETE ADMIN CANCEL ORDER (Merged with main) ──
// ── ANALYTICS ─────────────────────────────────────────────────────────────────
function getAnalytics(p) {
  var dateFrom = p.dateFrom, dateTo = p.dateTo;
  if (!dateFrom || !dateTo) return {success:false, error:"dateFrom and dateTo required"};
  var ss  = getSpreadsheet();
  var ws  = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };
  var rows = getAllRows(ws).filter(function(r) {
    var d = fmtDate(r.Order_Date);
    return d >= dateFrom && d <= dateTo && String(r.Payment_Status) !== "Cancelled";
  });
  var LUNCH_COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
    "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full","Dal","Rice","Salad","Curd"];
  var COL_DISP = {"Chapati":"Chapati","Without_Oil_Chapati":"WO Chapati","Phulka":"Phulka","Ghee_Phulka":"Ghee Phulka",
    "Jowar_Bhakri":"Jowar Bhakri","Bajra_Bhakri":"Bajra Bhakri","Dry_Sabji_Mini":"Dry Sabji Mini",
    "Dry_Sabji_Full":"Dry Sabji Full","Curry_Sabji_Mini":"Curry Sabji Mini","Curry_Sabji_Full":"Curry Sabji Full",
    "Dal":"Dal","Rice":"Rice","Salad":"Salad","Curd":"Curd"};
  var totalRev=0, totalPaid=0, custSet={}, dayMap={};
  var mealStats={Breakfast:{count:0,revenue:0},Lunch:{count:0,revenue:0},Dinner:{count:0,revenue:0}};
  var itemCounts={};
  rows.forEach(function(r) {
    var d=fmtDate(r.Order_Date), net=Number(r.Net_Total)||0;
    totalRev+=net; if(String(r.Payment_Status)==="Paid" || String(r.Payment_Status)==="Wallet Paid") totalPaid+=net;
    var ph=String(r.Phone||"").trim(); if(ph) custSet[ph]=true;
    var meal=String(r.Meal_Type||"");
    if(mealStats[meal]){mealStats[meal].count++;mealStats[meal].revenue+=net;}
    if(!dayMap[d])dayMap[d]={orders:0,revenue:0};
    dayMap[d].orders++;dayMap[d].revenue+=net;
    if(meal==="Breakfast"){
      for(var n=1;n<=4;n++){var bi=String(r["BF_Item_"+n]||"").trim(),bq=Number(r["BF_Qty_"+n])||0;if(bi&&bq>0)itemCounts[bi]=(itemCounts[bi]||0)+bq;}
      var cu=Number(r.Curd)||0; if(cu>0)itemCounts["Curd"]=(itemCounts["Curd"]||0)+cu;
    } else {
      LUNCH_COLS.forEach(function(col){var q=Number(r[col])||0;if(q>0){var dn=COL_DISP[col]||col;itemCounts[dn]=(itemCounts[dn]||0)+q;}});
    }
  });
  var days=Object.keys(dayMap).sort().map(function(d){return{date:d,orders:dayMap[d].orders,revenue:Math.round(dayMap[d].revenue)};});
  var topItems=Object.keys(itemCounts).map(function(k){return{name:k,count:Math.round(itemCounts[k])};}).sort(function(a,b){return b.count-a.count;}).slice(0,15);
  Object.keys(mealStats).forEach(function(m){mealStats[m].revenue=Math.round(mealStats[m].revenue);});
  return {success:true,
    summary:{orders:rows.length,customers:Object.keys(custSet).length,revenue:Math.round(totalRev),
      paid:Math.round(totalPaid),pending:Math.round(totalRev-totalPaid),
      avgPerDay:days.length>0?Math.round(totalRev/days.length):0},
    meals:mealStats,days:days,topItems:topItems};
}

// ── CHURN REPORT ──────────────────────────────────────────────────────────────
function getChurnReport(sinceDate) {
  if (!sinceDate) return {success:false, error:"sinceDate required"};
  var ss=getSpreadsheet(), ws=getOrCreateTab(ss,TAB_ORDERS,ORDERS_HEADERS);
  var fmtDate=function(v){return v instanceof Date?Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd"):String(v).trim();};
  var map={};
  getAllRows(ws).forEach(function(r){
    if(String(r.Payment_Status)==="Cancelled")return;
    var phone=String(r.Phone||"").trim(); if(!phone)return;
    var d=fmtDate(r.Order_Date);
    if(!map[phone])map[phone]={phone:phone,name:String(r.Customer_Name||"").trim(),area:String(r.Area||"").trim(),lastDate:"",orderCount:0};
    map[phone].orderCount++;
    if(d>map[phone].lastDate){map[phone].lastDate=d;map[phone].name=String(r.Customer_Name||map[phone].name).trim();}
  });
  var churned=Object.values(map).filter(function(c){return c.lastDate<sinceDate;}).sort(function(a,b){return b.lastDate.localeCompare(a.lastDate);});
  return {success:true,sinceDate:sinceDate,customers:churned,count:churned.length};
}


// ── LIVE TRACKER LOGIC: En Route & Delivered ──────────────────────────────────
function batchMarkEnRoute(body) {
  var sids      = body.submissionIds;
  var enRouteAt = body.enRouteAt;
  if (!sids || !sids.length) return {success:false, error:"submissionIds required"};

  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  var data = ws.getDataRange().getValues();
  var headers = data[0];
  var erIdx   = headers.indexOf("EnRoute_At");
  var sidIdx  = headers.indexOf("Submission_ID");

  if (erIdx === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
    erIdx = headers.length;
  }

  // Create lookup for existing rows
  var sidToRowMap = {};
  for (var i = 1; i < data.length; i++) {
    sidToRowMap[String(data[i][sidIdx])] = i + 1;
  }

  sids.forEach(function(sid) {
    var row = sidToRowMap[String(sid)];
    if (row) {
      ws.getRange(row, erIdx + 1).setValue(enRouteAt);
    } else {
      // Append if not found (though usually we expect them to be found if already rendered)
      ws.appendRow([sid, "", enRouteAt]);
    }
  });

  return {success:true};
}

function markEnRoute(body) {
  var sid         = body.submissionId;
  var enRouteAt   = body.enRouteAt;
  if (!sid) return {success:false, error:"submissionId required"};

  var ss  = getSpreadsheet();
  var ws  = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  var rows = getAllRows(ws);
  
  // Ensure "EnRoute_At" header exists
  var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  if (headers.indexOf("EnRoute_At") === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
  }
  var hIdx = headerIndex(ws);

  var existing = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Submission_ID) === String(sid)) { existing = rows[i]; break; }
  }
  if (existing) {
    ws.getRange(existing._row, hIdx["EnRoute_At"]).setValue(enRouteAt);
  } else {
    // Note: Append Row pushes based on sheet width. Best to be safe using indices.
    var newRowArr = [];
    newRowArr[hIdx["Submission_ID"] - 1] = sid;
    newRowArr[hIdx["Delivered_At"] - 1] = "";
    newRowArr[hIdx["EnRoute_At"] - 1] = enRouteAt;
    ws.appendRow(newRowArr);
  }
  return {success:true, submissionId:sid, enRouteAt:enRouteAt};
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

/**
 * ADMIN: Fetch unverified wallet recharges
 */
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

/**
 * ADMIN: Approve a wallet recharge
 */
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
      return {success:true};
    }
  }
  
  Logger.log(`Activation Failed: No match for Phone: ${phone}, TS: ${ts}. Rows scanned: ${rows.length - 1}`);
  return {success:false, error:"Recharge request not found or already verified"};
}

/**
 * ADMIN: Fetch all orders with "Pending" status (usually UPI)
 */
function getPendingUPIPayments() {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  // Return Payment_Status == "Pending" OR "Cancelled (Verify UPI)"
  return rows.filter(r => {
    const s = String(r.Payment_Status).trim();
    return s === "Pending" || s === "Cancelled (Verify UPI)";
  })
             .map(r => ({
               id: r.Submission_ID,
               date: r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd") : r.Order_Date,
               customer: r.Customer_Name,
               phone: r.Phone,
               amount: r.Net_Total,
               meal: r.Meal_Type,
               timestamp: r.Submitted_At,
               status: r.Payment_Status,
               refund_preference: r.Refund_Preference || ""
             }));
}

// ── MARK DELIVERED ────────────────────────────────────────────────────────────
function markDelivered(body) {
  var sid         = body.submissionId;
  var deliveredAt = body.deliveredAt;
  if (!sid) return {success:false, error:"submissionId required"};

  var ss  = getSpreadsheet();
  // Ensure we define the tab properly
  var ws  = getOrCreateTab(ss, "SK_Deliveries", ["Submission_ID","Delivered_At","EnRoute_At"]);
  
  var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  if (headers.indexOf("EnRoute_At") === -1) {
    ws.getRange(1, headers.length + 1).setValue("EnRoute_At");
  }
  
  var rows = getAllRows(ws);
  var hIdx = headerIndex(ws);

  var existing = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Submission_ID) === sid) { existing = rows[i]; break; }
  }
  if (existing) {
    ws.getRange(existing._row, hIdx["Delivered_At"]).setValue(deliveredAt);
  } else {
    var newRowArr = [];
    newRowArr[hIdx["Submission_ID"] - 1] = sid;
    newRowArr[hIdx["Delivered_At"] - 1] = deliveredAt;
    newRowArr[hIdx["EnRoute_At"] - 1] = "";
    ws.appendRow(newRowArr);
  }
  return {success:true, submissionId:sid, deliveredAt:deliveredAt};
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
    
    return rPhone === phone && rMeal === meal && orderDate === dateStr && status !== 'deleted' && status !== 'cancelled';
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

/**
 * Run this function once from the Apps Script editor to populate
 * dummy orders for Today and Tomorrow for testing prints/labels.
 */
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
      { date: tStr, meal: "Lunch", name: "Amit Kulkarni", area: "Bhosale Garden", society: "Laxmi Vihar", wing: "A", flat: "104", items: {"Chapati": 3, "Dry_Sabji_Mini": 1, "Dal": 1}, total: 71, notes: "Deliver at gate" },
      { date: tStr, meal: "Lunch", name: "Sneha Patil", area: "Magarpatta", society: "Cosmos", wing: "E", flat: "P-5", items: {"Phulka": 2, "Curry_Sabji_Full": 1, "Rice": 1}, total: 77, notes: "" },
      { date: tStr, meal: "Dinner", name: "Mayur Joshi", area: "DP Road", society: "Riverview", wing: "F", flat: "901", items: {"Jowar_Bhakri": 2, "Curry_Sabji_Mini": 1}, total: 62, notes: "Ring bell and leave" },
      { date: mStr, meal: "Breakfast", name: "Priya Rao", area: "Magarpatta", society: "Pentagon 3", wing: "A", flat: "610", items: {"Sabudana Khichdi": 1}, total: 40, notes: "" },
      { date: mStr, meal: "Lunch", name: "Vikram Shah", area: "Amanora", society: "Adreno", wing: "1", flat: "1502", items: {"Ghee_Phulka": 4, "Dry_Sabji_Full": 1, "Salad": 1}, total: 100, notes: "Call on arrival" },
      { date: mStr, meal: "Dinner", name: "Svaadh Test", area: "Bhosale Garden", society: "Self Pickup", wing: "-", flat: "-", items: {"Chapati": 2, "Dry_Sabji_Mini": 1, "Dal": 1}, total: 62, notes: "I will pick up" }
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
