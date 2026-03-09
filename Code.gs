// ============================================================
// SVAADH KITCHEN — Code.gs (New System)
// One Google Sheet, clean schema, no Tally dependency
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SHEET_ID       = "17X7JOrMe1Oj_QykH7mk6UGoBjuguLfyC1RmKYATDXlI";   // ← Create a new blank Google Sheet and paste its ID here
const ADMIN_PIN      = "7284";                       // ← Updated before go-live
const CODE_VERSION   = 2;
const LEDGER_FOLDER  = "Svaadh Customer Ledgers";

// Sheet tab names
const TAB_ORDERS     = "SK_Orders";
const TAB_CUSTOMERS  = "SK_Customers";
const TAB_MENU       = "SK_Daily_Menu";
const TAB_BF_MASTER  = "SK_Master_Breakfast";
const TAB_SABJI      = "SK_Master_Sabjis";

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

const ORDERS_HEADERS = [
  "Submission_ID","Submitted_At","Order_Date","Meal_Type",
  "Customer_Name","Phone","Area","Wing","Flat","Floor","Society","Full_Address","Maps_Link","Landmark",
  "Items_JSON",
  "Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
  "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full",
  "Dal","Rice","Salad","Curd",
  "BF_Item_1","BF_Qty_1","BF_Item_2","BF_Qty_2","BF_Item_3","BF_Qty_3","BF_Item_4","BF_Qty_4",
  "Special_Notes",
  "Food_Subtotal","Delivery_Charge","Discount_Amount","Net_Total",
  "Payment_Method","Payment_Status","Payment_Freq","First_Time","Source"
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
  about: "Svaadh Kitchen is a home-based vegetarian cloud kitchen in Hadapsar, Pune, serving fresh and wholesome homemade meals for over 2.5 years. We specialize in homemade vegetarian food, offering breakfast, lunch, and dinner with a changing daily sabji menu. We serve Bhosale Garden (free delivery), Magarpatta, Amanora Township, and DP Road areas. Our meals are cooked fresh daily using quality ingredients and traditional recipes.",
  vision: "To make homemade vegetarian meals easily accessible and affordable for everyone, while maintaining taste, quality, and consistency.",
  locations_served: ["Bhosale Garden (free delivery)", "Magarpatta", "Amanora Township", "DP Road"],
  order_cutoffs: { breakfast: "before 7:00 AM", lunch: "before 9:30 AM", dinner: "before 5:00 PM", closed_on: "Sunday" },
  delivery: {
    free_area: "Bhosale Garden",
    charge: "₹10 per meal for Magarpatta/Amanora/DP Road, only if that meal's subtotal is below ₹100. Free for Bhosale Garden always.",
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
    breakfast: "Rotating daily — items like Kanda Poha, Aloo Paratha, Paneer Paratha with Curd. Check the order form for today's options."
  },
  discounts: {
    tier1: "5% off when the day total is ₹300 or more",
    tier2: "7.5% off when the day total is ₹450 or more",
    note: "Discounts are applied automatically per day's total when placing an order."
  },
  payment: {
    options: ["UPI", "Cash on Delivery (COD)", "10-Day post-paid cycle"],
    upi_id: "shriharioze07-1@okhdfcbank",
    ten_day: "Amount accumulates over 10 days (1–10, 11–20, 21–end of month), settled once at period end."
  },
  ordering: {
    order_url: "https://www.svaadhkitchen.in/order.html",
    process: "Open the order form → enter phone number → fill address → pick dates → choose meals → review bill → pay via UPI or COD.",
    advance: "Select multiple dates on the calendar to order for the full week in one go.",
    edit_cancel: "Use 'View/Edit existing orders' on the order form home screen to edit or cancel before the cutoff.",
    no_login: "No login needed — phone number is your identity. Details are saved automatically."
  },
  contact: {
    phone_primary: "9930748908",
    phone_alt: "9819969682",
    whatsapp: "+91 99307 48908",
    whatsapp_link: "https://wa.me/919930748908",
    whatsapp_group: "https://chat.whatsapp.com/EpLv7mtYipm61ScKjbOiuk",
    email: "svaadh.kitchen@gmail.com",
    google_page: "https://share.google/UnZM2xcLOF2QVO9cj"
  }
};

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  const p = e.parameter;
  const action = p.action || "";
  try {
    if (action === "getCustomer")   return jsonRes(getCustomer(p.phone));
    if (action === "getMenu")       return jsonRes(getMenu(p.date));
    if (action === "getCustomerOrders") return jsonRes(getCustomerOrders(p.phone));
    if (action === "get10DayRunning")   return jsonRes(get10DayRunning(p.phone));
    if (action === "getAdminData") {
      if (p.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(getAdminData());
    }
    if (action === "getUnpaidCustomers") {
      if (p.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(getUnpaidCustomers(p));
    }
    return jsonRes({error:"Unknown action"});
  } catch(err) {
    return jsonRes({error: err.message});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body._action || "";
    if (action === "deleteOrder")       return jsonRes(deleteOrder(body.phone, body.rowId));
    if (action === "deleteBreakfastItem") {
      if (body.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(deleteBreakfastItem(body.id));
    }
    if (action === "saveBreakfastItem") {
      if (body.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(saveBreakfastItem(body));
    }
    if (action === "deleteSabjiItem") {
      if (body.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(deleteSabjiItem(body.id));
    }
    if (action === "saveSabjiItem") {
      if (body.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(saveSabjiItem(body));
    }
    if (action === "chat")            return jsonRes(handleChat(body));
    if (action === "markCustomersPaid") {
      if (body.pin !== ADMIN_PIN) return jsonRes({error:"Invalid PIN"});
      return jsonRes(markCustomersPaid(body));
    }
    if (body.pin === ADMIN_PIN)       return jsonRes(saveMenu(body));
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
    ws.appendRow(headers);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#c0392b")
      .setFontColor("white");
  }
  return ws;
}

function getISTTimestamp() {
  const now = new Date();
  const ist = new Date(now.getTime() + 5.5 * 3600 * 1000);
  return Utilities.formatDate(ist, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
}

function generateSubmissionID() {
  const ist = new Date(new Date().getTime() + 5.5 * 3600 * 1000);
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
  getOrCreateTab(ss, TAB_CUSTOMERS, [
    "Phone","Name","Area","Wing","Flat","Floor","Society","Full_Address",
    "Maps_Link","Landmark","Payment_Freq","Created_At","Ledger_Sheet_ID"
  ]);
  getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner"
  ]);
  getOrCreateTab(ss, TAB_BF_MASTER, ["ID","Name","Price","Active"]);
  getOrCreateTab(ss, TAB_SABJI,     ["ID","Name","Type","Active"]);
  return {success: true, message: "Schema initialised"};
}

// ── GET CUSTOMER ─────────────────────────────────────────────
function getCustomer(phone) {
  if (!phone) return {found: false};
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, []);
  const rows = getAllRows(ws);
  const r = rows.find(x => String(x.Phone).trim() === String(phone).trim());
  if (!r) return {found: false};
  return {
    found: true,
    name:               r.Name || "",
    area:               r.Area || "",
    wing:               r.Wing || "",
    flat:               r.Flat || "",
    floor:              r.Floor || "",
    society:            r.Society || "",
    maps:               r.Maps_Link || "",
    landmark:           r.Landmark || "",
    payment_preference: r.Payment_Freq || "Daily bill Payment",
  };
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
  if (r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
  if (r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
  if (r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);

  let bfJson = breakfast;
  try {
    if (r.Breakfast_JSON) bfJson = JSON.parse(r.Breakfast_JSON);
  } catch(e) {}

  return {
    breakfast:    bfJson,
    lunch_dry:    r.Lunch_Dry    || "",
    lunch_curry:  r.Lunch_Curry  || "",
    dinner_dry:   r.Dinner_Dry   || "",
    dinner_curry: r.Dinner_Curry || "",
    cutoff_overrides: co
  };
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
  const payFreq      = profile.payment_preference || "Daily bill Payment";

  // Build the header→index map once
  const hIdx = headerIndex(ordersWs);

  const submissionIds = [];

  for (const order of orders) {
    const orderDate = order.date;

    // Calculate day-level discount once across all meals for this date
    const dayTotal = order.meals.reduce((s, m) => s + (m.subtotal || 0), 0);
    let discRate = 0;
    if (dayTotal >= 450) discRate = 0.075;
    else if (dayTotal >= 300) discRate = 0.05;
    const totalDiscAmt = Math.round(dayTotal * discRate);
    // Pro-rate discount across meals proportionally
    const getDisc = (sub) => dayTotal > 0 ? Math.round(totalDiscAmt * (sub / dayTotal)) : 0;

    for (const meal of order.meals) {
      const sid = generateSubmissionID();
      submissionIds.push(sid);

      const items  = meal.items || [];   // [{colKey, qty}]
      const mealType = meal.type;
      const notes  = meal.notes || "";
      const sub    = meal.subtotal || 0;
      const mealArea  = meal.area || profile.area || "";

      // Delivery charge
      const FREE_AREA = "Bhosale Garden";
      const DELIVERY  = 10;
      const FREE_THR  = 100;
      const delCharge = (mealArea !== FREE_AREA && sub > 0 && sub < FREE_THR) ? DELIVERY : 0;
      const discAmt   = getDisc(sub);
      const netTotal  = sub + delCharge - discAmt;

      // Build items JSON
      const itemsObj = {};
      items.forEach(({colKey, qty}) => {
        const canonical = ITEM_COL_MAP[colKey] || colKey;
        itemsObj[canonical] = qty;
      });

      // Address
      const wing    = meal.wing    || profile.wing    || "";
      const flat    = meal.flat    || profile.flat    || "";
      const floor   = meal.floor   || profile.floor   || "";
      const society = meal.society || profile.society || "";
      const area    = mealArea;
      const fullAddr = [wing && `Wing ${wing}`, flat && `Flat ${flat}`, floor && `${floor} Floor`, society, area]
                        .filter(Boolean).join(", ");
      const mapsLink = meal.maps || profile.maps || "";
      const landmark = meal.landmark || profile.landmark || "";

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
      set("Items_JSON",          JSON.stringify(itemsObj));
      set("Special_Notes",       notes);
      set("Food_Subtotal",       sub);
      set("Delivery_Charge",     delCharge);
      set("Discount_Amount",     discAmt);
      set("Net_Total",           netTotal);
      set("Payment_Method",      payMethod);
      set("Payment_Status",      payStatus);
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

  // Update customer ledger (10-day)
  if (payFreq === "Post paid bill payment for every 10 days") {
    try { _updateLedger(ss, profile, orders); } catch(e) { /* non-fatal */ }
  }

  return {success: true, submissionId: submissionIds[0] || ""};
}

// ── UPSERT CUSTOMER ──────────────────────────────────────────
function _upsertCustomer(ss, profile) {
  const ws = getOrCreateTab(ss, TAB_CUSTOMERS, []);
  const rows = getAllRows(ws);
  const existing = rows.find(r => String(r.Phone).trim() === String(profile.phone).trim());

  const fullAddr = [
    profile.wing    && `Wing ${profile.wing}`,
    profile.flat    && `Flat ${profile.flat}`,
    profile.floor   && `${profile.floor} Floor`,
    profile.society, profile.area
  ].filter(Boolean).join(", ");

  if (existing) {
    const rowNum = existing._row;
    const hIdx = headerIndex(ws);
    const setCell = (col, val) => { if (hIdx[col]) ws.getRange(rowNum, hIdx[col]).setValue(val); };
    setCell("Name",        profile.name    || existing.Name);
    setCell("Area",        profile.area    || existing.Area);
    setCell("Wing",        profile.wing    !== undefined ? profile.wing    : existing.Wing);
    setCell("Flat",        profile.flat    !== undefined ? profile.flat    : existing.Flat);
    setCell("Floor",       profile.floor   !== undefined ? profile.floor  : existing.Floor);
    setCell("Society",     profile.society !== undefined ? profile.society: existing.Society);
    setCell("Full_Address",fullAddr || existing.Full_Address);
    setCell("Maps_Link",   profile.maps    !== undefined ? profile.maps   : existing.Maps_Link);
    setCell("Landmark",    profile.landmark!== undefined ? profile.landmark: existing.Landmark);
    setCell("Payment_Freq",profile.payment_preference || existing.Payment_Freq);
  } else {
    ws.appendRow([
      profile.phone || "",
      profile.name  || "",
      profile.area  || "",
      profile.wing  || "",
      profile.flat  || "",
      profile.floor || "",
      profile.society || "",
      fullAddr,
      profile.maps    || "",
      profile.landmark|| "",
      profile.payment_preference || "Daily bill Payment",
      getISTTimestamp(),
      ""   // Ledger_Sheet_ID filled later
    ]);
  }
}

// ── GET CUSTOMER ORDERS ──────────────────────────────────────
function getCustomerOrders(phone) {
  if (!phone) return {orders:[]};
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

  const upcoming = rows
    .filter(r => String(r.Phone).trim() === String(phone).trim() && r.Order_Date >= today)
    .map(r => ({
      rowId:              r.Submission_ID,
      date:               r.Order_Date,
      meal:               r.Meal_Type,
      summary:            _buildSummary(r),
      total:              r.Net_Total,
      customer_name:      r.Customer_Name      || "",
      area:               r.Area               || "",
      wing:               r.Wing               || "",
      flat:               r.Flat               || "",
      floor:              r.Floor              || "",
      society:            r.Society            || "",
      maps:               r.Maps_Link          || "",
      landmark:           r.Landmark           || "",
      items_json:         r.Items_JSON         || "{}",
      notes:              r.Special_Notes      || "",
      payment_preference: r.Payment_Freq       || "Daily bill Payment",
    }));

  return {orders: upcoming};
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

// ── DELETE ORDER ─────────────────────────────────────────────
function deleteOrder(phone, rowId) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const now = new Date(new Date().getTime() + 5.5 * 3600 * 1000);
  const today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
  const hourIST = now.getHours() + now.getMinutes() / 60;
  const CUTOFFS = { Breakfast: 7, Lunch: 9.5, Dinner: 17 };

  const r = rows.find(x =>
    String(x.Submission_ID) === String(rowId) &&
    String(x.Phone).trim() === String(phone).trim()
  );
  if (!r) return {success: false, error: "Order not found"};
  if (r.Order_Date < today) return {success: false, error: "Cannot delete past orders"};

  // Block deletion if cutoff has passed for today's orders
  if (r.Order_Date === today) {
    const cutoffHour = CUTOFFS[r.Meal_Type];
    if (cutoffHour !== undefined && hourIST >= cutoffHour) {
      return {success: false, error: `Cutoff for ${r.Meal_Type} has already passed`};
    }
  }

  ws.deleteRow(r._row);
  return {success: true};
}

// ── GET 10-DAY RUNNING TOTAL ─────────────────────────────────
function get10DayRunning(phone) {
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);

  const now = new Date(new Date().getTime() + 5.5 * 3600 * 1000);
  const day = now.getDate();
  const y   = now.getFullYear();
  const m   = now.getMonth();
  const monthName = now.toLocaleString("default", {month:"long"});
  const lastDayOfMonth = new Date(y, m + 1, 0).getDate();

  let periodStart, periodEnd, periodLabel;
  let prevStart, prevEnd, prevLabel;

  if (day <= 10) {
    periodStart = new Date(y, m, 1);
    periodEnd   = new Date(y, m, 10);
    periodLabel = `1–10 ${monthName}`;
    // Previous period: 21st–end of last month
    const prevMonthName = new Date(y, m - 1, 1).toLocaleString("default", {month:"long"});
    const lastDayPrevMonth = new Date(y, m, 0).getDate();
    prevStart = new Date(y, m - 1, 21);
    prevEnd   = new Date(y, m, 0);
    prevLabel = `21–${lastDayPrevMonth} ${prevMonthName}`;
  } else if (day <= 20) {
    periodStart = new Date(y, m, 11);
    periodEnd   = new Date(y, m, 20);
    periodLabel = `11–20 ${monthName}`;
    // Previous period: 1st–10th of this month
    prevStart = new Date(y, m, 1);
    prevEnd   = new Date(y, m, 10);
    prevLabel = `1–10 ${monthName}`;
  } else {
    periodStart = new Date(y, m, 21);
    periodEnd   = new Date(y, m + 1, 0);
    periodLabel = `21–${lastDayOfMonth} ${monthName}`;
    // Previous period: 11th–20th of this month
    prevStart = new Date(y, m, 11);
    prevEnd   = new Date(y, m, 20);
    prevLabel = `11–20 ${monthName}`;
  }

  const fmt = d => Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
  const ps = fmt(periodStart), pe = fmt(periodEnd);
  const pps = fmt(prevStart),  ppe = fmt(prevEnd);
  const phoneStr = String(phone).trim();

  const running = rows
    .filter(r =>
      String(r.Phone).trim() === phoneStr &&
      r.Order_Date >= ps && r.Order_Date <= pe &&
      r.Payment_Status !== "Paid"
    )
    .reduce((s, r) => s + (Number(r.Net_Total) || 0), 0);

  const prevDue = rows
    .filter(r =>
      String(r.Phone).trim() === phoneStr &&
      r.Order_Date >= pps && r.Order_Date <= ppe &&
      r.Payment_Status !== "Paid"
    )
    .reduce((s, r) => s + (Number(r.Net_Total) || 0), 0);

  // isDueDate: today is 10th, 20th, or last day of month
  const isDueDate = (day === 10 || day === 20 || day === lastDayOfMonth);

  return {running, period: periodLabel, prevDue, prevLabel, isDueDate};
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
    active: String(r.Active).toLowerCase() !== "false"
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
  const custWs  = getOrCreateTab(ss, TAB_CUSTOMERS, []);
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
  const headers = ["Date","Meal","Items Ordered","Subtotal (₹)","Delivery (₹)","Discount (₹)","Net Total (₹)"];
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
      const summary = Object.entries(JSON.parse(JSON.stringify(meal.items || [])))
        .map(([k,v]) => `${v}×${k}`).join(", ") || "—";
      const delCharge = meal.deliveryCharge || 0;
      const discAmt   = meal.discountAmount  || 0;
      const netTotal  = meal.subtotal + delCharge - discAmt;
      // Append after last data row in the correct period section
      ws.appendRow([order.date, meal.type, summary, meal.subtotal, delCharge, discAmt, netTotal]);
    }
  }
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
  const breads = B.menu.breads.map(function(i){ return i.name + " ₹" + i.price + " " + i.unit; }).join(" | ");
  const sabji  = B.menu.sabji.map(function(i){ return i.name + " ₹" + i.price; }).join(" | ");
  const basics = B.menu.basics.map(function(i){ return i.name + " ₹" + i.price; }).join(" | ");

  // Fetch today's live menu from the Sheet
  var todayMenuSection = "";
  try {
    var ist = new Date(new Date().getTime() + 5.5 * 3600 * 1000);
    var todayStr = Utilities.formatDate(ist, "Asia/Kolkata", "yyyy-MM-dd");
    var todayMenu = getMenu(todayStr);
    var bfItems = (todayMenu.breakfast || []).map(function(x){ return x.name + " ₹" + x.price; }).join(", ");
    var lDry    = todayMenu.lunch_dry    || "";
    var lCurry  = todayMenu.lunch_curry  || "";
    var dDry    = todayMenu.dinner_dry   || "";
    var dCurry  = todayMenu.dinner_curry || "";
    todayMenuSection = "TODAY'S LIVE MENU (" + todayStr + "):\n"
      + (bfItems  ? "- Breakfast items: " + bfItems + "\n" : "- Breakfast: not set yet for today\n")
      + (lDry     ? "- Lunch Dry Sabji: "    + lDry    + "\n" : "")
      + (lCurry   ? "- Lunch Curry Sabji: "  + lCurry  + "\n" : "")
      + (dDry     ? "- Dinner Dry Sabji: "   + dDry    + "\n" : "")
      + (dCurry   ? "- Dinner Curry Sabji: " + dCurry  + "\n" : "")
      + (!lDry && !lCurry && !dDry && !dCurry ? "- Today's sabji not set yet — direct user to WhatsApp group for updates.\n" : "")
      + "\n";
  } catch(e) {
    todayMenuSection = "TODAY'S LIVE MENU: Unable to fetch right now — direct user to WhatsApp group.\n\n";
  }

  return "You are a friendly and helpful customer service assistant for " + B.name + ", a " + B.type + " in Hadapsar, Pune.\n\n"
    + "ABOUT US:\n" + B.about + "\n"
    + "Our vision: " + B.vision + "\n\n"
    + "SERVING AREAS: " + B.locations_served.join(", ") + "\n\n"
    + "ORDER CUTOFF TIMES (orders must be placed BEFORE these times):\n"
    + "- Breakfast: " + B.order_cutoffs.breakfast + "\n"
    + "- Lunch: "     + B.order_cutoffs.lunch     + "\n"
    + "- Dinner: "    + B.order_cutoffs.dinner    + "\n"
    + "- Closed on: " + B.order_cutoffs.closed_on + "\n\n"
    + todayMenuSection
    + "FULL MENU & PRICES:\n"
    + "Breads: "    + breads  + "\n"
    + "Sabji: "     + sabji   + "\n"
    + "Basics: "    + basics  + "\n"
    + "Breakfast: " + B.menu.breakfast + "\n"
    + "Note: "      + B.menu.note     + "\n\n"
    + "DELIVERY:\n"
    + "- " + B.delivery.charge + "\n"
    + "- " + B.delivery.per_meal_address + "\n\n"
    + "DISCOUNTS (automatically applied per day total):\n"
    + "- " + B.discounts.tier1 + "\n"
    + "- " + B.discounts.tier2 + "\n"
    + "- " + B.discounts.note  + "\n\n"
    + "PAYMENT OPTIONS: " + B.payment.options.join(", ") + "\n"
    + "UPI ID: " + B.payment.upi_id + "\n"
    + "10-Day billing cycle: " + B.payment.ten_day + "\n\n"
    + "HOW TO ORDER:\n"
    + "- Order form: " + B.ordering.order_url + "\n"
    + "- Process: " + B.ordering.process + "\n"
    + "- Advance ordering: " + B.ordering.advance + "\n"
    + "- Edit/cancel: " + B.ordering.edit_cancel + "\n"
    + "- " + B.ordering.no_login + "\n\n"
    + "CONTACT:\n"
    + "- WhatsApp: " + B.contact.whatsapp + " (" + B.contact.whatsapp_link + ")\n"
    + "- WhatsApp group (daily menu updates): " + B.contact.whatsapp_group + "\n"
    + "- Phone: " + B.contact.phone_primary + " (alternate: " + B.contact.phone_alt + ")\n"
    + "- Email: " + B.contact.email + "\n"
    + "- Google page / reviews: " + B.contact.google_page + "\n\n"
    + "INSTRUCTIONS:\n"
    + "- Respond in the same language the customer writes in (English, Hindi, or Marathi).\n"
    + "- Keep responses concise, warm, and helpful.\n"
    + "- When a customer wants to place an order, always direct them to: " + B.ordering.order_url + "\n"
    + "- When asked about today's menu, use the TODAY'S LIVE MENU section above — if sabji is not set, direct them to the WhatsApp group.\n"
    + "- Do NOT invent prices, availability, or promotions not listed above.\n"
    + "- Do NOT discuss competitors or make comparisons.\n"
    + "- If unsure about anything, direct the customer to WhatsApp: " + B.contact.whatsapp + ".";
}

function callGemini(systemPrompt, history, userMessage) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return "I'm having trouble connecting right now. Please WhatsApp us at +91 99307 48908 for help!";
  }

  // Build contents array from history + current message
  const contents = [];
  (history || []).forEach(function(msg) {
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
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(response.getContentText());
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
    (r.Payment_Status === "10-Day Pending" ||
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
        (r.Payment_Status === "10-Day Pending" ||
         r.Payment_Status === "Pending"         ||
         !r.Payment_Status)) {
      ws.getRange(r._row, hIdx["Payment_Status"]).setValue("Paid");
      updated++;
    }
  });
  return {success:true, updatedRows:updated, customersMarked:phones.length};
}
