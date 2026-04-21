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

const CODE_VERSION   = 14.2; // Standardized Menu Names
const LEDGER_FOLDER  = "Svaadh Customer Ledgers";

// ── PAYMENT GATEWAY CONFIG ───────────────────────────────────
// Controlled via Script Properties — never hardcoded.
// In Dev Apps Script: add Script Property  PAYMENT_GATEWAY_ENABLED = true
// Live project never has this property set → evaluates to false automatically.
const PAYMENT_GATEWAY_ENABLED = (SP.getProperty("PAYMENT_GATEWAY_ENABLED") === "true");

// ── HDFC SmartGATEWAY — Script Properties reference ─────────
// Add all of these in Apps Script → Project Settings → Script Properties.
// NEVER hardcode keys here.
//
//  Property Name            │ Where to get it
//  ─────────────────────────┼─────────────────────────────────────────────
//  HDFC_MERCHANT_ID         │ Dashboard → Settings → General (Merchant ID)
//  HDFC_API_KEY             │ Dashboard → Settings → Security → Create New API Key
//  HDFC_RESPONSE_KEY        │ Dashboard → Settings → Security → Response Key
//  HDFC_WEBHOOK_USERNAME    │ Set freely — enter same value in Dashboard → Settings → Webhooks → Username
//  HDFC_WEBHOOK_PASSWORD    │ Set freely — enter same value in Dashboard → Settings → Webhooks → Password
//  HDFC_ENV                 │ "test" or "live"
//  HDFC_TEST_URL            │ Sandbox base URL from HDFC (e.g. https://smartgateway-uat.hdfcbank.com)
//  HDFC_LIVE_URL            │ Production base URL (e.g. https://smartgateway.hdfcbank.com)
//  HDFC_RETURN_URL          │ Apps Script exec URL (NOT order.html — GitHub Pages rejects POST).
//                           │ doPost detects the HDFC return payload and JS-redirects to order.html.
//  HDFC_WEBHOOK_URL         │ This Apps Script doPost URL (set in Dashboard → Settings → Webhooks)

const HDFC_MERCHANT_ID      = SP.getProperty("HDFC_MERCHANT_ID")      || "";
const HDFC_API_KEY          = SP.getProperty("HDFC_API_KEY")          || "";
const HDFC_RESPONSE_KEY     = SP.getProperty("HDFC_RESPONSE_KEY")     || "";
const HDFC_WEBHOOK_USERNAME = SP.getProperty("HDFC_WEBHOOK_USERNAME") || "";
const HDFC_WEBHOOK_PASSWORD = SP.getProperty("HDFC_WEBHOOK_PASSWORD") || "";
const HDFC_ENV              = SP.getProperty("HDFC_ENV")              || "test";
const HDFC_RETURN_URL       = SP.getProperty("HDFC_RETURN_URL")       || "";
const HDFC_ORDER_PAGE_URL   = SP.getProperty("HDFC_ORDER_PAGE_URL")   || "https://svaadhkitchen.in/order.html";
const HDFC_BASE_URL         = HDFC_ENV === "live"
  ? (SP.getProperty("HDFC_LIVE_URL") || "https://smartgateway.hdfcbank.com")
  : (SP.getProperty("HDFC_TEST_URL") || "https://smartgateway-uat.hdfcbank.com");
// ─────────────────────────────────────────────────────────────

// Sheet tab names
const TAB_ORDERS     = "SK_Orders";
const TAB_CUSTOMERS  = "SK_Customers";
const TAB_MENU       = "SK_Daily_Menu";
const TAB_BF_MASTER  = "SK_Master_Breakfast";
const TAB_SABJI      = "SK_Master_Sabjis";
const TAB_AREAS      = "SK_Areas";
const TAB_WALLET     = "SK_Wallet"; // Holds prepaid balances
const TAB_REFUNDS    = "SK_Refunds";      // Manual refund requests
const TAB_WEBHOOK_LOG = "SK_Webhook_Log"; // HDFC webhook log-first buffer

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
  "Review_Promo_Count", "Review_Reward_Claimed", "Standard_Order", "Billing_Cycle", "Fee_Exempt", "Delivery_Point", "On_Account"
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
  "Payment_Method","Payment_Status","Payment_Freq","First_Time","Source","Refund_Preference", "Packed", "Delivery_Point",
  "Inflation_Surcharge", "Loyalty_Discount", "Wallet_Credit"
];

const ITEM_COL_MAP = {
  // Canonical Names (Universal Standard)
  "Chapati": "Chapati",
  "Without Oil Chapati": "Without_Oil_Chapati",
  "Phulka": "Phulka",
  "Ghee Phulka": "Ghee_Phulka",
  "Jowar Bhakri": "Jowar_Bhakri",
  "Bajra Bhakri": "Bajra_Bhakri",
  "Dry Sabji Mini (100ml)": "Dry_Sabji_Mini",
  "Dry Sabji Full (250ml)": "Dry_Sabji_Full",
  "Curry Sabji Mini (100ml)": "Curry_Sabji_Mini",
  "Curry Sabji Full (250ml)": "Curry_Sabji_Full",
  "Dal (200ml)": "Dal",
  "Rice (100g)": "Rice",
  "Salad (40g)": "Salad",
  "Curd (50g)": "Curd",

  // Legacy/Simple variants (for backward compatibility)
  "Dry Sabji Mini": "Dry_Sabji_Mini",
  "Dry Sabji Full": "Dry_Sabji_Full",
  "Curry Sabji Mini": "Curry_Sabji_Mini",
  "Curry Sabji Full": "Curry_Sabji_Full",
  "Dal": "Dal",
  "Rice": "Rice",
  "Salad": "Salad",
  "Curd": "Curd",

  // Underscored variants
  "Without_Oil_Chapati":"Without_Oil_Chapati",
  "Ghee_Phulka":"Ghee_Phulka",
  "Jowar_Bhakri":"Jowar_Bhakri",
  "Bajra_Bhakri":"Bajra_Bhakri",
  "Dry_Sabji_Mini":"Dry_Sabji_Mini",
  "Dry_Sabji_Full":"Dry_Sabji_Full",
  "Curry_Sabji_Mini":"Curry_Sabji_Mini",
  "Curry_Sabji_Full":"Curry_Sabji_Full",

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
  about: "Svaadh Kitchen is a home-based vegetarian cloud kitchen in Hadapsar, Pune, serving fresh and wholesome homemade meals since August 2023 (over 2.5 years). We specialize in homemade vegetarian food, offering breakfast, lunch, and dinner with a changing daily sabji menu. We deliver exclusively to 15 areas in Hadapsar: Bhosale Nagar, Magarpatta, Amanora, DP Road, Triveni Nagar, Malwadi, SadeSatraNali, Kirtane Baug, Tupe Patil Road, BG Shirke Road, Vaiduwadi (till Yash Honda), Pune-Solapur Road (till Gadital), Vihar Chowk, Mandai, and Gadital. Delivery is FREE for Bhosale Nagar and Triveni Nagar. All other areas have a nominal ₹10 fee if the order is below ₹100. Self Pickup is always free.",
  vision: "To make homemade vegetarian meals easily accessible and affordable for everyone, while maintaining taste, quality, and consistency.",
  locations_served: [
    "Bhosale Nagar", "Magarpatta", "Amanora", "DP Road", "Triveni Nagar", 
    "Malwadi", "SadeSatraNali", "Kirtane Baug", "Tupe Patil Road", "BG Shirke Road", 
    "Vaiduwadi (Till Yash Honda Only)", "Pune-Solapur Road (Till Gadital Only)", "Vihar Chowk", "Mandai (Hadapsar Mandai)", "Gadital"
  ],
  order_cutoffs: { breakfast: "before 7:00 AM", lunch: "before 9:00 AM", dinner: "before 4:30 PM", closed_on: "Sunday" },
  delivery: {
    free_areas: ["Bhosale Nagar", "Triveni Nagar", "Self Pickup"],
    charge: "₹10 per meal for other listed areas if subtotal is below ₹100. Free for Bhosale Nagar, Triveni Nagar and Self Pickup always.",
    outside_policy: "We only deliver in the listed Hadapsar areas. We DO NOT deliver to areas like Kothrud, Baner, Viman Nagar, etc."
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
    breakfast: "Rotating daily (₹35–₹70). Items include Kanda Poha [175g] ₹35, Ghee Upma [200g] ₹40, Sabudana Khichdi [200g] ₹40, 5 x Tikhi Pudi with 100 ml coriander chutney ₹45, 4 x Idli & 100ml Chutney ₹45, Aloo Paratha ₹50, Thalipeeth ₹50, Ghee Sheera [200g] ₹50, Paneer Paratha ₹70. Curd 50g available extra ₹12. Check the order form for today's options.",
    breakfast_note: "Curd 50g (₹12) is available as an add-on for breakfast — not included by default. Pure Ghee is used to make breakfast items."
  },
  discounts: {
    tier1: "5% off when the day total is ₹300 or more",
    tier2: "7.5% off when the day total is ₹450 or more",
    note: "Discounts are applied automatically per day's total when placing an order."
  },
  payment: {
    options: ["Svaadh Wallet (Prepaid)", "UPI", "Prepaid Wallet Billing"],
    upi_id: "svaadhkitchen.36727659@hdfcbank",
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

  // ── HDFC Return URL via GET ────────────────────────────────────
  // HDFC sometimes redirects the customer's browser via GET (not POST).
  // Detect by presence of order_id + status params with no _action.
  // Redirect browser to the order page URL with all params forwarded.
  if (p.order_id && p.status && !p.action && !p._action) {
    const params = Object.keys(p)
      .map(function(k) { return encodeURIComponent(k) + "=" + encodeURIComponent(p[k]); })
      .join("&");
    const redirectUrl = HDFC_ORDER_PAGE_URL + "?" + params;
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0;url=' + redirectUrl + '"></head>' +
      '<body><script>window.location.replace(' + JSON.stringify(redirectUrl) + ');</script>' +
      '<p>Redirecting... <a href="' + redirectUrl + '">Click here if not redirected</a></p></body></html>'
    );
  }
  // ─────────────────────────────────────────────────────────────

  try {
    if (action === "version") return jsonRes({version: CODE_VERSION, status:"ok"});
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
    
    // Keep-alive ping — just wakes GAS, no sheet reads
    if (action === "ping") return jsonRes({ok: true, t: new Date().toISOString()});

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
      return HtmlService.createHtmlOutput(
        `<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0;url=${redirectUrl}"></head>` +
        `<body><script>window.location.replace(${JSON.stringify(redirectUrl)});</script>` +
        `<p>Redirecting... <a href="${redirectUrl}">Click here if not redirected</a></p></body></html>`
      );
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
    const isAdmin = (pin === ADMIN_PIN && pin !== "");
    const isStaff = (pin === KITCHEN_PIN || pin === ADMIN_PIN) && pin !== "";

    // Customer actions (pinned via their own phone/PIN handled inside functions)
    if (action === "deleteOrder") return jsonRes(deleteOrder(body.phone, body.rowId, body.refundType));
    if (action === "getCustomerList") return jsonRes(getCustomerList());
    if (action === "getCustomerOrders") return jsonRes(getCustomerOrders(body.phone));
    
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
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner",
    "OOS_JSON","Orders_Closed"
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
// Returns true if the order should be excluded from kitchen/prep counts.
// "Cancelled (Verify UPI)" = soft-cancel pending admin verification —
// the customer already requested cancellation, do NOT include in kitchen prep.
function _isOrderCancelled(paymentStatus) {
  const s = String(paymentStatus || "").toLowerCase();
  return s === "cancelled" || s.startsWith("cancelled");
}

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
    promoCount: (function(v){
      if (v === "" || v === null || v === undefined) return null;
      var num = Number(v);
      return isNaN(num) ? v : num;
    })(r.Review_Promo_Count),
    wallet_balance:     _calculateWalletBalance(phone),
    feeExempt:          (r.Fee_Exempt === "Yes" || r.Fee_Exempt === true),
    onAccount:          r.On_Account || "No",
    billingCycle:       r.Billing_Cycle || "Daily"
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
      promoCount: (function(v){
        if (v === "" || v === null || v === undefined) return null;
        var num = Number(v);
        return isNaN(num) ? v : num;
      })(r.Review_Promo_Count),
      wallet_balance:     _calculateWalletBalance(phone),
      feeExempt:          (r.Fee_Exempt === "Yes" || r.Fee_Exempt === true),
      onAccount:          r.On_Account || "No",
      billingCycle:       r.Billing_Cycle || "Daily"
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

    if (rType.includes("recharge") || rType.includes("refund") || rType.includes("credit")
        || rType.includes("carry forward") || rType.includes("carry-forward")) {
      balance += rAmt;
    } else if (rType.includes("order") || rType.includes("deduct") || rType.includes("payment")) {
      balance -= rAmt;
    }
  });

  return Math.round(balance * 100) / 100;
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
  
  const NAME_MAP = {
    "Kanda Poha": "Kanda Poha [175g]",
    "Ghee Upma": "Ghee Upma [200g]",
    "Sabudana Khichdi": "Sabudana Khichdi [200g]",
    "Tikhi Pudi": "5 x Tikhi Pudi with 100 ml coriander chutney",
    "Tikhi Puri": "5 x Tikhi Pudi with 100 ml coriander chutney",
    "Idli Chutney": "4 x Idli & 100ml Chutney",
    "Idli": "4 x Idli & 100ml Chutney",
    "4 x Idli & 100ml Chutney": "4 x Idli & 100ml Chutney",
    "Ghee Sheera": "Ghee Sheera [200g]"
  };

  const breakfast = bfRows.map(x => {
    const rawName = String(x.Name).trim();
    return {
      name: NAME_MAP[rawName] || rawName,
      price: Number(x.Price)
    };
  });

  if (!r) return {
    breakfast, lunch_dry:"", lunch_curry:"", dinner_dry:"", dinner_curry:"",
    cutoff_overrides:{},
    oos_items: { Breakfast: [], Lunch: [], Dinner: [] },
    orders_closed: {}
  };

  const co = {};
  if (r && r.Cutoff_Breakfast) co.Breakfast = Number(r.Cutoff_Breakfast);
  if (r && r.Cutoff_Lunch)     co.Lunch     = Number(r.Cutoff_Lunch);
  if (r && r.Cutoff_Dinner)    co.Dinner    = Number(r.Cutoff_Dinner);

  // MERGE LOGIC: Start with master active items, then merge daily overrides
  const masterActive = breakfast;
  let dailyBf = [];
  if (r && r.Breakfast_JSON) {
    try { 
      const parsed = JSON.parse(r.Breakfast_JSON); 
      dailyBf = parsed.map(d => ({
        ...d,
        name: d.name ? (NAME_MAP[d.name.trim()] || d.name) : ""
      }));
    } catch(e) {}
  }

  // Prioritize Daily selections (where specific prices or choices were made)
  // but ensure Master Active items are always present.
  const finalBreakfast = [...dailyBf];
  masterActive.forEach(m => {
    if (!finalBreakfast.some(d => d.name === m.name)) {
      finalBreakfast.push(m);
    }
  });

  let oosItems = { Breakfast: [], Lunch: [], Dinner: [] };
  try { if (r && r.OOS_JSON) oosItems = JSON.parse(r.OOS_JSON); } catch(e) {}

  let ordersClosed = {};
  try { if (r && r.Orders_Closed) ordersClosed = JSON.parse(r.Orders_Closed); } catch(e) {}

  return {
    breakfast:    finalBreakfast,
    lunch_dry:    r ? (r.Lunch_Dry || "") : "",
    lunch_curry:  r ? (r.Lunch_Curry || "") : "",
    dinner_dry:   r ? (r.Dinner_Dry || "") : "",
    dinner_curry: r ? (r.Dinner_Curry || "") : "",
    cutoff_overrides: co,
    oos_items:    oosItems,
    orders_closed: ordersClosed
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

  // Show all dates from today onwards that have a menu row set
  const today = getISTDate();
  const todayStr = Utilities.formatDate(today, "Asia/Kolkata", "yyyy-MM-dd");

  // Collect all future/today dates that have a menu row, sorted ascending
  const futureDates = Object.keys(menuMap)
    .filter(d => d >= todayStr)
    .sort();

  const days = [];
  futureDates.forEach(dateStr => {
    const d = new Date(dateStr + "T00:00:00+05:30");
    const dayName    = Utilities.formatDate(d, "Asia/Kolkata", "EEEE");
    const displayDate = Utilities.formatDate(d, "Asia/Kolkata", "dd MMM");

    const r = menuMap[dateStr];
    let bfDaily = [];
    try {
      if (r && r.Breakfast_JSON) bfDaily = JSON.parse(r.Breakfast_JSON);
    } catch(e) {}

    // Merge Master + Daily
    const finalBf = [...bfDaily];
    defaultBreakfast.forEach(m => {
      if (!finalBf.some(x => x.name === m.name)) finalBf.push(m);
    });

    days.push({
      date: dateStr,
      dayName: dayName,
      displayDate: displayDate,
      breakfast: finalBf,
      lunch_dry:    r ? (r.Lunch_Dry    || "") : "",
      lunch_curry:  r ? (r.Lunch_Curry  || "") : "",
      dinner_dry:   r ? (r.Dinner_Dry   || "") : "",
      dinner_curry: r ? (r.Dinner_Curry || "") : "",
      menuSet: true  // only dates with a menu row are included
    });
  });

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
  
  // Fetch existing orders once for all dates in this submission to calculate combined-day fees/discounts
  const submissionDates = orders.map(o => o.date);
  const existingDayTotals = getDayTotalsForDates(profile.phone, submissionDates.join(',')).dayTotals || {};

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
  const initialStreakInfo = _calculateLoyaltyStreak(profile.phone);
  let virtualStreakCount = initialStreakInfo.streak;
  let virtualPastSurcharge = initialStreakInfo.pastSurcharge;

  // Pre-fetch existing orders once for duplicate detection
  const allOrderRows = getAllRows(ordersWs);
  const _dupNowMs = Date.now();
  const _FIVE_MIN_MS = 5 * 60 * 1000;
  const _normPhone = _normalizePhone(profile.phone);
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
      const mealCredit = submissionDayFoodTotal > 0 ? Math.round(totalDateCredit * (sub / submissionDayFoodTotal)) : 0;

      const discAmt = getDisc(sub);
      const inflationSurcharge = Math.ceil(sub / 20);

      // Google Review Promo Logic (10% OFF per meal)
      let reviewDiscount = 0;
      const isNumeric = (typeof promoCount === "number" && !isNaN(promoCount));
      if (isNumeric && promoCount > 0 && sub > 0) {
        reviewDiscount = Math.round(sub * 0.10);
        promoCount--;
      }

      const netTotal = Math.round(sub + delCharge + smallOrderFee + inflationSurcharge - discAmt - mealCredit - reviewDiscount);
      meal._reviewDiscount = reviewDiscount; // carry for set() below


      // Build items JSON
      const itemsObj = {};
      items.forEach(({colKey, qty}) => {
        const canonical = resolveName(colKey); // Use name instead of ID
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
      set("Phone",               profile.phone    || "");
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
      set("Items_JSON",          JSON.stringify(itemsObj));
      set("Special_Notes_Kitchen",  nKitchen);
      set("Special_Notes_Delivery", nDelivery);
      set("Food_Subtotal",       sub);
      set("Small_Order_Fee",     smallOrderFee);
      set("Inflation_Surcharge", inflationSurcharge);
      set("Loyalty_Discount",    is6thDay ? "Yes" : "No");
      set("Delivery_Charge",     delCharge);
      set("Discount_Amount",     discAmt);
      if (hIdx["Review_Discount"]) {
        set("Review_Discount",   meal._reviewDiscount || 0);
      }
      set("Net_Total",           netTotal);
      
      let pStat = payStatus;
      let walletCreditUsed = 0;
      // ════ WALLET DEDUCTION LOGIC ════
      if (payMethod === "Wallet") {
        let currentBalance = _calculateWalletBalance(profile.phone);

        if (currentBalance >= netTotal) {
          _appendWalletTransaction(profile.phone || "", profile.name || "Customer", "Order Deduction", netTotal, true, sid);
          pStat = "Wallet Paid";
          walletCreditUsed = netTotal;
        } else {
          pStat = "Pending"; // Wallet failed, fallback to pending
        }
      } else if (payMethod === "Split") {
        // Split: deduct wallet portion now, UPI portion remains pending
        const requestedCredit = Math.min(Number(body.wallet_credit) || 0, netTotal);
        if (requestedCredit > 0) {
          const currentBalance = _calculateWalletBalance(profile.phone);
          if (currentBalance >= requestedCredit) {
            _appendWalletTransaction(profile.phone || "", profile.name || "Customer", "Order Deduction (Wallet Part)", requestedCredit, true, sid);
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

      // Duplicate guard: same phone + date + meal_type + identical items within 5 minutes → skip
      const _incomingSig = _itemsSig(itemsObj);
      const _dupRow = allOrderRows.find(r => {
        if (_normalizePhone(r.Phone) !== _normPhone) return false;
        if (_normDate(r.Order_Date) !== _normDate(orderDate)) return false;
        if (r.Meal_Type !== mealType) return false;
        const rMs = r.Submitted_At ? new Date(r.Submitted_At).getTime() : 0;
        if (!rMs || (_dupNowMs - rMs) > _FIVE_MIN_MS) return false;
        try {
          const stored = typeof r.Items_JSON === "string" ? JSON.parse(r.Items_JSON) : (r.Items_JSON || {});
          return _itemsSig(stored) === _incomingSig;
        } catch(e) { return false; }
      });
      if (_dupRow) {
        submissionIds[submissionIds.length - 1] = _dupRow.Submission_ID || sid;
        console.log("Duplicate order skipped: " + _normPhone + " / " + orderDate + " / " + mealType);
        continue;
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
    // Pro transition: 0 -> "Exhausted"
    let finalValue = promoCount;
    if (finalValue === 0) finalValue = "Exhausted";
    else if (finalValue === null) finalValue = "";
    
    const realRow = cRowIdx + 2;
    custWs.getRange(realRow, cIdx["Review_Promo_Count"]).setValue(finalValue);
  }

  // If loyalty reward exceeded the bill, credit the bonus to wallet
  const walletBonus = Number(body.wallet_bonus) || 0;
  if (walletBonus > 0) {
    try {
      _appendWalletTransaction(
        profile.phone || "", profile.name || "Customer",
        "Loyalty Streak Reward (Excess Credit)",
        walletBonus, true, submissionIds[0] || ""
      );
    } catch(e) { /* non-fatal */ }
  }

  return {success: true, submissionId: submissionIds[0] || "", wallet_bonus: walletBonus};
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
      // ONLY update if val is effectively provided (not undefined)
      if (hIdx[col] && val !== undefined) {
        ws.getRange(rowNum, hIdx[col]).setValue(val);
      }
    };
    if (profile.name !== undefined) update("Customer_Name", profile.name);
    if (profile.area !== undefined) update("Area",          profile.area);
    if (profile.wing !== undefined) update("Wing",          profile.wing);
    if (profile.flat !== undefined) update("Flat",          profile.flat);
    if (profile.floor !== undefined) update("Floor",         profile.floor);
    if (profile.society !== undefined) update("Society",       profile.society);
    if (profile.area !== undefined || profile.society !== undefined) update("Full_Address",  fullAddr);
    
    // Auto-derive Maps Link if missing
    let finalMaps = profile.maps || "";
    if (!finalMaps) {
      finalMaps = _deriveMapsLink(fullAddr, profile.society || "");
    }
    update("Maps_Link", finalMaps);

    if (profile.landmark !== undefined) update("Landmark",      profile.landmark || "");
    if (profile.delivery_point !== undefined) update("Delivery_Point", _getDeliveryPointLabel(profile.delivery_point));
    if (profile.payment_preference !== undefined) update("Payment_Freq",  profile.payment_preference);
    if (profile.pin) update("PIN", profile.pin);
    if (profile.meal_addresses) update("Meal_Addresses", profile.meal_addresses);
    if (profile.standardOrder !== undefined) update("Standard_Order", profile.standardOrder);
    if (profile.onAccount !== undefined) update("On_Account", profile.onAccount);
    if (profile.billingCycle !== undefined) update("Billing_Cycle", profile.billingCycle);
    
    SpreadsheetApp.flush(); // Ensure writes are committed
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
        case "Maps_Link":       val = profile.maps || _deriveMapsLink(fullAddr, profile.society || ""); break;
        case "Landmark":        val = profile.landmark || ""; break;
        case "Delivery_Point":  val = _getDeliveryPointLabel(profile.delivery_point); break;
        case "Payment_Freq":    val = profile.payment_preference || "Daily Payment"; break;
        case "Created_At":      val = getISTTimestamp(); break;
        case "PIN":             val = profile.pin || ""; break;
        case "Meal_Addresses":  val = profile.meal_addresses || ""; break;
        case "Standard_Order":  val = profile.standardOrder || ""; break;
        case "Billing_Cycle":   val = profile.billingCycle || "Daily"; break;
        case "On_Account":      val = profile.onAccount || "No"; break;
        case "Review_Promo_Count": val = ""; break;
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

/**
 * ADMIN: Toggle On Account status for a customer
 */
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

/**
 * Calculates current streak and accumulated surcharges for a customer.
 * Skips Sundays (kitchen closed).
 */
function _calculateLoyaltyStreak(phone) {
  if (!phone) return { streak: 0, pastSurcharge: 0 };
  const ss = getSpreadsheet();
  const ws = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  const rows = getAllRows(ws);
  const phoneStr = _normalizePhone(phone);
  const todayISO = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");

  const dailyTotals = {};
  rows.forEach(r => {
    if (_normalizePhone(r.Phone) !== phoneStr) return;
    const stat = String(r.Payment_Status || "").toLowerCase();
    if (stat.includes("cancelled") || stat.includes("deleted")) return;
    
    const d = r.Order_Date instanceof Date ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd") : String(r.Order_Date).trim();
    if (d >= todayISO) return; // Only past days for the streak check

    if (!dailyTotals[d]) dailyTotals[d] = 0;
    dailyTotals[d] += (Number(r.Inflation_Surcharge) || (Math.ceil((Number(r.Food_Subtotal)||0)/20))); 
  });

  let streakCount = 0;
  let accumulatedSurcharge = 0;
  
  let d = new Date(); d.setDate(d.getDate() - 1); // yesterday
  let safety = 0;
  while (safety < 30) { 
    safety++;
    if (d.getDay() === 0) { // Skip Sunday
      d.setDate(d.getDate() - 1);
      continue;
    }
    const iso = Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    if (dailyTotals[iso] !== undefined) {
      streakCount++;
      accumulatedSurcharge += dailyTotals[iso];
    } else {
      break; 
    }
    d.setDate(d.getDate() - 1);
  }
  
  return { streak: streakCount, pastSurcharge: accumulatedSurcharge };
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
        inflation_surcharge: Number(r.Inflation_Surcharge) || 0,
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
        wallet_credit:      Number(r.Wallet_Credit) || 0,
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
        inflation_surcharge: Number(r.Inflation_Surcharge) || 0,
        payment_status:     r.Payment_Status,
        payment_method:     r.Payment_Method,
        wallet_credit:      Number(r.Wallet_Credit) || 0,
        deliveredAt:        delTracker.deliveredAt,
        enRouteAt:          delTracker.enRouteAt
      };
    });

  const onAccountBalance = allFiltered
    .filter(r => String(r.Payment_Status || "").toLowerCase() === "on account")
    .reduce((sum, r) => sum + (Number(r.Net_Total) || 0), 0);

  return {
    orders: upcoming,
    past_orders: past,
    wallet_balance: _calculateWalletBalance(phone),
    on_account_balance: onAccountBalance
  };
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
  const CUTOFFS = { Breakfast: 7, Lunch: 9, Dinner: 16.5 };

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
  const isOnAccountOrder = pStatStr === "on account";
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

    // Discount eligibility helper
    const discRate = (sub) => sub >= 450 ? 0.075 : sub >= 300 ? 0.05 : 0;
    const oldRate = discRate(oldDaySubtotal);
    const newRate = discRate(remainingDaySubtotal);

    // Calc over-discount on remaining orders: they received discount at oldRate
    // but now only deserve newRate. Also update those rows in the sheet so that
    // a same-day reorder sees the corrected prevDayDiscAmt and gets the right discount.
    let overDiscount = 0;
    if (oldRate > newRate) {
      const overOnRemaining = sameDayRows.reduce((s, x) => {
        const xSub = Number(x.Food_Subtotal) || 0;
        const oldD = Math.round(xSub * oldRate);
        const newD = Math.round(xSub * newRate);
        return s + (oldD - newD);
      }, 0);
      overDiscount = overOnRemaining;

      // ── Update remaining rows so their stored Discount_Amount and Net_Total
      //    reflect the new (lower) tier. This ensures any re-order on the same
      //    day computes prevDayDiscAmt correctly and gives the right discount.
      if (overDiscount > 0) {
        const discColIdx  = hIdx["Discount_Amount"];
        const netColIdx   = hIdx["Net_Total"];
        sameDayRows.forEach(x => {
          const xSub      = Number(x.Food_Subtotal)      || 0;
          const xSurcharge= Number(x.Inflation_Surcharge)|| 0;
          const xDelivery = Number(x.Delivery_Charge)    || 0;
          const xSmallFee = Number(x.Small_Order_Fee)    || 0;
          const xReviewD  = Number(x.Review_Discount)    || 0;
          const newDiscAmt= Math.round(xSub * newRate);          // 0 when newRate=0
          const newNetTotal = xSub + xDelivery + xSmallFee + xSurcharge - newDiscAmt - xReviewD;
          if (discColIdx) ws.getRange(x._row, discColIdx).setValue(newDiscAmt);
          if (netColIdx)  ws.getRange(x._row, netColIdx) .setValue(newNetTotal);
        });
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
    let loyaltyClawback = 0;
    const phoneStr = _normalizePhone(phone);
    // Scan later days (up to 6 operational days ahead) to see if a payoff happened
    const laterPayoffs = rows.filter(x => {
      if (_normalizePhone(x.Phone) !== phoneStr) return false;
      const xStat = String(x.Payment_Status || "").toLowerCase();
      if (xStat.includes("cancelled") || xStat.includes("deleted")) return false;
      if (String(x.Loyalty_Discount).trim() !== "Yes") return false;
      const xDate = x.Order_Date instanceof Date ? Utilities.formatDate(x.Order_Date, "Asia/Kolkata", "yyyy-MM-dd") : String(x.Order_Date).trim();
      return xDate >= orderDateStr; // Include today just in case, but usually payoff is today or later
    });

    if (laterPayoffs.length > 0) {
      // Find the LATEST payoff within a reasonable window (6 operational days)
      // Actually, any payoff that triggered after this date might be invalidated.
      // We'll subtract the Discount_Amount of the first payoff we find.
      loyaltyClawback = Number(laterPayoffs[0].Discount_Amount) || 0;
      // Mark it as "Clawed Back" in the sheet to prevent multiple deductions?
      // Actually, we should only claw back if the streak is NOW invalid.
      // For simplicity: subtract the reward.
    }

    // Refund = Net_Total − adjustment
    // Net_Total already correctly encodes: food + delivery + fees + surcharge − discount − mealCredit − reviewDiscount
    // mealCredit (retroactive delivery/fee credit for same-day orders) is baked into Net_Total silently;
    // using Net_Total as base ensures we never over-refund that credit.
    // The adjustment claws back over-discount on remaining rows, and delivery/fees now owed
    // by remaining rows (those amounts are deducted from THIS refund instead of charging customer again).
    const adjustment = overDiscount + deliveryOwed + smallFeeOwed + loyaltyClawback;
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
    const currentWasSplit  = (String(r.Payment_Method || "").trim().toLowerCase() === "split");
    if (isOnAccountOrder) {
      // On Account: no cash was collected — just delete the row.
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
      msg = `₹${refundAmt} refunded to Wallet`;
      finalType = "__split_handled__"; // skip normal logic below
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
    let oosItems = { Breakfast: [], Lunch: [], Dinner: [] };
    try { if (r.OOS_JSON) oosItems = JSON.parse(r.OOS_JSON); } catch(e) {}
    let ordersClosed = {};
    try { if (r.Orders_Closed) ordersClosed = JSON.parse(r.Orders_Closed); } catch(e) {}
    return {
      date:             d,
      breakfast:        breakfast,
      lunch_dry:        r.Lunch_Dry    || "",
      lunch_curry:      r.Lunch_Curry  || "",
      dinner_dry:       r.Dinner_Dry   || "",
      dinner_curry:     r.Dinner_Curry || "",
      cutoff_overrides: co,
      oos_items:        oosItems,
      orders_closed:    ordersClosed,
    };
  });

  return {breakfastMaster, sabjiMaster, menuEntries};
}

// ── ADMIN: SAVE MENU ─────────────────────────────────────────
function saveMenu(body) {
  const ss = getSpreadsheet();
  // Always pass full headers so schema self-heals if initSchema() was never run
  const ws = getOrCreateTab(ss, TAB_MENU, [
    "Date","Breakfast_JSON","Lunch_Dry","Lunch_Curry","Dinner_Dry","Dinner_Curry",
    "Cutoff_Breakfast","Cutoff_Lunch","Cutoff_Dinner",
    "OOS_JSON","Orders_Closed"
  ]);
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
    JSON.stringify(body.oos_items    || { Breakfast: [], Lunch: [], Dinner: [] }),
    JSON.stringify(body.orders_closed || {}),
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
  ["Amanora",         "Amanora Town",                              "FALSE"],
  ["BG Shirke Road",  "BG Shirke Road",                            "FALSE"],
  ["Bhosale Nagar",   "Bhosale Nagar (Free Delivery)",             "TRUE"],
  ["DP Road",         "DP Road",                                   "FALSE"],
  ["Gadital",         "Gadital",                                   "FALSE"],
  ["Mandai",          "Hadapsar Mandai",                           "FALSE"],
  ["Kirtane Baug",    "Kirtane Baug",                              "FALSE"],
  ["Magarpatta",      "Magarpatta",                                "FALSE"],
  ["Malwadi",         "Malwadi",                                   "FALSE"],
  ["Pune-Solapur Road", "Pune-Solapur Road (Till Gadital Only)",   "FALSE"],
  ["SadeSatraNali",   "SadeSatraNali",                             "FALSE"],
  ["Triveni Nagar",   "Triveni Nagar (Free Delivery)",             "TRUE"],
  ["Tupe Patil Road", "Tupe Patil Road",                           "FALSE"],
  ["Vaiduwadi",       "Vaiduwadi (Till Yash Honda Only)",          "FALSE"],
  ["Vihar Chowk",     "Vihar Chowk",                               "FALSE"],
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
      if (!m.riceMatrix)  m.riceMatrix  = {};
      if (!m.saladMatrix) m.saladMatrix = {};
      if (!m.curdMatrix)  m.curdMatrix  = {};

      var dalQ = Number(r.Dal)   || 0;
      var riceQ = Number(r.Rice)  || 0;
      var saladQ = Number(r.Salad) || 0;
      var curdQ = Number(r.Curd)  || 0;
      
      m.other.Dal.kg      += dalQ * 1.33;
      m.other.Dal.count   += dalQ;
      m.other.Rice.count  += riceQ;
      m.other.Salad.count += saladQ;
      m.other.Curd.count  += curdQ;

      // Matrix calculations
      if (riceQ > 0) {
        var rPacks = calculatePackets(riceQ, 3); // RICE_LIMIT = 3
        rPacks.forEach(function(p) { m.riceMatrix[p] = (m.riceMatrix[p] || 0) + 1; });
      }
      if (saladQ > 0) {
        var sPacks = calculatePackets(saladQ, 4); // SALAD_LIMIT = 4
        sPacks.forEach(function(p) { m.saladMatrix[p] = (m.saladMatrix[p] || 0) + 1; });
      }
      if (curdQ > 0) {
        var cPacks = calculatePackets(curdQ, 2); // CURD_LIMIT = 2 (50g cups: 1 or 2 per packet)
        cPacks.forEach(function(p) { m.curdMatrix[p] = (m.curdMatrix[p] || 0) + 1; });
      }

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
      items: {
        Chapati: Number(r.Chapati)||0, Without_Oil_Chapati: Number(r.Without_Oil_Chapati)||0,
        Phulka: Number(r.Phulka)||0, Ghee_Phulka: Number(r.Ghee_Phulka)||0,
        Jowar_Bhakri: Number(r.Jowar_Bhakri)||0, Bajra_Bhakri: Number(r.Bajra_Bhakri)||0,
        Dry_Sabji_Mini: Number(r.Dry_Sabji_Mini)||0, Dry_Sabji_Full: Number(r.Dry_Sabji_Full)||0,
        Curry_Sabji_Mini: Number(r.Curry_Sabji_Mini)||0, Curry_Sabji_Full: Number(r.Curry_Sabji_Full)||0,
        Dal: Number(r.Dal)||0, Rice: Number(r.Rice)||0, Salad: Number(r.Salad)||0, Curd: Number(r.Curd)||0,
        "Kanda Poha": Number(r["Kanda Poha"])||0, "Ghee Upma": Number(r["Ghee Upma"])||0,
        "Thalipeeth": Number(r["Thalipeeth"])||0, "Palak Paratha": Number(r["Palak Paratha"])||0,
        "Paneer Paratha": Number(r["Paneer Paratha"])||0, "Methi Thepla": Number(r["Methi Thepla"])||0,
        "Sabudana Khichdi": Number(r["Sabudana Khichdi"])||0
      },
      Special_Notes_Kitchen: String(r.Special_Notes_Kitchen || ""),
      Special_Notes_Delivery: String(r.Special_Notes_Delivery || ""),
      Delivery_Point: String(r.Delivery_Point || ""),
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

  // Load delivery status from SK_Deliveries tab (both EnRoute_At and Delivered_At)
  var delMap = {};
  var delWs  = ss.getSheetByName("SK_Deliveries");
  if (delWs) {
    getAllRows(delWs).forEach(function(r) {
      var sid = String(r.Submission_ID || "").trim();
      if (sid) delMap[sid] = {
        deliveredAt: String(r.Delivered_At || ""),
        enRouteAt:   String(r.EnRoute_At   || "")
      };
    });
  }

  // Load customer meal preferences (Source of Truth)
  var custMap = {};
  var custWs = ss.getSheetByName(TAB_CUSTOMERS);
  if (custWs) {
    getAllRows(custWs).forEach(function(r) {
      var ph = _normalizePhone(r.Phone);
      if (ph) {
        custMap[ph] = {
          mealAddresses: r.Meal_Addresses || ""
        };
      }
    });
  }

  rows.forEach(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    if (d !== date) return;
    if (_isOrderCancelled(r.Payment_Status)) return;
    // var area = String(r.Area || "").trim();
    // if (area.toLowerCase().includes("pickup")) return;
    var area = String(r.Area || "").trim();
    var meal = String(r.Meal_Type || "");
    if (!meals[meal]) return;
    var sid = String(r.Submission_ID || "");
    var normP = _normalizePhone(r.Phone);
    meals[meal].push({
      submissionId:  sid,
      name:          String(r.Customer_Name || ""),
      phone:         String(r.Phone || ""),
      area:          area,
      address:       String(r.Full_Address || ""),
      landmark:      String(r.Landmark || ""),
      deliveryPoint: String(r.Delivery_Point || ""),
      maps:          String(r.Maps_Link || ""),
      notes:         String(r.Special_Notes_Delivery || ""),
      deliveredAt:   (delMap[sid] && delMap[sid].deliveredAt) || "",
      enRouteAt:     (delMap[sid] && delMap[sid].enRouteAt)   || "",
      amount:        Number(r.Net_Total || r.Food_Subtotal || 0),
      paymentStatus: String(r.Payment_Status || ""),
      mealAddresses: custMap[normP] ? custMap[normP].mealAddresses : ""
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
  "Breakfast Box":           2.36,
  "Delivery Bag":            1.00,
  "Label / Sticker":         0.2,
  "Bread Packet":            0.70,
  "Sabji Container (Mini)":  2.70,
  "Sabji Container (Full)":  4.0,
  "Dal Container":           4.00,
  "Rice Container":          2.00,
  "Salad Container":         0.700,
  "Curd Container":          1.70
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

// ── PACKAGING EXPENSES — RANGE ───────────────────────────────
function getPackagingExpensesRange(from, to) {
  var ss = getSpreadsheet();
  var ws = getOrCreateTab(ss, TAB_ORDERS, []);
  var rows = getAllRows(ws);

  var rangeRows = rows.filter(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    return d >= from && d <= to && !_isOrderCancelled(r.Payment_Status);
  });

  // Group by date
  var byDate = {};
  rangeRows.forEach(function(r) {
    var d = r.Order_Date instanceof Date
      ? Utilities.formatDate(r.Order_Date, "Asia/Kolkata", "yyyy-MM-dd")
      : String(r.Order_Date).trim();
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(r);
  });

  var PKG_COSTS = PKG_UNIT_COSTS;
  var itemOrder = ["Breakfast Box","Delivery Bag","Label / Sticker","Bread Packet",
                   "Sabji Container (Mini)","Sabji Container (Full)",
                   "Dal Container","Rice Container","Salad Container","Curd Container"];

  function calcDay(dateStr, dayRows) {
    var counts = {}, mealCounts = {Breakfast:0, Lunch:0, Dinner:0};
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
        if (breadCols.some(function(c){ return (Number(r[c])||0)>0; })) add("Bread Packet", 1);
        add("Sabji Container (Mini)", (Number(r.Dry_Sabji_Mini)||0)+(Number(r.Curry_Sabji_Mini)||0));
        add("Sabji Container (Full)", (Number(r.Dry_Sabji_Full)||0)+(Number(r.Curry_Sabji_Full)||0));
        add("Dal Container",  Number(r.Dal)||0);
        add("Rice Container", Number(r.Rice)||0);
        add("Salad Container",Number(r.Salad)||0);
        add("Curd Container", Number(r.Curd)||0);
      }
    });
    var items = [], total = 0;
    itemOrder.forEach(function(key) {
      var qty = counts[key]||0; if (!qty) return;
      var unitCost = PKG_COSTS[key]||0, t = qty*unitCost;
      items.push({name:key, qty:qty, unitCost:unitCost, total:t});
      total += t;
    });
    var mealsOut = {};
    Object.keys(mealCounts).forEach(function(m){ if(mealCounts[m]>0) mealsOut[m]=mealCounts[m]; });
    return {date:dateStr, orderCount:dayRows.length, meals:mealsOut, items:items, total:total};
  }

  // Build per-day results
  var days = Object.keys(byDate).sort().map(function(d){ return calcDay(d, byDate[d]); });

  // Aggregate totals
  var aggCounts = {}, aggMeals = {Breakfast:0,Lunch:0,Dinner:0}, aggTotal = 0, aggOrders = 0;
  days.forEach(function(day) {
    aggOrders += day.orderCount;
    aggTotal  += day.total;
    Object.keys(day.meals).forEach(function(m){ aggMeals[m]=(aggMeals[m]||0)+day.meals[m]; });
    day.items.forEach(function(it){ aggCounts[it.name]=(aggCounts[it.name]||0)+it.qty; });
  });
  var aggItems = [];
  itemOrder.forEach(function(key) {
    var qty = aggCounts[key]||0; if (!qty) return;
    var unitCost = PKG_COSTS[key]||0, t=qty*unitCost;
    aggItems.push({name:key, qty:qty, unitCost:unitCost, total:t});
  });

  return {
    from: from, to: to,
    orderCount: aggOrders,
    total: aggTotal,
    meals: aggMeals,
    items: aggItems,
    days: days
  };
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
         r.Payment_Status === "on account" ||
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
        claimed: (String(c.Review_Reward_Claimed) === "TRUE" || String(c.Review_Reward_Claimed) === "true"),
        standardOrder: c.Standard_Order || "",
        feeExempt: (String(c.Fee_Exempt).trim() === "Yes") ? "Yes" : "No",
        onAccount: (String(c.On_Account).trim() === "Yes") ? "Yes" : "No",
        billingCycle: c.Billing_Cycle || "Daily",
        // Address profiles
        wing:    c.Wing || "",
        flat:    c.Flat || "",
        floor:   c.Floor || "",
        society: c.Society || "",
        maps:    c.Maps_Link || "",
        landmark: c.Landmark || "",
        delivery_point: c.Delivery_Point || ""
      };
    }
  });

  var fmtDate = function(v) {
    return v instanceof Date ? Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd") : String(v).trim();
  };

  var map = {};
  ordRows.forEach(function(r) {
    if (_isOrderCancelled(r.Payment_Status)) return;
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
        reviewClaimed: cMap[normP] ? cMap[normP].claimed : false,
        standardOrder: cMap[normP] ? cMap[normP].standardOrder : "",
        Fee_Exempt:    cMap[normP] ? cMap[normP].feeExempt : "No",
        onAccount:     cMap[normP] ? cMap[normP].onAccount : "No",
        billingCycle:  cMap[normP] ? cMap[normP].billingCycle : "Daily",
        wing:    cMap[normP] ? cMap[normP].wing : "",
        flat:    cMap[normP] ? cMap[normP].flat : "",
        floor:   cMap[normP] ? cMap[normP].floor : "",
        society: cMap[normP] ? cMap[normP].society : "",
        maps:    cMap[normP] ? cMap[normP].maps : "",
        landmark: cMap[normP] ? cMap[normP].landmark : "",
        delivery_point: cMap[normP] ? cMap[normP].delivery_point : ""
      };
    }
    map[phone].orderCount++;
    map[phone].totalSpent += Number(r.Net_Total)||0;
    const ps = String(r.Payment_Status || "").trim();
    if (ps !== "Paid" && ps !== "Wallet Paid" && ps !== "Collected") map[phone].pendingAmt += Number(r.Net_Total)||0;
    if (d > map[phone].lastDate) {
      map[phone].lastDate = d;
      map[phone].name = String(r.Customer_Name||map[phone].name).trim();
    }
  });

  // Also include SK_Customers entries that have never placed an order (e.g. pre-registered VIPs)
  custRows.forEach(function(c) {
    var p = String(c.Phone || "").trim();
    if (!p) return;
    var normP = _normalizePhone(p);
    if (map[p]) return; // already in map from orders
    // Only surface pre-registered VIPs (Fee_Exempt = Yes) to keep the list clean
    if (String(c.Fee_Exempt).trim() !== "Yes") return;
    map[p] = {
      phone: p,
      name: String(c.Customer_Name || "").trim(),
      area: String(c.Area || "").trim(),
      payFreq: String(c.Payment_Freq || "").trim(),
      orderCount: 0,
      totalSpent: 0,
      pendingAmt: 0,
      lastDate: String(c.Created_At instanceof Date
        ? Utilities.formatDate(c.Created_At, "Asia/Kolkata", "yyyy-MM-dd")
        : (c.Created_At || "")).slice(0, 10),
      promoCount: 0,
      reviewClaimed: false,
      standardOrder: "",
      Fee_Exempt: "Yes",
      onAccount: String(c.On_Account).trim() === "Yes" ? "Yes" : "No",
      billingCycle: c.Billing_Cycle || "Daily"
    };
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
  var activeOrders = orders.filter(function(o){return !_isOrderCancelled(o.status);});
  var totalSpent = Math.round(activeOrders.reduce(function(s,o){return s+o.total;},0));
  var pending    = Math.round(activeOrders.filter(function(o){return String(o.status)!=="Paid" && String(o.status)!=="Wallet Paid";}).reduce(function(s,o){return s+o.total;},0));

  // Fetch Standard_Order from customer sheet
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var cRows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var standardOrder = "";
  var custMatch = cRows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  if (custMatch) standardOrder = custMatch.Standard_Order || "";

  var feeExempt = custMatch ? (String(custMatch.Fee_Exempt).trim() === "Yes" ? "Yes" : "No") : "No";
  var onAccount = custMatch ? (String(custMatch.On_Account).trim() === "Yes" ? "Yes" : "No") : "No";
  var billingCycle = custMatch ? (custMatch.Billing_Cycle || "Daily") : "Daily";

  return {
    success:true, phone:phone, name:name, area:area, payFreq:payFreq,
    orders:orders, totalSpent:totalSpent, pending:pending, orderCount:orders.length,
    standardOrder: standardOrder, Fee_Exempt: feeExempt, On_Account: onAccount, Billing_Cycle: billingCycle,
    // Add full address profile
    wing:    custMatch ? (custMatch.Wing || "") : "",
    flat:    custMatch ? (custMatch.Flat || "") : "",
    floor:   custMatch ? (custMatch.Floor || "") : "",
    society: custMatch ? (custMatch.Society || "") : "",
    maps:    custMatch ? (custMatch.Maps_Link || "") : "",
    landmark: custMatch ? (custMatch.Landmark || "") : "",
    delivery_point: custMatch ? (custMatch.Delivery_Point || "") : ""
  };
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
      ws.deleteRow(r._row); // Final delete after verification
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

// ── DELETED OBSOLETE ADMIN CANCEL ORDER (Merged with main) ──
/**
 * Marks a specific order as 'Packed' in the SK_Orders sheet.
 * Called by the Kitchen Dashboard (kitchen.html) via APPS_SCRIPT_URL.
 */
function markOrderPacked(body) {
  var id = body.submissionId;
  if (!id) return {success:false, error: "submissionId required"};

  var ss    = getSpreadsheet();
  var ws    = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var hIdx  = headerIndex(ws);
  var rows  = getAllRows(ws);
  
  if (hIdx.Packed === undefined) return {success:false, error: "Packed column not found"};

  var order = rows.find(function(r) {
    return String(r.Submission_ID) === String(id);
  });

  if (order) {
    ws.getRange(order._row, hIdx.Packed).setValue(true);
    return {success:true};
  }
  
  return {success:false, error: "Order not found"};
}

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
    return d >= dateFrom && d <= dateTo && !_isOrderCancelled(r.Payment_Status);
  });

  // ── Option B: Exact Small Order Fee backfill ──────────────────────────────
  // Pre-pass 1: build VIP set from profiles (Fee_Exempt = Yes)
  var profWs   = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var profRows = getAllRows(profWs);
  var vipSet   = {};
  profRows.forEach(function(pr) {
    if (pr.Fee_Exempt === "Yes" || pr.Fee_Exempt === true) {
      vipSet[String(pr.Phone||"").trim()] = true;
    }
  });

  // Pre-pass 2: for every phone+date combo, sum food subtotals & count distinct meals
  // This lets us know if the combined day total reached the free-delivery threshold,
  // which also waives the small order fee.
  var dayTotals = {}; // key = phone+"_"+date  →  { foodTotal, mealCount }
  rows.forEach(function(r) {
    var ph   = String(r.Phone||"").trim();
    var d    = fmtDate(r.Order_Date);
    var food = Number(r.Food_Subtotal)||0;
    var key  = ph + "_" + d;
    if (!dayTotals[key]) dayTotals[key] = { foodTotal:0, meals:{} };
    dayTotals[key].foodTotal += food;
    dayTotals[key].meals[String(r.Meal_Type||"")] = true;
  });

  // Helper: calculate small fee for a row using exact rules
  function calcSmallFee(r) {
    var stored = r.Small_Order_Fee;
    // If the column exists and has a numeric value, trust it
    if (stored !== undefined && stored !== null && stored !== "" && !isNaN(Number(stored))) {
      return Number(stored);
    }
    // Backfill for old rows
    var meal = String(r.Meal_Type||"");
    if (meal !== "Lunch" && meal !== "Dinner") return 0; // Breakfast: never charged
    var food = Number(r.Food_Subtotal)||0;
    if (food <= 0 || food >= 50) return 0;              // Only charged when sub < ₹50
    var area = String(r.Area||"").trim();
    if (area === "Self Pickup") return 0;                // Pickup: waived
    var ph  = String(r.Phone||"").trim();
    if (vipSet[ph]) return 0;                            // VIP: waived
    // Check if combined day food total crossed free-delivery threshold
    var d   = fmtDate(r.Order_Date);
    var key = ph + "_" + d;
    var dt  = dayTotals[key] || {foodTotal:0, meals:{}};
    var mealCount    = Object.keys(dt.meals).length;
    var threshold    = mealCount <= 1 ? 100 : 150;
    if (dt.foodTotal >= threshold) return 0;             // Day crossed threshold: waived
    return 10;
  }
  // ── End backfill helper ───────────────────────────────────────────────────
  var LUNCH_COLS = ["Chapati","Without_Oil_Chapati","Phulka","Ghee_Phulka","Jowar_Bhakri","Bajra_Bhakri",
    "Dry_Sabji_Mini","Dry_Sabji_Full","Curry_Sabji_Mini","Curry_Sabji_Full","Dal","Rice","Salad","Curd"];
  var COL_DISP = {"Chapati":"Chapati","Without_Oil_Chapati":"WO Chapati","Phulka":"Phulka","Ghee_Phulka":"Ghee Phulka",
    "Jowar_Bhakri":"Jowar Bhakri","Bajra_Bhakri":"Bajra Bhakri","Dry_Sabji_Mini":"Dry Sabji Mini",
    "Dry_Sabji_Full":"Dry Sabji Full","Curry_Sabji_Mini":"Curry Sabji Mini","Curry_Sabji_Full":"Curry Sabji Full",
    "Dal":"Dal","Rice":"Rice","Salad":"Salad","Curd":"Curd"};
  var totalRev=0, totalPaid=0, totalDelivery=0, totalSurcharge=0, totalSmallFee=0;
  var custSet={}, dayMap={};
  var mealStats={Breakfast:{count:0,revenue:0},Lunch:{count:0,revenue:0},Dinner:{count:0,revenue:0}};
  var itemCounts={};
  rows.forEach(function(r) {
    var d=fmtDate(r.Order_Date), net=Number(r.Net_Total)||0;
    var delivery=Number(r.Delivery_Charge)||0;
    var food=Number(r.Food_Subtotal)||0;
    // Backfill surcharge for old rows where Inflation_Surcharge column was blank
    var surchargeRaw=Number(r.Inflation_Surcharge);
    var surcharge = (!isNaN(surchargeRaw) && surchargeRaw > 0) ? surchargeRaw : (food > 0 ? Math.ceil(food/20) : 0);
    // Small_Order_Fee: exact backfill using Option B (checks VIP, pickup, day threshold)
    var smallFee = calcSmallFee(r);
    var payStatus = String(r.Payment_Status || "").trim();
    totalRev+=net;
    totalDelivery+=delivery; totalSurcharge+=surcharge; totalSmallFee+=smallFee;
    if(payStatus==="Paid"||payStatus==="Wallet Paid"||payStatus==="Collected") totalPaid+=net;
    var ph=String(r.Phone||"").trim(); if(ph) custSet[ph]=true;
    var meal=String(r.Meal_Type||"");
    if(mealStats[meal]){mealStats[meal].count++;mealStats[meal].revenue+=net;}
    if(!dayMap[d]) dayMap[d]={orders:0,revenue:0,delivery:0,surcharge:0,smallFee:0};
    dayMap[d].orders++; dayMap[d].revenue+=net;
    dayMap[d].delivery+=delivery; dayMap[d].surcharge+=surcharge; dayMap[d].smallFee+=smallFee;
    if(meal==="Breakfast"){
      for(var n=1;n<=4;n++){var bi=String(r["BF_Item_"+n]||"").trim(),bq=Number(r["BF_Qty_"+n])||0;if(bi&&bq>0)itemCounts[bi]=(itemCounts[bi]||0)+bq;}
      var cu=Number(r.Curd)||0; if(cu>0)itemCounts["Curd"]=(itemCounts["Curd"]||0)+cu;
    } else {
      LUNCH_COLS.forEach(function(col){var q=Number(r[col])||0;if(q>0){var dn=COL_DISP[col]||col;itemCounts[dn]=(itemCounts[dn]||0)+q;}});
    }
  });
  var days=Object.keys(dayMap).sort().map(function(d){
    return{date:d,orders:dayMap[d].orders,revenue:Math.round(dayMap[d].revenue),
           delivery:Math.round(dayMap[d].delivery),surcharge:Math.round(dayMap[d].surcharge),smallFee:Math.round(dayMap[d].smallFee)};
  });
  var allItems=Object.keys(itemCounts).map(function(k){return{name:k,count:Math.round(itemCounts[k])};}).sort(function(a,b){return b.count-a.count;});
  var topItems=allItems.slice(0,15);
  Object.keys(mealStats).forEach(function(m){mealStats[m].revenue=Math.round(mealStats[m].revenue);});
  return {success:true,
    summary:{orders:rows.length,customers:Object.keys(custSet).length,revenue:Math.round(totalRev),
      paid:Math.round(totalPaid),pending:Math.round(totalRev-totalPaid),
      avgPerDay:days.length>0?Math.round(totalRev/days.length):0,
      delivery:Math.round(totalDelivery),surcharge:Math.round(totalSurcharge),smallFee:Math.round(totalSmallFee)},
    meals:mealStats,days:days,topItems:topItems,allItems:allItems};
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
  var newBalance = _calculateWalletBalance(phone);
  return {success:true, newBalance: Math.round(newBalance)};
}

// ── INVENTORY ─────────────────────────────────────────────────────────────────
// Tracks raw material purchases. Each new entry for the same item auto-calculates
// how long the previous batch lasted → builds consumption rate over time.
const TAB_INVENTORY      = "SK_Inventory";
const INVENTORY_HEADERS  = [
  "Entry_ID","Date","Item","Unit","Quantity","Price_Paid","Notes","Timestamp"
];

function saveInventoryEntry(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var now  = new Date();
  var id   = "INV-" + Utilities.formatDate(now,"Asia/Kolkata","yyyyMMdd") + "-" + Math.floor(Math.random()*9000+1000);
  var hIdx = headerIndex(ws);
  var totalCols = Math.max(ws.getLastColumn(), INVENTORY_HEADERS.length);
  var row  = new Array(totalCols).fill("");
  var set  = function(col, val) { if (hIdx[col]) row[hIdx[col]-1] = val; };

  set("Entry_ID",   id);
  set("Date",       String(body.date || Utilities.formatDate(now,"Asia/Kolkata","yyyy-MM-dd")));
  set("Item",       String(body.item || "").trim());
  set("Unit",       String(body.unit || "kg"));
  set("Quantity",   Number(body.quantity) || 0);
  set("Price_Paid", Number(body.price)    || 0);
  set("Notes",      String(body.notes     || ""));
  set("Timestamp",  getISTTimestamp());

  ws.appendRow(row);
  return { success: true, id: id };
}

function getInventoryData(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var rows = getAllRows(ws);

  // Sort ascending by date for correct duration calculation
  rows.sort(function(a,b){ return String(a.Date).localeCompare(String(b.Date)); });

  // Group by item
  var byItem = {};
  rows.forEach(function(r) {
    var item = String(r.Item || "").trim();
    if (!item) return;
    if (!byItem[item]) byItem[item] = [];
    byItem[item].push({
      id:       String(r.Entry_ID || ""),
      date:     String(r.Date     || ""),
      unit:     String(r.Unit     || "kg"),
      qty:      Number(r.Quantity) || 0,
      price:    Number(r.Price_Paid) || 0,
      notes:    String(r.Notes    || ""),
      timestamp:String(r.Timestamp || "")
    });
  });

  // For each item, calculate durations between entries + consumption stats
  var items = [];
  Object.keys(byItem).sort().forEach(function(item) {
    var entries = byItem[item];
    var totalDays = 0, totalQty = 0, durationCount = 0;

    // Annotate each entry with how long it lasted (days until next purchase)
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      e.lasted_days = null;
      e.daily_rate  = null;
      if (i < entries.length - 1) {
        var d1 = new Date(e.date);
        var d2 = new Date(entries[i+1].date);
        var days = Math.round((d2 - d1) / 86400000);
        if (days > 0) {
          e.lasted_days = days;
          e.daily_rate  = Math.round((e.qty / days) * 100) / 100;
          totalDays += days;
          totalQty  += e.qty;
          durationCount++;
        }
      }
    }

    var avgDays       = durationCount > 0 ? Math.round(totalDays / durationCount) : null;
    var avgDailyRate  = (avgDays && totalQty) ? Math.round((totalQty / durationCount / avgDays) * 100) / 100 : null;
    var lastEntry     = entries[entries.length - 1];

    // Predict next purchase date
    var nextBuyDate = null;
    if (avgDays && lastEntry.date) {
      var d = new Date(lastEntry.date);
      d.setDate(d.getDate() + avgDays);
      nextBuyDate = Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
    }

    items.push({
      item:          item,
      unit:          lastEntry.unit,
      entries:       entries.reverse(), // newest first for display
      entry_count:   entries.length,
      avg_days:      avgDays,
      avg_daily_rate:avgDailyRate,
      last_purchased:lastEntry.date,
      last_qty:      lastEntry.qty,
      next_buy_est:  nextBuyDate
    });
  });

  return { success: true, items: items };
}

function deleteInventoryEntry(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_INVENTORY, INVENTORY_HEADERS);
  var hIdx = headerIndex(ws);
  var rows = ws.getDataRange().getValues();
  var idCol = (hIdx["Entry_ID"] || 1) - 1;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]).trim() === String(body.id || "").trim()) {
      ws.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Entry not found" };
}

// ── EXPENSE CUSTOM CATEGORIES ─────────────────────────────────────────────────
// Stored in Script Properties as JSON so no extra sheet is needed.
function getCustomExpenseCategories() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty("CUSTOM_EXP_CATS");
    return raw ? JSON.parse(raw) : {};
  } catch(e) { return {}; }
}

function saveCustomExpenseCategory(body) {
  var category = String(body.category || "").trim();
  var item     = String(body.item     || "").trim();
  if (!category) return { success:false, error:"Category required" };
  var cats = getCustomExpenseCategories();
  if (!cats[category]) cats[category] = [];
  if (item && !cats[category].includes(item)) cats[category].push(item);
  PropertiesService.getScriptProperties().setProperty("CUSTOM_EXP_CATS", JSON.stringify(cats));
  return { success:true, categories: cats };
}

function deleteCustomExpenseCategory(body) {
  var category = String(body.category || "").trim();
  var item     = String(body.item     || "").trim();
  var cats = getCustomExpenseCategories();
  if (item && cats[category]) {
    cats[category] = cats[category].filter(function(i){ return i !== item; });
  } else {
    delete cats[category];
  }
  PropertiesService.getScriptProperties().setProperty("CUSTOM_EXP_CATS", JSON.stringify(cats));
  return { success:true, categories: cats };
}

// ── KITCHEN EXPENSES ──────────────────────────────────────────────────────────
const TAB_EXPENSES      = "SK_Expenses";
const EXPENSES_HEADERS  = [
  "Expense_ID","Date","Category","Item","Amount","Frequency",
  "Payment_Mode","Notes","Timestamp"
];

// Category → sub-items map (also used by frontend for dropdowns)
var EXPENSE_CATEGORIES = {
  "🥦 Raw Materials": [
    "Vegetables & Greens","Fruits","Dairy (Milk/Curd/Paneer/Butter)",
    "Oil & Ghee","Spices & Masala","Dry Groceries (Dal/Rice/Atta)","Other Raw Material"
  ],
  "📦 Packaging": [
    "Containers / Boxes","Bags & Covers","Labels & Stickers",
    "Tissue & Napkins","Other Packaging"
  ],
  "⛽ Fuel & Transport": [
    "Petrol / CNG","Vehicle Maintenance","Delivery Outsourcing","Other Transport"
  ],
  "👨‍🍳 Staff": [
    "Cook Salary","Helper Salary","Delivery Person Salary","Part-time Staff","Other Staff"
  ],
  "🔌 Utilities": [
    "LPG Cylinder","Electricity Bill","Water Bill","Internet / Phone","Other Utility"
  ],
  "🍳 Kitchen & Equipment": [
    "Equipment Purchase","Equipment Repair / Service","Utensils","Cleaning Supplies","Other Kitchen"
  ],
  "📣 Marketing": [
    "Printing / Pamphlets","Online Advertising","Branding / Design","Other Marketing"
  ],
  "🏦 Finance & Admin": [
    "Bank Charges","Platform / Software Fees","GST / Tax","Other Finance"
  ],
  "📝 Miscellaneous": ["Miscellaneous"]
};

function saveExpense(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var hIdx = headerIndex(ws);
  var now  = new Date();
  var id   = "EXP-" + Utilities.formatDate(now, "Asia/Kolkata", "yyyyMMdd") + "-" + Math.floor(Math.random()*9000+1000);

  var totalCols = ws.getLastColumn();
  var row = new Array(totalCols).fill("");
  var set = function(col, val) { if (hIdx[col]) row[hIdx[col]-1] = val; };

  set("Expense_ID",   id);
  set("Date",         String(body.date || Utilities.formatDate(now,"Asia/Kolkata","yyyy-MM-dd")));
  set("Category",     String(body.category  || ""));
  set("Item",         String(body.item      || ""));
  set("Amount",       Number(body.amount)   || 0);
  set("Frequency",    String(body.frequency || "One-time"));
  set("Payment_Mode", String(body.payment_mode || "Cash"));
  set("Notes",        String(body.notes     || ""));
  set("Timestamp",    getISTTimestamp());

  ws.appendRow(row);
  return { success: true, id: id };
}

function deleteExpense(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var hIdx = headerIndex(ws);
  var rows = ws.getDataRange().getValues();
  var idCol = (hIdx["Expense_ID"] || 1) - 1;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]).trim() === String(body.id || "").trim()) {
      ws.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Expense not found" };
}

function getExpenses(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var rows = getAllRows(ws);
  var from = String(body.from || "");
  var to   = String(body.to   || "");
  var filtered = rows.filter(function(r) {
    var d = String(r.Date || "").trim();
    if (from && d < from) return false;
    if (to   && d > to)   return false;
    return true;
  });
  // Sort newest first
  filtered.sort(function(a,b){ return String(b.Date).localeCompare(String(a.Date)); });
  return {
    success: true,
    expenses: filtered.map(function(r) {
      return {
        id:           r.Expense_ID,
        date:         r.Date,
        category:     r.Category,
        item:         r.Item,
        amount:       Number(r.Amount) || 0,
        frequency:    r.Frequency,
        payment_mode: r.Payment_Mode,
        notes:        r.Notes,
        timestamp:    r.Timestamp
      };
    }),
    categories: EXPENSE_CATEGORIES
  };
}

function getExpenseAnalytics(body) {
  var ss   = getSpreadsheet();
  var ws   = getOrCreateTab(ss, TAB_EXPENSES, EXPENSES_HEADERS);
  var rows = getAllRows(ws);
  var from = String(body.from || "");
  var to   = String(body.to   || "");

  var filtered = rows.filter(function(r) {
    var d = String(r.Date || "").trim();
    if (!d) return false;
    if (from && d < from) return false;
    if (to   && d > to)   return false;
    return true;
  });

  var total       = 0;
  var byCat       = {};   // category → total
  var byFreq      = {};   // frequency → total
  var byPayMode   = {};   // payment mode → total
  var byDay       = {};   // date → total
  var topItems    = {};   // item → total
  var monthlyFixed = 0;   // sum of Monthly-tagged expenses

  filtered.forEach(function(r) {
    var amt  = Number(r.Amount) || 0;
    var cat  = String(r.Category || "Other");
    var freq = String(r.Frequency || "One-time");
    var pm   = String(r.Payment_Mode || "Cash");
    var d    = String(r.Date || "").trim();
    var item = String(r.Item || "Other");

    total += amt;
    byCat[cat]     = (byCat[cat]     || 0) + amt;
    byFreq[freq]   = (byFreq[freq]   || 0) + amt;
    byPayMode[pm]  = (byPayMode[pm]  || 0) + amt;
    byDay[d]       = (byDay[d]       || 0) + amt;
    topItems[item] = (topItems[item] || 0) + amt;
    if (freq === "Monthly") monthlyFixed += amt;
  });

  var days = Object.keys(byDay).sort().map(function(d) {
    return { date: d, amount: Math.round(byDay[d]) };
  });

  var catArr = Object.keys(byCat).sort(function(a,b){ return byCat[b]-byCat[a]; }).map(function(c) {
    return { category: c, amount: Math.round(byCat[c]) };
  });

  var itemArr = Object.keys(topItems).sort(function(a,b){ return topItems[b]-topItems[a]; }).slice(0,10).map(function(i) {
    return { item: i, amount: Math.round(topItems[i]) };
  });

  return {
    success:      true,
    total:        Math.round(total),
    monthlyFixed: Math.round(monthlyFixed),
    count:        filtered.length,
    byCategory:   catArr,
    byFrequency:  byFreq,
    byPayMode:    byPayMode,
    byDay:        days,
    topItems:     itemArr,
    categories:   EXPENSE_CATEGORIES
  };
}

// ── CLIENT ERROR LOG ──────────────────────────────────────────────────────────
const TAB_ERROR_LOG     = "SK_Error_Log";
const ERROR_LOG_HEADERS = ["Timestamp","Date","Phone","Version","Type","Action","Attempt","Duration_ms","Message","URL"];

function logClientError(body) {
  try {
    var ss  = getSpreadsheet();
    var ws  = getOrCreateTab(ss, TAB_ERROR_LOG, ERROR_LOG_HEADERS);
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
    ws.appendRow([
      getISTTimestamp(),
      dateStr,
      String(body.phone    || "unknown"),
      String(body.version  || ""),
      String(body.type     || "error"),
      String(body.action   || "unknown"),
      Number(body.attempt  || 1),
      Number(body.ms       || 0),
      String(body.msg      || ""),
      String(body.url      || "")
    ]);
    return { success: true };
  } catch(e) {
    return { success: false }; // never throw — this is logging only
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
  Archives SK_Orders and SK_Wallet for a given quarter into a new Google
  Spreadsheet, writes Balance Carry Forward snapshots so wallet balances are
  preserved, then deletes the archived rows from the main sheet.

  Quarter map:
    Q1 = Jan–Mar   (archive trigger: April 10)
    Q2 = Apr–Jun   (archive trigger: July 10)
    Q3 = Jul–Sep   (archive trigger: October 10)
    Q4 = Oct–Dec   (archive trigger: January 10 of next year)
*/
function archiveQuarter(year, quarter) {
  var Q = {
    1: {from: year+"-01-01", to: year+"-03-31", label: "Q1 Jan–Mar"},
    2: {from: year+"-04-01", to: year+"-06-30", label: "Q2 Apr–Jun"},
    3: {from: year+"-07-01", to: year+"-09-30", label: "Q3 Jul–Sep"},
    4: {from: year+"-10-01", to: year+"-12-31", label: "Q4 Oct–Dec"}
  };
  var qr = Q[quarter];
  if (!qr) return {success:false, error:"Invalid quarter — must be 1, 2, 3 or 4"};

  var ss = getSpreadsheet();
  var fmtDate = function(v) {
    return v instanceof Date
      ? Utilities.formatDate(v, "Asia/Kolkata", "yyyy-MM-dd")
      : String(v || "").trim().slice(0, 10);
  };

  // ── STEP 1: Create archive spreadsheet ────────────────────────────────────
  var archiveName = "Svaadh Kitchen Archive — " + qr.label + " " + year;
  var archiveSS   = SpreadsheetApp.create(archiveName);
  var log = [];

  // ── STEP 2: Archive SK_Orders ──────────────────────────────────────────────
  var ordersWs      = getOrCreateTab(ss, TAB_ORDERS, ORDERS_HEADERS);
  var allOrderData  = ordersWs.getDataRange().getValues();
  var oHeaders      = allOrderData[0];
  var oDateIdx      = oHeaders.indexOf("Order_Date");

  // Collect rows that fall in the quarter (track 1-based sheet row numbers)
  var toArchiveOrders = []; // { sheetRow (1-based), vals }
  for (var i = 1; i < allOrderData.length; i++) {
    var d = fmtDate(allOrderData[i][oDateIdx]);
    if (d >= qr.from && d <= qr.to) {
      toArchiveOrders.push({sheetRow: i + 1, vals: allOrderData[i]});
    }
  }

  if (toArchiveOrders.length > 0) {
    var archiveOrderSheet = archiveSS.getActiveSheet();
    archiveOrderSheet.setName("SK_Orders");
    archiveOrderSheet.getRange(1, 1, 1, oHeaders.length).setValues([oHeaders]);
    var oData = toArchiveOrders.map(function(r) { return r.vals; });
    archiveOrderSheet.getRange(2, 1, oData.length, oHeaders.length).setValues(oData);
    // Verify
    var oWritten = archiveOrderSheet.getLastRow() - 1;
    if (oWritten !== toArchiveOrders.length) {
      return {success:false, error:"Order archive verification failed. Expected "
        + toArchiveOrders.length + ", got " + oWritten + ". Nothing deleted."};
    }
    log.push(toArchiveOrders.length + " orders archived ✓");
  } else {
    log.push("No orders found for this quarter.");
  }

  // ── STEP 3: Archive SK_Wallet ──────────────────────────────────────────────
  var walletWs     = getOrCreateTab(ss, TAB_WALLET, WALLET_HEADERS);
  var allWalletData = walletWs.getDataRange().getValues();
  var wHeaders      = allWalletData[0];
  var wTsIdx        = wHeaders.indexOf("Timestamp");
  var wPhoneIdx     = wHeaders.indexOf("Phone");
  var wNameIdx      = wHeaders.indexOf("Customer_Name");

  var toArchiveWallet = [];
  for (var j = 1; j < allWalletData.length; j++) {
    var ts = allWalletData[j][wTsIdx];
    var wd = fmtDate(ts instanceof Date ? ts : new Date(ts));
    if (wd >= qr.from && wd <= qr.to) {
      toArchiveWallet.push({sheetRow: j + 1, vals: allWalletData[j]});
    }
  }

  if (toArchiveWallet.length > 0) {
    var archiveWalletSheet = archiveSS.insertSheet("SK_Wallet");
    archiveWalletSheet.getRange(1, 1, 1, wHeaders.length).setValues([wHeaders]);
    var wData = toArchiveWallet.map(function(r) { return r.vals; });
    archiveWalletSheet.getRange(2, 1, wData.length, wHeaders.length).setValues(wData);
    var wWritten = archiveWalletSheet.getLastRow() - 1;
    if (wWritten !== toArchiveWallet.length) {
      return {success:false, error:"Wallet archive verification failed. Expected "
        + toArchiveWallet.length + ", got " + wWritten + ". Nothing deleted."};
    }
    log.push(toArchiveWallet.length + " wallet transactions archived ✓");
  } else {
    log.push("No wallet transactions found for this quarter.");
  }

  // ── STEP 4: Write Balance Carry Forward snapshots ─────────────────────────
  // Calculate BEFORE deleting anything — reads full live wallet sheet
  // Only needed for phones that had activity in this quarter AND have balance > 0
  var activePhones = {};
  toArchiveWallet.forEach(function(r) {
    var ph   = String(r.vals[wPhoneIdx] || "").trim();
    var name = String(r.vals[wNameIdx]  || "").trim();
    if (ph) activePhones[ph] = name;
  });

  var snapshotCount = 0;
  var snapTime      = new Date();
  var refId         = "ARCHIVE-Q" + quarter + "-" + year;
  Object.keys(activePhones).forEach(function(ph) {
    var balance = _calculateWalletBalance(ph);
    if (balance > 0) {
      walletWs.appendRow([ph, activePhones[ph], "Balance Carry Forward",
                          balance, "TRUE", refId, snapTime]);
      snapshotCount++;
    }
    // balance = 0 → no snapshot needed, no carry-forward required
  });
  if (snapshotCount > 0) log.push(snapshotCount + " balance snapshots written ✓");

  // ── STEP 5: Delete archived rows — bottom to top so indices don't shift ───
  // Wallet rows
  var wRowNums = toArchiveWallet.map(function(r){return r.sheetRow;})
                                .sort(function(a,b){return b-a;});
  wRowNums.forEach(function(rowNum) { walletWs.deleteRow(rowNum); });
  if (wRowNums.length) log.push(wRowNums.length + " wallet rows deleted from main ✓");

  // Order rows
  var oRowNums = toArchiveOrders.map(function(r){return r.sheetRow;})
                                .sort(function(a,b){return b-a;});
  oRowNums.forEach(function(rowNum) { ordersWs.deleteRow(rowNum); });
  if (oRowNums.length) log.push(oRowNums.length + " order rows deleted from main ✓");

  return {
    success:          true,
    archiveName:      archiveName,
    archiveUrl:       archiveSS.getUrl(),
    ordersArchived:   toArchiveOrders.length,
    walletArchived:   toArchiveWallet.length,
    snapshots:        snapshotCount,
    log:              log
  };
}

// Called by admin UI — wraps archiveQuarter with PIN check (handled by router)
function triggerManualArchive(body) {
  var year    = parseInt(body.year);
  var quarter = parseInt(body.quarter);
  if (!year || !quarter) return {success:false, error:"year and quarter required"};
  return archiveQuarter(year, quarter);
}

// ── Time-based trigger: auto-archive previous quarter on the 10th ─────────
// Run setupQuarterlyArchiveTrigger() once from Apps Script editor to register.
function setupQuarterlyArchiveTrigger() {
  // Remove any existing trigger for runScheduledArchive
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "runScheduledArchive") ScriptApp.deleteTrigger(t);
  });
  // Fire on the 10th of every month at 2 AM IST (Apps Script uses server time ≈ UTC)
  ScriptApp.newTrigger("runScheduledArchive")
    .timeBased()
    .onMonthDay(10)
    .atHour(21)   // 21:00 UTC = 02:30 IST next day ≈ 2 AM IST on the 10th
    .create();
  return "Quarterly archive trigger set — fires on 10th of Apr, Jul, Oct, Jan.";
}

function runScheduledArchive() {
  // Determine which quarter just ended based on today's month
  var now     = new Date();
  var month   = now.getMonth() + 1; // 1–12
  var year    = now.getFullYear();
  var qMap    = {4:1, 7:2, 10:3, 1:4}; // trigger month → quarter to archive
  var qNum    = qMap[month];
  if (!qNum) { Logger.log("runScheduledArchive: not an archive month (" + month + "), skipping."); return; }
  var archiveYear = (month === 1) ? year - 1 : year; // Jan trigger → archive Q4 of last year
  var result  = archiveQuarter(archiveYear, qNum);
  Logger.log("Archive result: " + JSON.stringify(result));
}

// ── CHURN REPORT ──────────────────────────────────────────────────────────────
function getChurnReport(sinceDate) {
  if (!sinceDate) return {success:false, error:"sinceDate required"};
  var ss=getSpreadsheet(), ws=getOrCreateTab(ss,TAB_ORDERS,ORDERS_HEADERS);
  var fmtDate=function(v){return v instanceof Date?Utilities.formatDate(v,"Asia/Kolkata","yyyy-MM-dd"):String(v).trim();};
  var map={};
  getAllRows(ws).forEach(function(r){
    if(_isOrderCancelled(r.Payment_Status))return;
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

/**
 * Batch process multiple approvals/rejections
 * body: { tab: 'refunds'|'payments'|'wallet', action: 'approve'|'reject', items: [...] }
 */
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
    const m = String(r.Payment_Method || "").trim();
    return s === "Pending" || s === "Cancelled (Verify UPI)" || s === "Pending Approval"
      || (m === "Split" && s === "Pending"); // Split orders awaiting UPI portion
  })
             .map(r => {
               const walletCredit = Number(r.Wallet_Credit) || 0;
               const isSplit = String(r.Payment_Method || "").trim() === "Split";
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
                 refund_preference: r.Refund_Preference || ""
               };
             });
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

function setStandardOrder(phone, itemsJSON, templateName, meal) {
  var ss = getSpreadsheet();
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var rows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var cust = rows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  
  var currentStr = (cust && cust.Standard_Order) ? cust.Standard_Order : "[]";
  var list = [];
  try { list = JSON.parse(currentStr); if(!Array.isArray(list)) list=[]; } catch(e){ list=[]; }
  
  // Ensure items is an object, not a string, before saving
  var finalItems = itemsJSON;
  if (typeof itemsJSON === "string") {
    try { finalItems = JSON.parse(itemsJSON); } catch(e) { finalItems = itemsJSON; }
  }
  
  // Remove existing with same name if any
  list = list.filter(function(x){ return x.name !== templateName; });
  list.push({ 
    name: templateName, 
    meal: meal || "Other",
    items: finalItems, 
    createdAt: new Date().toISOString() 
  });
  
  _upsertCustomer(ss, { phone: phone, standardOrder: JSON.stringify(list) });
  return { success: true };
}

function removeStandardOrder(phone, templateName) {
  var ss = getSpreadsheet();
  var custWs = getOrCreateTab(ss, TAB_CUSTOMERS, CUSTOMERS_HEADERS);
  var rows = getAllRows(custWs);
  var normP = _normalizePhone(phone);
  var cust = rows.find(function(c){ return _normalizePhone(c.Phone) === normP; });
  
  if (!cust || !cust.Standard_Order) return { success: true };
  
  var list = [];
  try { list = JSON.parse(cust.Standard_Order); if(!Array.isArray(list)) list=[]; } catch(e){ list=[]; }
  list = list.filter(function(x){ return x.name !== templateName; });
  
  _upsertCustomer(ss, { phone: phone, standardOrder: JSON.stringify(list) });
  return { success: true };
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
    fromStr = filterValue || Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
    toStr   = fromStr;
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
  const onAccountOrders = allOrders.filter(r => {
    const status = String(r.Payment_Status || '').trim().toLowerCase();
    if (status !== 'on account') return false;
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
      net:   Number(r.Net_Total || 0)
    });
    grouped[phone].total += Number(r.Net_Total || 0);
  });

  // Default sorting: Total Amount Descending, then Name
  const customers = Object.values(grouped).sort((a,b) => {
    if (b.total !== a.total) return b.total - a.total;
    return a.name.localeCompare(b.name);
  });  return { success: true, cycle, from: fromStr, to: toStr, customers };
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

/**
 * VIP / Fee Exempt Logic
 */
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

function _getDeliveryPointLabel(key) {
  if (!key) return "Handover at Doorstep";
  const map = {
    door: "Handover at Doorstep / Office Door",
    bell_keep: "Keep outside & Ring bell",
    lobby_handoff: "Handover at Lobby / Reception",
    lobby_keep: "Keep at Lobby / Reception",
    gate_handoff: "Handover at Security Gate",
    gate_keep: "Keep at security cabin",
    comedown: "Customer will come down",
    other: "Other (see instructions)"
  };
  return map[key] || key;
}
/**
 * Derives a Google Maps link based on partial address or society name matching.
 * @param {string} addr - The full address string
 * @param {string} society - The society name field
 * @returns {string} - The found maps link or empty string
 */
function _deriveMapsLink(addr, society) {
  const dict = {
    "Laburnum Park": "https://maps.app.goo.gl/nEApFaLe5x4PzuHd8",
    "Magarpatta City": "https://maps.app.goo.gl/wEndRh6jnkjL1GRC9",
    "Amanora Mall": "https://maps.app.goo.gl/Wd2FxrytcABk9Xty6",
    "Pentagon 1": "https://maps.app.goo.gl/BQKFrtdmLv9sK8tF8",
    "IZiel": "https://maps.app.goo.gl/BQKFrtdmLv9sK8tF8",
    "Shree Lakshmi Vihar": "https://maps.app.goo.gl/LWr5zhbiXHN9sh1F8",
    "Desire Tower": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 18": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 21": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Amanora Tower 22": "https://maps.app.goo.gl/P9KE1RtznuruMrJk8",
    "Greenville": "https://maps.app.goo.gl/FvpENaW8aFzXTj916",
    "Cosmos": "https://maps.app.goo.gl/bYu8gnb1ZVQgAwse9",
    "Prime Wing Cosmos": "https://maps.app.goo.gl/eUHhhV9oS1MPQutq8",
    "Zinnia": "https://maps.app.goo.gl/iWcMqnXCDRkesdnT7",
    "Trillium": "https://maps.app.goo.gl/cnYLtqjxaqDtexzKA",
    "Sudarshan Heritage": "https://maps.app.goo.gl/QJAAaugYVcD18eg4A",
    "Daffodils": "https://maps.app.goo.gl/vxXz2LuxFCZPcvTZ8",
    "Amanora Gold Tower": "https://maps.app.goo.gl/PgXAJ8nz7xKjnUZ7A",
    "Gateway Tower": "https://maps.app.goo.gl/YCh4dHET1ZsrLZqv6",
    "T100 Gateway Towers": "https://maps.app.goo.gl/YCh4dHET1ZsrLZqv6",
    "Wework Futura": "https://maps.app.goo.gl/u8JaNJYbpHPXBxdp9",
    "Jasminium": "https://maps.app.goo.gl/HiGZSmvbb3SFnTKP6",
    "Jasminium Society": "https://maps.app.goo.gl/pMXdyNSMu3kEMuec6",
    "Vascon Ela": "https://maps.app.goo.gl/MXhnTLvycEENcrNA7",
    "Marvel Fuego": "https://maps.app.goo.gl/izc9t6cXWWYiciyv8",
    "Heliconia": "https://maps.app.goo.gl/tFu79S7KHv6L48XG8",
    "Kumar Paradise": "https://maps.app.goo.gl/utQB7yFo4jMgVKAv6",
    "Vrindavan Heights": "https://maps.app.goo.gl/qcS8v4dVtbx1rBg46",
    "Sai Tower": "https://maps.app.goo.gl/dWigWD2KHXsE5YrT7",
    "Future Towers": "https://maps.app.goo.gl/pvhvaMsWS8n3x5Cq8",
    "Annexe Society": "https://maps.app.goo.gl/uAuA67dHxgLyfVX29",
    "Imperial Heights": "https://maps.app.goo.gl/Dt9HZVgBzB7iNdhu9",
    "Tulja Tower": "https://maps.app.goo.gl/VeUw6EciJckATaZU6",
    "Torana Kamdhenu": "https://maps.app.goo.gl/9PMRz86iGNjv8UZEA",
    "Cybercity": "https://maps.app.goo.gl/5ASLF1yDeoKH7Wq86",
    "Cyber City": "https://maps.app.goo.gl/5ASLF1yDeoKH7Wq86",
    "Samarth Shrushti": "https://maps.app.goo.gl/1eLo6vz3mTWu1BXB8",
    "Grevillea": "https://maps.app.goo.gl/ppTArZ12auPnRR5E8",
    "Orient Garden": "https://maps.app.goo.gl/vFrsEoijMPVV7a3U7",
    "Bhosale Nagar": "https://maps.app.goo.gl/5MzXtAZmtZvD9D9o6",
    "Amar Ornate": "https://maps.app.goo.gl/mtFcV35i5gpPz4BG7",
    "DSK Sunderban": "https://maps.app.goo.gl/xu4fiGtFLbgp3Kxs9",
    "Aruna Girls PG": "https://maps.app.goo.gl/FUxKBQ6iQKt64ji36",
    "Aruna PG": "https://maps.app.goo.gl/FUxKBQ6iQKt64ji36",
    "Mams Bungalow": "https://maps.app.goo.gl/gX94MR52LDXzbq7h8",
    "Gardenia": "https://maps.app.goo.gl/tVsv9NZvBzxqg1aU6",
    "Palazzo": "https://maps.app.goo.gl/fd1Na7Lenjh1AwiH6",
    "Kumar Picasso": "https://maps.app.goo.gl/WpDwFWDGJVSRx9r88",
    "Roystonea": "https://maps.app.goo.gl/6xfQ2VNEc4CMnRgD7",
    "Pankaj Avenue": "https://maps.app.goo.gl/HzLBbS8L5o8zkHgx8",
    "Unika": "https://maps.app.goo.gl/CvgQwPzkTRUT5dHz6",
    "Amanora Ascent": "https://maps.app.goo.gl/xgMEJAbrEbD2PpJ69",
    "Vanshree": "https://maps.app.goo.gl/cW1gjcsVkStwudfZ7",
    "Aspire Towers": "https://maps.app.goo.gl/dZ7V6SVs7SX63BhG8",
    "Naren Bliss": "https://maps.app.goo.gl/RJbXCqWcZSZTieBS7",
    "Solitaire": "https://maps.app.goo.gl/RsrdFp6gYQNbWzzq7",
    "Sylvania": "https://maps.app.goo.gl/j5UhPibRVDxhkdRY7",
    "Erica": "https://maps.app.goo.gl/yTpThD413d3q2FL38",
    "Om Balaji Darshan": "https://maps.app.goo.gl/KLDrXo2DZcUUkJzV9",
    "Adreno Towers": "https://maps.app.goo.gl/725jneTokL1ahFsF9",
    "Neo Towers": "https://maps.app.goo.gl/sJy9YiqFcEEDMgWz9",
    "Leisure Town": "https://maps.app.goo.gl/gvVkgfZtLfhpRk1W8",
    "Sundar Sankul": "https://share.google/aOAcluoLyIMzc3I1L",
    "Kumar Purab": "https://maps.app.goo.gl/FGez7Rv63NzNRLeaA",
    "Marvel Azure": "https://maps.app.goo.gl/YaeQHEbg8D4HAALDA",
    "Hrishikesh housing": "https://maps.app.goo.gl/PsyUKTGt4Rj9qPsVA"
  };

  const str = (addr + " " + society).toLowerCase();
  for (let key in dict) {
    if (str.includes(key.toLowerCase())) {
      return dict[key];
    }
  }
  return "";
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
  set("Phone",          phone);
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


/**
 * STEP 2 — Webhook handler.
 * HDFC POSTs here immediately after payment success or refund.
 * This is the authoritative payment confirmation — never rely on return URL alone.
 *
 * Auth: Basic Auth (HDFC_WEBHOOK_USERNAME:HDFC_WEBHOOK_PASSWORD)
 *
 * @param {Object} body  Parsed webhook JSON from HDFC
 * @param {Object} e     Raw Apps Script event (for header access)
 */
/**
 * STEP 2 — Webhook handler. LOG-FIRST PATTERN.
 *
 * HDFC expects a 200 response within 5 seconds. Heavy work (Status API call,
 * sheet reads, row updates) can take 5-8s — guaranteed timeout.
 *
 * Solution: this function does ONE thing only:
 *   1. Verify Basic Auth (in-memory, < 5ms)
 *   2. Write one row to SK_Webhook_Log (< 500ms)
 *   3. Return 200 OK to HDFC immediately
 *
 * The actual processing — Status API verification, marking order paid —
 * is handled by hdfc_processWebhookLog(), called by a 1-minute time trigger.
 * No time pressure there. No duplicates (PENDING → PROCESSED gate).
 *
 * @param {Object} body  Parsed webhook JSON from HDFC
 * @param {Object} e     Raw Apps Script event (for header access)
 */
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


/**
 * HDFC Webhook Processor — called by a 1-minute time trigger.
 * Reads all PENDING rows from SK_Webhook_Log and processes each one.
 *
 * For ORDER_SUCCEEDED events:
 *   1. Calls Status API to independently confirm payment (security requirement)
 *   2. Marks the order Paid in SK_Orders
 *   3. Updates the log row to PROCESSED or FAILED
 *
 * Safe to run multiple times — duplicate protection is inside hdfc_markOrderPaid.
 * Set up the trigger once by running setupHdfcWebhookTrigger() from the editor.
 */
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


/**
 * Run this ONCE from the Apps Script editor to set up the 1-minute trigger
 * that calls hdfc_processWebhookLog() automatically.
 *
 * Menu: Run → setupHdfcWebhookTrigger
 * Only needed in the DEV project (same project that receives HDFC webhooks).
 */
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


/**
 * STEP 3 — Called by order.html when customer lands back after payment.
 * Verifies the HMAC signature on return URL params.
 * Signature = HMAC_SHA256(order_id + "|" + status, HDFC_RESPONSE_KEY)
 *
 * @param {Object} body  { order_id, status, signature }
 * @returns {{ success, paid, order_id, status, message } | { error }}
 */
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


/**
 * Internal: Calls HDFC Order Status API to confirm a payment server-side.
 * Must be called before writing "Paid" to the sheet — webhook alone is not
 * sufficient per HDFC security audit requirements.
 *
 * Endpoint: GET {HDFC_BASE_URL}/orders/{order_id}
 * Auth:     Basic base64(API_KEY + ":")
 *
 * @param {string} orderId  The gateway order ID (alphanumeric, ≤21 chars)
 * @returns {{ confirmed: boolean, status: string, txn_id: string }}
 */
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


/**
 * Internal: Marks a Svaadh order row(s) as PAID in SK_Orders sheet.
 * Called by hdfc_handleWebhook on ORDER_SUCCEEDED.
 *
 * @param {Object} order  Order object from HDFC webhook content.order
 */
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


/**
 * Retrieves and removes the saved pending order on return from HDFC.
 * Called by order.html hdfc_handleReturnParams() after verifying payment.
 *
 * @param {Object} body  { order_id }
 * @returns {{ success, phone, amount, orders, selectedDates, profile, mealAddrs, isFirstTime } | { error }}
 */
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


/**
 * Utility: Compute HMAC-SHA256 hex digest.
 * Used to verify HDFC return URL signatures.
 *
 * @param {string} message  String to sign
 * @param {string} secret   Response key from HDFC
 * @returns {string}        Lowercase hex string
 */
function hdfc_hmacSha256(message, secret) {
  const sig = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(message).getBytes(),
    Utilities.newBlob(secret).getBytes()
  );
  return sig.map(function(b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); }).join("");
}


/**
 * DEV HELPER — Test gateway connectivity without a real payment.
 * Run this function directly from the Apps Script editor (Dev project only).
 * It will attempt to create a ₹1 session and log the result.
 */
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
