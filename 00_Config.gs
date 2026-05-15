// ============================================================
// 00_Config.gs
// Constants, Script Property references, sheet column layouts, item maps,
  meal item config (FIXED_MEAL_ITEMS), and chatbot knowledge base.
  All environment-specific values come from Script Properties — never
  hardcoded.
// ============================================================
// This file is part of a modular split of the original Code.gs.
// Apps Script merges all .gs files into one global scope at
// load time, so cross-file function calls work without imports.
// ============================================================

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
const GA4_PROPERTY_ID       = "396771381"; // User provided Property ID
const CODE_VERSION   = 14.9; // Soft-cancel: orders are never deleted, marked with status remark for audit trail
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
const TAB_GA4_METRICS = "SK_Analytics_Data"; // Google Analytics Storage
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
  "Inflation_Surcharge", "Loyalty_Discount", "Wallet_Credit",
  // HDFC SmartGateway columns (auto-healed by getOrCreateTab).
  // Gateway_Order_ID provides idempotent dedup at submitOrder time so
  // a single HDFC charge can never produce two SK_Orders rows.
  // Meal_Credit tracks how much retroactive day-free delivery credit
  // has already been refunded for this customer's day total — prevents
  // the double-refund bug when multiple meals are placed on the same date.
  "Gateway_Order_ID", "Meal_Credit"
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
// ── GOOGLE ANALYTICS 4 INTEGRATION ──────────────────────────
const GA4_HEADERS = ["Date", "Source", "Device", "Active_Users", "Sessions", "Page_Views", "Engagement_Rate", "Avg_Session_Duration", "Event_Count"];
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
// ── INVENTORY ─────────────────────────────────────────────────────────────────
// Tracks raw material purchases. Each new entry for the same item auto-calculates
// how long the previous batch lasted → builds consumption rate over time.
const TAB_INVENTORY      = "SK_Inventory";
const INVENTORY_HEADERS  = [
  "Entry_ID","Date","Item","Unit","Quantity","Price_Paid","Notes","Timestamp"
];
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
// ── CLIENT ERROR LOG ──────────────────────────────────────────────────────────
const TAB_ERROR_LOG     = "SK_Error_Log";
// Column layout: structured JSON fields extracted for easy Sheets filtering.
// "Extra_JSON" holds any additional fields the client sends beyond the core set.
const ERROR_LOG_HEADERS = [
  "Timestamp","Date","Phone","Version","Type","Action",
  "Attempt","Duration_ms","Message","URL","Extra_JSON"
];
