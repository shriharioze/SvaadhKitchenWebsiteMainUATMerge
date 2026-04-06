# Svaadh Kitchen — New System Setup Guide

## What's in this package

| File | Purpose |
|------|---------|
| `Code.gs` | Google Apps Script — backend for order.html + admin.html |
| `auto_svaadh_summary.ipynb` | Master Colab orchestrator |
| `customer_info_manager.py` | Builds customer master from first-time orders |
| `kitchen_live_manager.py` | Pushes live totals to Kitchen Dashboard sheet |
| `inventory_manager.py` | Calculates packaging costs per day |
| `dispatch_manager.py` | Generates driver dispatch Excel |
| `label_generator.py` | Generates PDF labels from Sheet directly |
| `label_generator.py` | Generates PDF labels from Sheet directly |

---

## Step 1 — Create the new Google Sheet

1. Go to sheets.google.com → New blank spreadsheet
2. Name it: **Svaadh Kitchen Orders (New)**
3. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit`

---

## Step 2 — Deploy Code.gs

1. Open your new Google Sheet → Extensions → Apps Script
2. Delete all existing code, paste in `Code.gs`
3. Find the top of the file, replace:
   ```
   const SHEET_ID = "17X7JOrMe1Oj_QykH7mk6UGoBjuguLfyC1RmKYATDXlI";
   ```
   with your actual Sheet ID from Step 1

4. Save (Ctrl+S)

5. Run `initSchema()` once to create all tabs with correct headers:
   - Click the function dropdown → select `initSchema` → Run
   - You'll see 5 tabs created: SK_Orders, SK_Customers, SK_Daily_Menu, SK_Master_Breakfast, SK_Master_Sabjis

6. Deploy as Web App:
   - Click Deploy → New Deployment
   - Type: Web App
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy → copy the web app URL

---

## Step 3 — Update order.html and admin.html

In both files, find:
```js
const APPS_SCRIPT_URL = "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
```
Replace with the URL from Step 2.

> ⚠️ Every time you change Code.gs and redeploy, you must create a **New Version**
> in the deployment settings — the URL stays the same but the new code won't run
> unless you do this.

---

## Step 4 — Google Drive folder for Colab

Create this folder structure in your Drive:
```
My Drive/
└── Svaadh Kitchen/
    └── New System/              ← all .py files go here
        ├── service_account.json
        ├── auto_svaadh_summary.ipynb
        ├── customer_info_manager.py
        ├── kitchen_live_manager.py
        ├── inventory_manager.py
        ├── dispatch_manager.py
        ├── label_generator.py
        └── label_generator.py
```

---

## Step 5 — Update Sheet ID in Python files

In each `.py` file, find:
```python
SHEET_ID = "17X7JOrMe1Oj_QykH7mk6UGoBjuguLfyC1RmKYATDXlI"
```
Replace with your actual Sheet ID from Step 1.

Files to update:
- `customer_info_manager.py`
- `label_generator.py`
- `auto_svaadh_summary.ipynb` (Cell 2)

> `kitchen_live_manager.py`, `dispatch_manager.py`, and `inventory_manager.py`
> receive data passed in from the orchestrator — no Sheet ID needed in those files.

---

## Step 6 — First run checklist

1. Open `auto_svaadh_summary.ipynb` in Colab
2. Run Cell 1 (mount Drive)
3. Run Cell 2 (auth) — confirm "✅ Auth OK"
4. Run Cell 3 (load modules) — confirm "✅ All modules loaded"
5. Run Cell 4 (choose date)
6. Run Cell 5 (define functions)
7. Run Cell 6 (start live loop)

---

## Sheet tab reference

### SK_Orders (main tally)
One row per meal per customer per date.

| Column | What it is |
|--------|-----------|
| Submission_ID | `SK-YYYYMMDD-XXXX` |
| Order_Date | `YYYY-MM-DD` |
| Meal_Type | Breakfast / Lunch / Dinner |
| Customer_Name | Customer's name |
| Phone | 10-digit mobile |
| Area | Delivery area |
| Wing/Flat/Floor/Society | Address parts |
| Full_Address | Combined address string |
| Maps_Link | Google Maps URL (optional) |
| Landmark | Directions for driver (optional) |
| Items_JSON | `{"Chapati":3,"Dal":1}` |
| Chapati → Curd | Individual item quantities |
| BF_Item_1/BF_Qty_1 … | Breakfast items (up to 4) |
| Special_Notes | e.g. "No onion" |
| Food_Subtotal | Before delivery/discount |
| Delivery_Charge | 0 or ₹10 |
| Discount_Amount | 5% or 7.5% applied |
| Net_Total | Final amount |
| Payment_Status | **Pending** (update manually to Paid or Wallet Paid) |
| First_Time | Yes / No |
| Source | WebApp |

### SK_Customers
One row per phone number. Updated on every order.

### SK_Daily_Menu
One row per date. Set via admin.html.

### SK_Master_Breakfast / SK_Master_Sabjis
Master item lists. Managed via admin.html.

---

## Parallel operation with old Tally system

The old Tally form → Sheet `1qWIMHMKMbx0TqPto80iY8n4PVUYWZ8s0Q3KIYSpvgmM` keeps running.
The new system writes to a completely separate Sheet.
The old Python orchestrator keeps running from its old folder.
The new orchestrator runs from `New System/` folder.

When you're ready to cut over fully:
1. Stop the old orchestrator (interrupt kernel)
2. Switch the `order.html` link on your website to the new version
3. Archive the old Tally form

---

## Changing the admin PIN

In `Code.gs`, line 5:
```js
const ADMIN_PIN = "1234";
```
Change to any 4-digit code. Then redeploy (new version).

---

## Generating labels manually

Either run from the Colab notebook (last cell) or:
```bash
cd "Svaadh Kitchen/New System"
python label_generator.py
```
It will prompt for date, meal, and language.

