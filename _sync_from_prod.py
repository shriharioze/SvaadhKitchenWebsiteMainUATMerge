"""
Sync non-payment .gs function bodies in this merged repo from the
latest prod Code.gs. Protected functions (payment / our merge-specific
work) are kept as-is. Everything else is overwritten with prod's body.

Run from the repo root:  python _sync_from_prod.py
"""
import io
import re
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PROD = r"C:\Users\admin\Downloads\LLM Course\Projects\ChatBot - final\Code.gs"
MODULES = [
    "00_Config.gs",
    "01_Router.gs",
    "02_Utils.gs",
    "03_Customer.gs",
    "04_Menu_Admin.gs",
    "05_Orders.gs",
    "06_Wallet_Billing.gs",
    "07_Kitchen_Delivery.gs",
    "08_Analytics_Inventory.gs",
    "09_Archive.gs",
    "12_Admin_Tools.gs",
    "13_Maintenance.gs",
]

# Functions where the MERGED repo's version intentionally diverges from
# prod and MUST NOT be overwritten:
#   - All HDFC gateway / reconciler / popup-related code lives in 10/11.
#   - submitOrder + _submitOrderInternal carry the Gateway_Order_ID
#     idempotency guard and the Meal_Credit double-refund fix.
#   - getDayTotalsForDates surfaces meal_credit_applied.
#   - get{Unpaid,Order,Customer}History + getBillingData carry our
#     archive-aware lookup additions.
#   - reconcileTransactions has the tier-2 multi-row sum logic.
#   - doGet/doPost contain HDFC routing.
PROTECTED = {
    "doGet", "doPost",
    "hdfc_hmacSha256", "_hdfcReturnRedirectHtml", "_checkWebhookLogForCharge",
    "_computeAuthoritativeTotal", "hdfc_createSession", "hdfc_getOrderStatus",
    "hdfc_verifyReturnPayload", "hdfc_savePendingOrder", "hdfc_getPendingOrder",
    "hdfc_handleWebhook", "hdfc_processWebhookLog", "hdfc_markOrderPaid",
    "setupHdfcWebhookTrigger", "testHdfcConnection",
    "hdfc_createWalletRechargeSession", "hdfc_finalizeWalletRecharge",
    "setupReconcileTrigger", "reconcilePendingOrders",
    "_reconcileSingleEntry", "_buildSubmitBodyFromPending",
    "submitOrder", "_submitOrderInternal",
    "getDayTotalsForDates",
    "getUnpaidCustomers", "getBillingData", "getOrderHistory", "getCustomerHistory",
    "reconcileTransactions",
    "voidOrderRow",
}

# Functions where merged needs prod's version EVEN THOUGH the name is
# also somewhere protected nearby. Currently empty.
FORCE_OVERWRITE = set()

_DECL = re.compile(r"^function\s+([A-Za-z_$][\w$]*)\s*\(")


def parse_functions(path):
    with open(path, "r", encoding="utf-8") as f:
        lines = f.read().split("\n")
    idx = []
    for i, ln in enumerate(lines):
        m = _DECL.match(ln)
        if m:
            idx.append((m.group(1), i))
    fns = {}
    for k, (name, start) in enumerate(idx):
        end = idx[k + 1][1] if k + 1 < len(idx) else len(lines)
        # Walk back through immediately-preceding contiguous comment lines
        # so the doc-block moves with the function body.
        s = start
        j = start - 1
        while j >= 0:
            stripped = lines[j].strip()
            if stripped == "":
                break
            if not (
                stripped.startswith("//")
                or stripped.startswith("/*")
                or stripped.startswith("*")
                or stripped.endswith("*/")
            ):
                break
            s = j
            j -= 1
        # Trim trailing blank/comment block that belongs to next function
        e = end
        while e > start + 1 and lines[e - 1].strip() == "":
            e -= 1
        while e > start + 1:
            t = lines[e - 1].strip()
            if (
                t.startswith("//")
                or t.startswith("/*")
                or t.startswith("*")
                or t.endswith("*/")
            ):
                e -= 1
                continue
            break
        while e > start + 1 and lines[e - 1].strip() == "":
            e -= 1
        fns[name] = "\n".join(lines[s:e])
    return fns


def normalise(b):
    """Loose equality — ignore leading/trailing whitespace + final
    newline so trivial trailing-whitespace diffs don't trigger a rewrite."""
    return b.strip()


def main():
    prod = parse_functions(PROD)
    rewrites = 0
    touched_modules = []

    for mod in MODULES:
        with open(mod, "r", encoding="utf-8") as f:
            src = f.read()
        lines = src.split("\n")
        idx = []
        for i, ln in enumerate(lines):
            m = _DECL.match(ln)
            if m:
                idx.append((m.group(1), i))
        # Build the (start_line, end_line, name) ranges, INCLUDING the
        # preceding contiguous comment block per function, mirroring
        # parse_functions above.
        ranges = []
        for k, (name, decl) in enumerate(idx):
            end = idx[k + 1][1] if k + 1 < len(idx) else len(lines)
            s = decl
            j = decl - 1
            while j >= 0:
                stripped = lines[j].strip()
                if stripped == "":
                    break
                if not (
                    stripped.startswith("//")
                    or stripped.startswith("/*")
                    or stripped.startswith("*")
                    or stripped.endswith("*/")
                ):
                    break
                s = j
                j -= 1
            e = end
            while e > decl + 1 and lines[e - 1].strip() == "":
                e -= 1
            while e > decl + 1:
                t = lines[e - 1].strip()
                if (
                    t.startswith("//")
                    or t.startswith("/*")
                    or t.startswith("*")
                    or t.endswith("*/")
                ):
                    e -= 1
                    continue
                break
            while e > decl + 1 and lines[e - 1].strip() == "":
                e -= 1
            ranges.append((name, s, e))

        # Walk ranges in REVERSE so earlier line-numbers stay valid.
        changed_in_mod = []
        for name, s, e in reversed(ranges):
            if name in PROTECTED and name not in FORCE_OVERWRITE:
                continue
            prod_body = prod.get(name)
            if prod_body is None:
                continue
            merged_body = "\n".join(lines[s:e])
            if normalise(merged_body) == normalise(prod_body):
                continue
            lines[s:e] = prod_body.split("\n")
            changed_in_mod.append(name)

        if changed_in_mod:
            new_src = "\n".join(lines)
            if not new_src.endswith("\n"):
                new_src += "\n"
            with open(mod, "w", encoding="utf-8") as f:
                f.write(new_src)
            rewrites += len(changed_in_mod)
            touched_modules.append((mod, changed_in_mod))
            print(f"  {mod}: replaced {len(changed_in_mod)} fn(s)")
            for n in changed_in_mod:
                print(f"     - {n}")

    print(f"\nTotal: {rewrites} function bodies overwritten from prod main.")


if __name__ == "__main__":
    main()
