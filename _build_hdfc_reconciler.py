"""Build 11_Hdfc_Reconciler.gs from UAT's reconciler functions."""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, r"C:\Users\admin\Downloads\LLM Course\Projects\SvaadhKitchenMerged")
from _extract_uat_fns import extract_function

UAT_SRC = r"C:\Users\admin\Downloads\SvaadhUAT\DevCode.gs"
OUT = r"C:\Users\admin\Downloads\LLM Course\Projects\SvaadhKitchenMerged\11_Hdfc_Reconciler.gs"

FUNCS = [
    "setupReconcileTrigger",
    "reconcilePendingOrders",
    "_reconcileSingleEntry",
    "_buildSubmitBodyFromPending",
]


def main():
    with open(UAT_SRC, "r", encoding="utf-8") as f:
        src = f.read()
    header = (
        "// ============================================================\n"
        "// 11_Hdfc_Reconciler.gs\n"
        "// Self-healing reconciliation for HDFC payments that charged\n"
        "// successfully at the gateway but never wrote a row in SK_Orders\n"
        "// (e.g. user closed the popup before the post-charge round-trip\n"
        "// completed). A 5-minute time-based trigger sweeps the pending\n"
        "// log, confirms each entry against the Status API / Webhook Log,\n"
        "// and writes the SK_Orders row using the cached cart state.\n"
        "// ============================================================\n"
        "// Gated by PAYMENT_GATEWAY_ENABLED — safe to deploy on live.\n"
        "// Sourced from SvaadhKitchenUAT v14.8.\n"
        "// ============================================================\n\n"
    )
    parts = [header]
    for fn in FUNCS:
        t = extract_function(src, fn)
        if t is None:
            print(f"  ! NOT FOUND in UAT: {fn}", file=sys.__stderr__)
            continue
        parts.append(t)
        if not t.endswith("\n"):
            parts.append("\n")
        parts.append("\n")
        print(f"  + {fn}  ({t.count(chr(10))} lines)", file=sys.__stderr__)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("".join(parts))
    total = "".join(parts).count("\n")
    print(f"\nWrote {OUT}  ({total} lines)", file=sys.__stderr__)


if __name__ == "__main__":
    main()
