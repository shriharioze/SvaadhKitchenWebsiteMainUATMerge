"""Build 10_Hdfc_Gateway.gs entirely from UAT's HDFC functions."""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

sys.path.insert(0, r"C:\Users\admin\Downloads\LLM Course\Projects\SvaadhKitchenMerged")
from _extract_uat_fns import extract_function

UAT_SRC = r"C:\Users\admin\Downloads\SvaadhUAT\DevCode.gs"
OUT = r"C:\Users\admin\Downloads\LLM Course\Projects\SvaadhKitchenMerged\10_Hdfc_Gateway.gs"

# Order matters for readability — group by purpose.
HDFC_FUNCTIONS_IN_ORDER = [
    # Helpers used by everything else
    "hdfc_hmacSha256",
    "_hdfcReturnRedirectHtml",
    "_checkWebhookLogForCharge",
    "_computeAuthoritativeTotal",

    # Core session flow
    "hdfc_createSession",
    "hdfc_getOrderStatus",
    "hdfc_verifyReturnPayload",

    # Pending-order state
    "hdfc_savePendingOrder",
    "hdfc_getPendingOrder",

    # Webhook handling
    "hdfc_handleWebhook",
    "hdfc_processWebhookLog",
    "hdfc_markOrderPaid",

    # Setup and test
    "setupHdfcWebhookTrigger",
    "testHdfcConnection",
]


def main():
    with open(UAT_SRC, "r", encoding="utf-8") as f:
        src = f.read()

    header = (
        "// ============================================================\n"
        "// 10_Hdfc_Gateway.gs\n"
        "// HDFC SmartGateway integration — session creation, webhook\n"
        "// handling, Status API, HMAC, post-payment verification,\n"
        "// authoritative server-side recompute (Burp tamper protection).\n"
        "// ============================================================\n"
        "// All functions here are gated by PAYMENT_GATEWAY_ENABLED.\n"
        "// Sourced from SvaadhKitchenUAT v14.8 — includes:\n"
        "//   - Bug-fix: verification no longer trusts client-sent status\n"
        "//     (requires Status API confirmed=true OR HMAC-verified webhook).\n"
        "//   - Server-authoritative pricing recompute that ignores any\n"
        "//     client-side amount tampering.\n"
        "//   - Webhook-log fallback when Status API is quota-exhausted.\n"
        "// ============================================================\n\n"
    )

    out_parts = [header]
    for fn in HDFC_FUNCTIONS_IN_ORDER:
        text = extract_function(src, fn)
        if text is None:
            print(f"  ! NOT FOUND in UAT: {fn}", file=sys.__stderr__)
            continue
        out_parts.append(text)
        if not text.endswith("\n"):
            out_parts.append("\n")
        out_parts.append("\n")
        print(f"  + {fn}  ({text.count(chr(10))} lines)", file=sys.__stderr__)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("".join(out_parts))

    total_lines = "".join(out_parts).count("\n")
    print(f"\nWrote {OUT}  ({total_lines} lines)", file=sys.__stderr__)


if __name__ == "__main__":
    main()
