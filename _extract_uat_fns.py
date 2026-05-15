"""
Extract specific top-level functions from UAT's DevCode.gs.
Used to pull HDFC improvements over to the merged repo.

Line-range-based extraction: finds the target function's declaration line,
then takes everything up to the next top-level `function ` declaration.
This avoids the fragility of brace-counting through JS string literals
that contain `{` and `}` (e.g. CSS in template literals).

Usage:
  python _extract_uat_fns.py FUNC1 FUNC2 ...
"""
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SRC = r"C:\Users\admin\Downloads\SvaadhUAT\DevCode.gs"

_DECL_RE = re.compile(r"^function\s+([A-Za-z_$][\w$]*)\s*\(")


def _index_functions(lines):
    """Return list of (name, line_idx) for all top-level function decls."""
    out = []
    for i, line in enumerate(lines):
        m = _DECL_RE.match(line)
        if m:
            out.append((m.group(1), i))
    return out


def extract_function(src, name):
    """Return the full text of function `name`, including an immediately
    preceding contiguous comment block (no blank-line gap)."""
    lines = src.split("\n")
    index = _index_functions(lines)

    decl_idx = None
    next_idx = len(lines)
    for pos, (fname, i) in enumerate(index):
        if fname == name:
            decl_idx = i
            if pos + 1 < len(index):
                next_idx = index[pos + 1][1]
            break

    if decl_idx is None:
        return None

    # Walk back through immediately-preceding contiguous comment lines.
    start = decl_idx
    j = decl_idx - 1
    while j >= 0:
        stripped = lines[j].strip()
        if stripped == "":
            break
        if not (stripped.startswith("//")
                or stripped.startswith("/*")
                or stripped.startswith("*")
                or stripped.endswith("*/")):
            break
        start = j
        j -= 1

    # Take everything up to (but not including) the next top-level function.
    end = next_idx
    # Trim trailing blank lines.
    while end > decl_idx + 1 and lines[end - 1].strip() == "":
        end -= 1
    # Trim a trailing contiguous comment block — it belongs to the NEXT
    # function, not ours.
    while end > decl_idx + 1:
        stripped = lines[end - 1].strip()
        if (stripped.startswith("//")
                or stripped.startswith("/*")
                or stripped.startswith("*")
                or stripped.endswith("*/")):
            end -= 1
            continue
        break
    # Then trim any blank lines exposed by that.
    while end > decl_idx + 1 and lines[end - 1].strip() == "":
        end -= 1

    return "\n".join(lines[start:end]) + "\n"


def main():
    if len(sys.argv) < 2:
        print("usage: _extract_uat_fns.py FUNC1 FUNC2 ...")
        sys.exit(1)
    with open(SRC, "r", encoding="utf-8") as f:
        src = f.read()
    for name in sys.argv[1:]:
        text = extract_function(src, name)
        if text is None:
            print(f"## NOT FOUND: {name}", file=sys.stderr)
            continue
        print(f"// === Extracted from UAT DevCode.gs: {name} ===")
        print(text)
        print()


if __name__ == "__main__":
    main()
