"""Counts Pyinex's own C++ lines under src/ and fails above the budget.

A line counts if it holds code: blank lines and lines that are only a comment do not, so
the budget never discourages explaining the code. Microsoft's SDK files are not under
src/ and do not count.

    python scripts/count_lines.py
"""

import re
import sys
from pathlib import Path

BUDGET = 2500
ROOT = Path(__file__).resolve().parent.parent


def code_lines(text: str) -> int:
    # Remove block comments, keeping their line breaks, then line comments.
    text = re.sub(r"/\*.*?\*/", lambda m: "\n" * m.group().count("\n"), text, flags=re.S)
    return sum(1 for line in text.splitlines() if line.strip() and not line.strip().startswith("//"))


def main() -> int:
    files = sorted(p for p in (ROOT / "src").rglob("*") if p.suffix in {".cpp", ".h", ".hpp"})
    total = 0
    for path in files:
        n = code_lines(path.read_text(encoding="utf-8"))
        total += n
        print(f"{n:6}  {path.relative_to(ROOT).as_posix()}")
    print(f"{total:6}  total (budget {BUDGET})")
    if total > BUDGET:
        print(f"over budget by {total - BUDGET} lines", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
