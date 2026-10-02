"""Excel functions in the Python runtime (engine/xlruntime.py) where a model's labels or figures depend on them.

  TEXT      number formats, rounded half up as Excel does (1234.5 -> "1235", not Python's "1234"); date formats
            of day, month and year codes ("dd mmm yyyy"): a label like ="Valuation date: "&TEXT(F7,"dd mmm yyyy")
            reads as Excel shows it, not as the date's serial number

    uv run python tests/check_xlruntime.py
"""
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "engine"))
from xlruntime import serial, xl  # noqa: E402


def main() -> None:
    d = serial(date(2025, 6, 30))  # a Monday
    cases = [(d, "dd mmm yyyy", "30 Jun 2025"), (d, "d/mm/yy", "30/06/25"), (d, "mmmm yyyy", "June 2025"),
             (d, "yyyy-mm-dd", "2025-06-30"), (d, "dddd d mmmm", "Monday 30 June"), (d, "ddd", "Mon"),
             (d, '"FY"yy', "FY25"), (d, "mmm-yy", "Jun-25"),
             (0.0725, "0.00%", "7.25%"), (1234.5, "#,##0.0", "1,234.5"), (1234.5, "0", "1235"), (2.5, "0", "3"),
             (0.00125, "0.00%", "0.13%"), (-1234.5, "#,##0", "-1,235")]
    for v, fmt, want in cases:
        got = xl.TEXT(v, fmt)
        assert got == want, (v, fmt, got, want)
    print(f"TEXT: ok ({len(cases)} formats: dates by their day, month and year codes, numbers rounded half up "
          "as Excel does)")


def offset_check() -> None:
    """OFFSET with a negative height or width runs up or to the left from the cell it lands on, as Excel takes it (a
    terminal value on the average of the last N years: SUM(OFFSET(DO105,,,,-4*N))/N); nil height or width is #REF!."""
    from xlruntime import Rng
    at = Rng(None, "", "Val", 105, 119, 105, 119)  # DO105
    back = xl.OFFSET(at, 0, 0, 1, -8.0)  # the last 8 quarters, ending at DO
    assert (back.r1, back.c1, back.r2, back.c2) == (105, 112, 105, 119), (back.r1, back.c1, back.r2, back.c2)
    up = xl.OFFSET(at, -1, 2, -3, 2)
    assert (up.r1, up.c1, up.r2, up.c2) == (102, 121, 104, 122), (up.r1, up.c1, up.r2, up.c2)
    fwd = xl.OFFSET(at, 0, 1, 1, 4)
    assert (fwd.c1, fwd.c2) == (120, 123), (fwd.c1, fwd.c2)
    assert getattr(xl.OFFSET(at, 0, 0, 1, 0), "code", None) == "#REF!"
    assert getattr(xl.OFFSET(at, 0, 0, 1, -200), "code", None) == "#REF!"  # past the sheet's first column
    print("OFFSET: ok (a negative height or width runs up or to the left, as Excel; nil is #REF!)")


if __name__ == "__main__":
    main()
    offset_check()
