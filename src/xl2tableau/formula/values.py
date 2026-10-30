"""Excel value model and coercion rules.

Scalar values are: float, str, bool, None (blank cell) or XlError.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import List, Optional, Sequence, Tuple


@dataclass(frozen=True)
class XlError:
    code: str

    def __str__(self) -> str:
        return self.code


DIV0 = XlError("#DIV/0!")
NA = XlError("#N/A")
VALUE = XlError("#VALUE!")
REF = XlError("#REF!")
NAME = XlError("#NAME?")
NUM = XlError("#NUM!")
CIRC = XlError("#CIRC!")  # not a real Excel code; marks cells in a reference cycle

_NUMERIC_TEXT = re.compile(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")


@dataclass
class RangeVal:
    """A materialised rectangular range. ``key`` identifies it for caching."""

    key: Tuple
    rows: List[List[object]]
    first_row: int
    first_col: int

    def flat(self):
        for r in self.rows:
            yield from r

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.rows[0]) if self.rows else 0


def parse_numeric_text(s: str) -> Optional[float]:
    if _NUMERIC_TEXT.fullmatch(s):
        return float(s)
    return None


def to_number(v) -> object:
    """Coerce a scalar to float the way Excel arithmetic does."""
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return 0.0
    if isinstance(v, XlError):
        return v
    if isinstance(v, str):
        n = parse_numeric_text(v.strip())
        return VALUE if n is None else n
    return VALUE


def number_to_text(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return f"{x:.15g}"


def to_text(v) -> object:
    if isinstance(v, XlError):
        return v
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return number_to_text(float(v))
    return str(v)


def to_bool(v) -> object:
    if isinstance(v, XlError):
        return v
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        u = v.upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
    return VALUE


def excel_round(x: float, digits: int) -> float:
    """ROUND with half-away-from-zero, like Excel (Python's round is half-even)."""
    if math.isnan(x) or math.isinf(x):
        return x
    q = Decimal(1).scaleb(-digits)
    d = Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP)
    return float(d)


def _type_rank(v) -> int:
    # Excel orders mixed types as: numbers < text < logicals.
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def compare(op: str, a, b) -> object:
    if isinstance(a, XlError):
        return a
    if isinstance(b, XlError):
        return b
    # A blank takes the type of the other side.
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
    ra, rb = _type_rank(a), _type_rank(b)
    if ra != rb:
        c = (ra > rb) - (ra < rb)
    elif ra == 1:
        la, lb = a.lower(), b.lower()
        c = (la > lb) - (la < lb)
    else:
        fa, fb = float(a), float(b)
        c = (fa > fb) - (fa < fb)
    return {
        "=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0,
    }[op]


def lookup_key(v):
    """Normalised key used for exact-match lookups (case-insensitive text)."""
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    if isinstance(v, str):
        return ("s", v.lower())
    if v is None:
        return ("blank",)
    return ("e", str(v))


# ---------------------------------------------------------------- criteria

@dataclass(frozen=True)
class Criterion:
    """A parsed SUMIF/COUNTIF criterion such as 5, "East", ">=10" or "A*"."""

    op: str
    operand: object  # float, str or bool
    wildcard: Optional[re.Pattern]

    @property
    def index_key(self):
        """Hashable key when this is a plain equality test, else None."""
        if self.op != "=" or self.wildcard is not None:
            return None
        if isinstance(self.operand, str) and self.operand == "":
            return None
        return criterion_cell_key(self.operand)


def criterion_cell_key(v):
    """Key under which a cell value is stored in a criteria index."""
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", float(v))
    if isinstance(v, str):
        n = parse_numeric_text(v)
        return ("n", n) if n is not None else ("s", v.lower())
    if v is None:
        return ("blank",)
    return None  # errors never match


def _wildcard_regex(s: str) -> Optional[re.Pattern]:
    if "*" not in s and "?" not in s:
        return None
    out, i = [], 0
    while i < len(s):
        ch = s[i]
        if ch == "~" and i + 1 < len(s) and s[i + 1] in "*?~":
            out.append(re.escape(s[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


def parse_criterion(c) -> Criterion:
    if isinstance(c, bool) or isinstance(c, (int, float)):
        return Criterion("=", float(c) if not isinstance(c, bool) else c, None)
    if c is None:
        return Criterion("=", 0.0, None)
    s = str(c)
    op = "="
    for cand in ("<=", ">=", "<>", "<", ">", "="):
        if s.startswith(cand):
            op, s = cand, s[len(cand):]
            break
    n = parse_numeric_text(s)
    if n is not None:
        return Criterion(op, n, None)
    if s.upper() in ("TRUE", "FALSE"):
        return Criterion(op, s.upper() == "TRUE", None)
    wc = _wildcard_regex(s) if op in ("=", "<>") else None
    return Criterion(op, s, wc)


def criterion_matches(crit: Criterion, v) -> bool:
    if isinstance(v, XlError):
        return False
    op, x = crit.op, crit.operand
    if isinstance(x, str) and x == "" and op in ("=", "<>"):
        is_blank = v is None or v == ""
        return is_blank if op == "=" else not is_blank
    if crit.wildcard is not None:
        hit = isinstance(v, str) and crit.wildcard.fullmatch(v) is not None
        return hit if op == "=" else not hit
    if isinstance(x, bool):
        if op == "=":
            return isinstance(v, bool) and v == x
        if op == "<>":
            return not (isinstance(v, bool) and v == x)
        return False
    if isinstance(x, float):
        if isinstance(v, bool):
            cv = None
        elif isinstance(v, (int, float)):
            cv = float(v)
        elif isinstance(v, str):
            cv = parse_numeric_text(v) if op in ("=", "<>") else None
        else:
            cv = None
        if cv is None:
            return op == "<>"
        return {"=": cv == x, "<>": cv != x, "<": cv < x, ">": cv > x,
                "<=": cv <= x, ">=": cv >= x}[op]
    # text operand, no wildcard
    if not isinstance(v, str) or parse_numeric_text(v) is not None:
        return op == "<>"
    a, b = v.lower(), x.lower()
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]


def flatten_args(values: Sequence) -> List:
    out = []
    for v in values:
        if isinstance(v, RangeVal):
            out.extend(v.flat())
        else:
            out.append(v)
    return out
