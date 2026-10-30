"""Excel worksheet function implementations.

Eager functions receive already-evaluated arguments (scalars or RangeVal).
Functions that need the evaluation context (for caches) take ``ctx`` first.
IF, IFERROR, AND and OR are handled lazily in the evaluator.
"""
from __future__ import annotations

import math
from typing import Callable, Dict, List

from .values import (
    DIV0, NA, NUM, VALUE, RangeVal, XlError, compare, criterion_cell_key, criterion_matches,
    excel_round, lookup_key, parse_criterion, to_bool, to_number, to_text,
)


class ArgError(Exception):
    """Raised to short-circuit a function with an Excel error value."""

    def __init__(self, err: XlError):
        self.err = err


def _num(v) -> float:
    if isinstance(v, RangeVal):
        raise ArgError(VALUE)
    n = to_number(v)
    if isinstance(n, XlError):
        raise ArgError(n)
    return n


def _text(v) -> str:
    if isinstance(v, RangeVal):
        raise ArgError(VALUE)
    t = to_text(v)
    if isinstance(t, XlError):
        raise ArgError(t)
    return t


def _numbers_for_agg(args) -> List[float]:
    """Collect numbers per Excel aggregate rules.

    Inside ranges only real numbers count (text, logicals, blanks are skipped);
    direct arguments are coerced. Errors anywhere propagate.
    """
    out: List[float] = []
    for a in args:
        if isinstance(a, RangeVal):
            for v in a.flat():
                if isinstance(v, XlError):
                    raise ArgError(v)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(float(v))
        elif a is None:
            continue
        else:
            out.append(_num(a))
    return out


def fn_sum(*args):
    return math.fsum(_numbers_for_agg(args))


def fn_average(*args):
    nums = _numbers_for_agg(args)
    return math.fsum(nums) / len(nums) if nums else DIV0


def fn_min(*args):
    nums = _numbers_for_agg(args)
    return min(nums) if nums else 0.0


def fn_max(*args):
    nums = _numbers_for_agg(args)
    return max(nums) if nums else 0.0


def fn_count(*args):
    n = 0
    for a in args:
        if isinstance(a, RangeVal):
            n += sum(1 for v in a.flat() if isinstance(v, (int, float)) and not isinstance(v, bool))
        elif not isinstance(to_number(a), XlError) and a is not None:
            n += 1
    return float(n)


def fn_counta(*args):
    n = 0
    for a in args:
        if isinstance(a, RangeVal):
            n += sum(1 for v in a.flat() if v is not None)
        elif a is not None:
            n += 1
    return float(n)


def fn_not(a):
    b = to_bool(a)
    if isinstance(b, XlError):
        return b
    return not b


def fn_round(x, d=0.0):
    return excel_round(_num(x), int(_num(d)))


def fn_roundup(x, d=0.0):
    x, d = _num(x), int(_num(d))
    f = 10.0 ** d
    return math.copysign(math.ceil(abs(x) * f - 1e-9) / f, x)


def fn_rounddown(x, d=0.0):
    x, d = _num(x), int(_num(d))
    f = 10.0 ** d
    return math.copysign(math.floor(abs(x) * f + 1e-9) / f, x)


def fn_abs(x):
    return abs(_num(x))


def fn_sqrt(x):
    x = _num(x)
    return NUM if x < 0 else math.sqrt(x)


def fn_int(x):
    return float(math.floor(_num(x)))


def fn_mod(a, b):
    a, b = _num(a), _num(b)
    if b == 0:
        return DIV0
    return a - b * math.floor(a / b)


def fn_power(a, b):
    return _pow(_num(a), _num(b))


def _pow(a: float, b: float):
    if a == 0 and b < 0:
        return DIV0
    try:
        r = a ** b
    except OverflowError:
        return NUM
    if isinstance(r, complex):
        return NUM
    return float(r)


def fn_exp(x):
    try:
        return math.exp(_num(x))
    except OverflowError:
        return NUM


def fn_ln(x):
    x = _num(x)
    return NUM if x <= 0 else math.log(x)


def fn_left(s, n=1.0):
    n = int(_num(n))
    return VALUE if n < 0 else _text(s)[:n]


def fn_right(s, n=1.0):
    n = int(_num(n))
    if n < 0:
        return VALUE
    t = _text(s)
    return t[len(t) - n:] if n else ""


def fn_mid(s, start, n):
    start, n = int(_num(start)), int(_num(n))
    if start < 1 or n < 0:
        return VALUE
    return _text(s)[start - 1:start - 1 + n]


def fn_len(s):
    return float(len(_text(s)))


def fn_upper(s):
    return _text(s).upper()


def fn_lower(s):
    return _text(s).lower()


def fn_trim(s):
    # Excel TRIM also collapses runs of internal spaces.
    return " ".join(p for p in _text(s).split(" ") if p)


def fn_concatenate(*args):
    return "".join(_text(a) for a in args)


def fn_isblank(v):
    return v is None


def fn_isnumber(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def fn_istext(v):
    return isinstance(v, str)


def fn_iserror(v):
    return isinstance(v, XlError)


# ---------------------------------------------------------- conditional sums

def _same_shape(a: RangeVal, b: RangeVal) -> bool:
    return a.height == b.height and a.width == b.width


def _criteria_index(ctx, sum_rng, pairs):
    """Hash index: tuple of criterion keys -> [sum of numbers, numeric count, row count].

    Built once per (ranges) per recalculation, so N SUMIF cells that scan the same
    N-row column cost O(N) total instead of O(N^2).
    """
    key = ("ifs-idx", sum_rng.key if sum_rng is not None else None,
           tuple(r.key for r, _ in pairs))
    idx = ctx.cache.get(key)
    if idx is not None:
        return idx
    idx = {}
    flats = [list(r.flat()) for r, _ in pairs]
    sums = list(sum_rng.flat()) if sum_rng is not None else None
    for i in range(len(flats[0])):
        k = tuple(criterion_cell_key(f[i]) for f in flats)
        slot = idx.get(k)
        if slot is None:
            slot = idx[k] = [0.0, 0, 0, None]
        slot[2] += 1
        if sums is not None:
            v = sums[i]
            if isinstance(v, XlError):
                if slot[3] is None:
                    slot[3] = v
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                slot[0] += v
                slot[1] += 1
    ctx.cache[key] = idx
    return idx


def _conditional(ctx, sum_rng, pairs, mode: str):
    for r, _ in pairs:
        if not isinstance(r, RangeVal):
            return VALUE
        if not _same_shape(r, pairs[0][0]):
            return VALUE
    if sum_rng is not None and (not isinstance(sum_rng, RangeVal) or not _same_shape(sum_rng, pairs[0][0])):
        return VALUE
    crits = [parse_criterion(c) for _, c in pairs]
    if ctx.use_index and all(c.index_key is not None for c in crits):
        idx = _criteria_index(ctx, sum_rng, pairs)
        slot = idx.get(tuple(c.index_key for c in crits))
        total, numeric, matched, err = slot if slot else (0.0, 0, 0, None)
    else:
        flats = [list(r.flat()) for r, _ in pairs]
        sums = list(sum_rng.flat()) if sum_rng is not None else None
        total, numeric, matched, err = 0.0, 0, 0, None
        for i in range(len(flats[0])):
            if all(criterion_matches(c, f[i]) for c, f in zip(crits, flats)):
                matched += 1
                if sums is not None:
                    v = sums[i]
                    if isinstance(v, XlError):
                        err = err or v
                    elif isinstance(v, (int, float)) and not isinstance(v, bool):
                        total += v
                        numeric += 1
    if mode == "count":
        return float(matched)
    if err is not None:
        return err
    if mode == "sum":
        return total
    return total / numeric if numeric else DIV0


def fn_sumif(ctx, rng, crit, sum_rng=None):
    return _conditional(ctx, sum_rng if sum_rng is not None else rng, [(rng, crit)], "sum")


def fn_averageif(ctx, rng, crit, avg_rng=None):
    return _conditional(ctx, avg_rng if avg_rng is not None else rng, [(rng, crit)], "avg")


def fn_countif(ctx, rng, crit):
    return _conditional(ctx, None, [(rng, crit)], "count")


def fn_sumifs(ctx, sum_rng, *rest):
    if len(rest) % 2 or not rest:
        return VALUE
    return _conditional(ctx, sum_rng, list(zip(rest[0::2], rest[1::2])), "sum")


def fn_countifs(ctx, *rest):
    if len(rest) % 2 or not rest:
        return VALUE
    return _conditional(ctx, None, list(zip(rest[0::2], rest[1::2])), "count")


# ----------------------------------------------------------------- lookups

def _exact_index(ctx, values: List, cache_key):
    idx = ctx.cache.get(cache_key) if ctx.use_index else None
    if idx is None:
        idx = {}
        for i, v in enumerate(values):
            idx.setdefault(lookup_key(v), i)
        if ctx.use_index:
            ctx.cache[cache_key] = idx
    return idx


def _approx_pos(values: List, x) -> int:
    """Binary search for the last position whose value <= x (sorted ascending)."""
    def same_kind(v) -> bool:
        if isinstance(x, str):
            return isinstance(v, str)
        return isinstance(v, (int, float)) and not isinstance(v, bool)

    lo, hi, best = 0, len(values) - 1, -1
    while lo <= hi:
        mid = (lo + hi) // 2
        v = values[mid]
        if same_kind(v) and compare("<=", v, x) is True:
            best, lo = mid, mid + 1
        else:
            hi = mid - 1
    return best


def fn_vlookup(ctx, x, table, col, approx=True):
    if isinstance(x, XlError):
        return x
    if not isinstance(table, RangeVal):
        return VALUE
    c = int(_num(col))
    if c < 1:
        return VALUE
    if c > table.width:
        return XlError("#REF!")
    first = [row[0] for row in table.rows]
    exact = to_bool(approx) is False
    if exact:
        pos = _exact_index(ctx, first, ("vl-idx", table.key)).get(lookup_key(x), -1)
    else:
        pos = _approx_pos(first, x)
    if pos < 0:
        return NA
    v = table.rows[pos][c - 1]
    return 0.0 if v is None else v


def fn_match(ctx, x, rng, mtype=1.0):
    if isinstance(x, XlError):
        return x
    if not isinstance(rng, RangeVal) or (rng.height > 1 and rng.width > 1):
        return NA
    values = list(rng.flat())
    t = int(_num(mtype))
    if t == 0:
        pos = _exact_index(ctx, values, ("match-idx", rng.key)).get(lookup_key(x), -1)
    elif t == 1:
        pos = _approx_pos(values, x)
    else:
        return NA  # descending match type -1 not supported
    return NA if pos < 0 else float(pos + 1)


def fn_index(rng, row, col=None):
    if not isinstance(rng, RangeVal):
        return VALUE
    r = int(_num(row))
    c = int(_num(col)) if col is not None else None
    if c is None:
        if rng.height == 1:
            r, c = 1, r
        else:
            c = 1
    if r < 1 or c < 1 or r > rng.height or c > rng.width:
        return XlError("#REF!")
    v = rng.rows[r - 1][c - 1]
    return 0.0 if v is None else v


FUNCTIONS: Dict[str, Callable] = {
    "SUM": fn_sum, "AVERAGE": fn_average, "MIN": fn_min, "MAX": fn_max,
    "COUNT": fn_count, "COUNTA": fn_counta, "NOT": fn_not,
    "ROUND": fn_round, "ROUNDUP": fn_roundup, "ROUNDDOWN": fn_rounddown,
    "ABS": fn_abs, "SQRT": fn_sqrt, "INT": fn_int, "MOD": fn_mod, "POWER": fn_power,
    "EXP": fn_exp, "LN": fn_ln,
    "LEFT": fn_left, "RIGHT": fn_right, "MID": fn_mid, "LEN": fn_len,
    "UPPER": fn_upper, "LOWER": fn_lower, "TRIM": fn_trim, "CONCATENATE": fn_concatenate,
    "ISBLANK": fn_isblank, "ISNUMBER": fn_isnumber, "ISTEXT": fn_istext, "ISERROR": fn_iserror,
    "INDEX": fn_index,
}

# Functions that take the evaluation context as their first argument.
CTX_FUNCTIONS: Dict[str, Callable] = {
    "SUMIF": fn_sumif, "AVERAGEIF": fn_averageif, "COUNTIF": fn_countif,
    "SUMIFS": fn_sumifs, "COUNTIFS": fn_countifs, "VLOOKUP": fn_vlookup, "MATCH": fn_match,
}

# Functions whose arguments may legitimately be ranges (no implicit intersection).
RANGE_ARG_FUNCTIONS = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "INDEX",
                       "SUMIF", "AVERAGEIF", "COUNTIF", "SUMIFS", "COUNTIFS", "VLOOKUP", "MATCH"}

# Pure aggregates whose result depends only on their range arguments: memoised per recalc.
MEMO_AGGREGATES = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"}

LAZY_FUNCTIONS = {"IF", "IFERROR", "AND", "OR"}

SUPPORTED = set(FUNCTIONS) | set(CTX_FUNCTIONS) | LAZY_FUNCTIONS
