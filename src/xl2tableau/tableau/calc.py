"""Tableau calculated-field expressions: AST, renderer, parser and evaluator.

The evaluator is a model of Tableau's row-level semantics, used to check that
a translated calculation gives the same answer as the Excel formula it came
from. It implements the parts that matter for that check: NULL propagation,
division by zero returning NULL, IF/ELSEIF/ELSE/END, and FIXED level-of-detail
expressions. It is not Tableau itself; see the README limitations.
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from ..formula.values import excel_round, number_to_text


# ------------------------------------------------------------------- AST

@dataclass(frozen=True)
class TLit:
    value: object  # float, str, bool or None


@dataclass(frozen=True)
class TField:
    name: str


@dataclass(frozen=True)
class TParam:
    name: str


@dataclass(frozen=True)
class TFunc:
    name: str
    args: Tuple = ()


@dataclass(frozen=True)
class TBin:
    op: str
    left: object
    right: object


@dataclass(frozen=True)
class TNot:
    operand: object


@dataclass(frozen=True)
class TNeg:
    operand: object


@dataclass(frozen=True)
class TIf:
    branches: Tuple  # ((cond, value), ...)
    else_: object = None


@dataclass(frozen=True)
class TLod:
    dims: Tuple[str, ...]
    agg: TFunc


AGGREGATES = {"SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTD"}


# -------------------------------------------------------------- rendering

def bracket(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


def _lit(v) -> str:
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return number_to_text(v) if v >= 0 else f"({number_to_text(v)})"
    return '"' + str(v).replace('"', '""') + '"'


def render(node, field: Callable[[str], str] = bracket,
           param: Callable[[str], str] = lambda n: "[Parameters]." + bracket(n)) -> str:
    def r(n):
        return render(n, field, param)

    if isinstance(node, TLit):
        return _lit(node.value)
    if isinstance(node, TField):
        return field(node.name)
    if isinstance(node, TParam):
        return param(node.name)
    if isinstance(node, TFunc):
        return f"{node.name}({', '.join(r(a) for a in node.args)})"
    if isinstance(node, TBin):
        return f"({r(node.left)} {node.op} {r(node.right)})"
    if isinstance(node, TNot):
        return f"NOT {r(node.operand)}"
    if isinstance(node, TNeg):
        return f"-{r(node.operand)}"
    if isinstance(node, TIf):
        parts = []
        for i, (c, v) in enumerate(node.branches):
            parts.append(f"{'IF' if i == 0 else 'ELSEIF'} {r(c)} THEN {r(v)}")
        if node.else_ is not None:
            parts.append(f"ELSE {r(node.else_)}")
        return " ".join(parts) + " END"
    if isinstance(node, TLod):
        dims = ", ".join(field(d) for d in node.dims)
        return "{FIXED " + (dims + " " if dims else "") + ": " + r(node.agg) + "}"
    raise TypeError(node)


# ---------------------------------------------------------------- parsing

class CalcParseError(ValueError):
    pass


_KEYWORDS = {"IF", "THEN", "ELSE", "ELSEIF", "END", "AND", "OR", "NOT", "TRUE", "FALSE",
             "NULL", "FIXED"}


def _tokenize(text: str):
    toks, i, n = [], 0, len(text)
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
        elif ch == "[":
            parts = []
            while True:
                j, buf = i + 1, []
                while True:
                    if j >= n:
                        raise CalcParseError("unterminated [name]")
                    if text[j] == "]":
                        if j + 1 < n and text[j + 1] == "]":
                            buf.append("]")
                            j += 2
                            continue
                        break
                    buf.append(text[j])
                    j += 1
                parts.append("".join(buf))
                i = j + 1
                if i + 1 < n and text[i] == "." and text[i + 1] == "[":
                    i += 1
                    continue
                break
            toks.append(("FIELD", tuple(parts)))
        elif ch in "\"'":
            j, buf = i + 1, []
            while True:
                if j >= n:
                    raise CalcParseError("unterminated string")
                if text[j] == ch:
                    if j + 1 < n and text[j + 1] == ch:
                        buf.append(ch)
                        j += 2
                        continue
                    break
                buf.append(text[j])
                j += 1
            toks.append(("STR", "".join(buf)))
            i = j + 1
        elif ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit()):
            j = i
            while j < n and (text[j].isdigit() or text[j] == "."):
                j += 1
            if j < n and text[j] in "eE":
                j += 1
                if j < n and text[j] in "+-":
                    j += 1
                while j < n and text[j].isdigit():
                    j += 1
            toks.append(("NUM", float(text[i:j])))
            i = j
        elif ch.isalpha() or ch == "_":
            j = i
            while j < n and (text[j].isalnum() or text[j] == "_"):
                j += 1
            word = text[i:j].upper()
            toks.append(("KW" if word in _KEYWORDS else "IDENT", word))
            i = j
        else:
            two = text[i:i + 2]
            if two in ("==", "!=", "<>", "<=", ">="):
                toks.append(("OP", two))
                i += 2
            elif ch in "+-*/%<>=(),{}:":
                toks.append(("OP", ch))
                i += 1
            else:
                raise CalcParseError(f"unexpected {ch!r}")
    return toks


class _P:
    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        if tok[0] is None:
            raise CalcParseError("unexpected end")
        self.i += 1
        return tok

    def expect(self, kind, val=None):
        tok = self.take()
        if tok[0] != kind or (val is not None and tok[1] != val):
            raise CalcParseError(f"expected {val or kind}, got {tok[1]!r}")
        return tok

    def accept(self, kind, val):
        if self.peek() == (kind, val):
            self.i += 1
            return True
        return False

    def expr(self):
        node = self.and_()
        while self.accept("KW", "OR"):
            node = TBin("OR", node, self.and_())
        return node

    def and_(self):
        node = self.not_()
        while self.accept("KW", "AND"):
            node = TBin("AND", node, self.not_())
        return node

    def not_(self):
        if self.accept("KW", "NOT"):
            return TNot(self.not_())
        return self.cmp()

    def cmp(self):
        node = self.add()
        while self.peek()[0] == "OP" and self.peek()[1] in ("==", "=", "!=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            op = {"=": "==", "<>": "!="}.get(op, op)
            node = TBin(op, node, self.add())
        return node

    def add(self):
        node = self.mul()
        while self.peek()[0] == "OP" and self.peek()[1] in "+-" and len(self.peek()[1]) == 1:
            op = self.take()[1]
            node = TBin(op, node, self.mul())
        return node

    def mul(self):
        node = self.unary()
        while self.peek()[0] == "OP" and self.peek()[1] in ("*", "/", "%"):
            op = self.take()[1]
            node = TBin(op, node, self.unary())
        return node

    def unary(self):
        if self.accept("OP", "-"):
            return TNeg(self.unary())
        return self.primary()

    def primary(self):
        kind, val = self.take()
        if kind == "NUM":
            return TLit(val)
        if kind == "STR":
            return TLit(val)
        if kind == "FIELD":
            if len(val) == 2 and val[0] == "Parameters":
                return TParam(val[1])
            return TField(val[-1])
        if kind == "KW" and val in ("TRUE", "FALSE"):
            return TLit(val == "TRUE")
        if kind == "KW" and val == "NULL":
            return TLit(None)
        if kind == "KW" and val == "IF":
            branches = [(self.expr(), None)]
            self.expect("KW", "THEN")
            branches[0] = (branches[0][0], self.expr())
            else_ = None
            while True:
                if self.accept("KW", "ELSEIF"):
                    c = self.expr()
                    self.expect("KW", "THEN")
                    branches.append((c, self.expr()))
                elif self.accept("KW", "ELSE"):
                    else_ = self.expr()
                    self.expect("KW", "END")
                    break
                else:
                    self.expect("KW", "END")
                    break
            return TIf(tuple(branches), else_)
        if kind == "OP" and val == "(":
            node = self.expr()
            self.expect("OP", ")")
            return node
        if kind == "OP" and val == "{":
            self.expect("KW", "FIXED")
            dims = []
            while self.peek()[0] == "FIELD":
                dims.append(self.take()[1][-1])
                if not self.accept("OP", ","):
                    break
            self.expect("OP", ":")
            agg = self.expr()
            self.expect("OP", "}")
            if not isinstance(agg, TFunc) or agg.name not in AGGREGATES:
                raise CalcParseError("LOD body must be an aggregate")
            return TLod(tuple(dims), agg)
        if kind == "IDENT":
            self.expect("OP", "(")
            args = []
            if not self.accept("OP", ")"):
                while True:
                    args.append(self.expr())
                    if self.accept("OP", ")"):
                        break
                    self.expect("OP", ",")
            return TFunc(val, tuple(args))
        raise CalcParseError(f"unexpected token {val!r}")


def parse_calc(text: str):
    p = _P(_tokenize(text))
    node = p.expr()
    if p.peek()[0] is not None:
        raise CalcParseError(f"trailing token {p.peek()[1]!r}")
    return node


# ------------------------------------------------------------- evaluation

class CalcEvalError(Exception):
    pass


def _round_half_away(x, n=0):
    return excel_round(x, int(n))


def _fn_str(x):
    if isinstance(x, bool):
        return "true" if x else "false"
    if isinstance(x, float):
        return number_to_text(x)
    return str(x)


def _safe(f):
    def g(*a):
        try:
            return f(*a)
        except (ValueError, OverflowError, ZeroDivisionError):
            return None
    return g


def _power(a, b):
    if a == 0 and b < 0:
        return None
    r = a ** b
    return None if isinstance(r, complex) else float(r)


_NULL_PROPAGATING = {
    "ABS": abs, "ROUND": _round_half_away, "SQRT": _safe(lambda x: math.sqrt(x) if x >= 0 else None),
    "FLOOR": lambda x: float(math.floor(x)), "CEILING": lambda x: float(math.ceil(x)),
    "POWER": _safe(_power), "EXP": _safe(math.exp), "LN": _safe(lambda x: math.log(x) if x > 0 else None),
    "INT": lambda x: float(int(x)), "FLOAT": float, "STR": _fn_str,
    "LEFT": lambda s, n: s[:max(int(n), 0)],
    "RIGHT": lambda s, n: s[len(s) - int(n):] if int(n) > 0 else "",
    "MID": lambda s, st, n: s[int(st) - 1:int(st) - 1 + int(n)],
    "LEN": lambda s: float(len(s)), "UPPER": str.upper, "LOWER": str.lower,
    "TRIM": str.strip,  # Tableau TRIM only strips the ends
}


class CalcEvaluator:
    """Evaluates calculated fields over an in-memory data source.

    ``rows`` holds the raw data columns (as Tableau would read them from the
    CSV, with NULL for blanks). ``calcs`` maps calc field name to its AST.
    """

    def __init__(self, rows: List[Dict[str, object]], calcs: Dict[str, object],
                 params: Optional[Dict[str, object]] = None):
        self.rows = rows
        self.calcs = calcs
        self.params = params or {}
        self._memo: Dict[Tuple[str, int], object] = {}
        self._lod: Dict[object, Dict] = {}
        self._active: set = set()

    def field(self, name: str, i: int):
        if name in self.calcs:
            key = (name, i)
            if key in self._memo:
                return self._memo[key]
            if key in self._active:
                raise CalcEvalError(f"circular calculation {name}")
            self._active.add(key)
            try:
                v = self.eval(self.calcs[name], i)
            finally:
                self._active.discard(key)
            self._memo[key] = v
            return v
        row = self.rows[i]
        if name not in row:
            raise CalcEvalError(f"unknown field [{name}]")
        return row[name]

    def column(self, name: str) -> List[object]:
        return [self.field(name, i) for i in range(len(self.rows))]

    def eval(self, node, i: int):
        if isinstance(node, TLit):
            return node.value
        if isinstance(node, TField):
            return self.field(node.name, i)
        if isinstance(node, TParam):
            if node.name not in self.params:
                raise CalcEvalError(f"unknown parameter {node.name}")
            return self.params[node.name]
        if isinstance(node, TNeg):
            v = self.eval(node.operand, i)
            return None if v is None else -v
        if isinstance(node, TNot):
            v = self.eval(node.operand, i)
            return None if v is None else (not v)
        if isinstance(node, TIf):
            for cond, val in node.branches:
                if self.eval(cond, i) is True:
                    return self.eval(val, i)
            return None if node.else_ is None else self.eval(node.else_, i)
        if isinstance(node, TBin):
            return self._bin(node, i)
        if isinstance(node, TFunc):
            return self._func(node, i)
        if isinstance(node, TLod):
            return self._lod_value(node, i)
        raise CalcEvalError(f"cannot evaluate {node!r}")

    def _bin(self, node: TBin, i: int):
        op = node.op
        a = self.eval(node.left, i)
        if op == "AND":
            if a is False:
                return False
            b = self.eval(node.right, i)
            if b is False:
                return False
            return None if a is None or b is None else True
        if op == "OR":
            if a is True:
                return True
            b = self.eval(node.right, i)
            if b is True:
                return True
            return None if a is None or b is None else False
        b = self.eval(node.right, i)
        if a is None or b is None:
            return None
        if op == "+":
            if isinstance(a, str) != isinstance(b, str):
                raise CalcEvalError("cannot add string and number")
            return a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if op == "/":
            return None if b == 0 else a / b
        if op == "%":
            return None if b == 0 else math.fmod(a, b)
        if isinstance(a, str) != isinstance(b, str):
            raise CalcEvalError("cannot compare string and number")
        return {"==": a == b, "!=": a != b, "<": a < b, ">": a > b,
                "<=": a <= b, ">=": a >= b}[op]

    def _func(self, node: TFunc, i: int):
        name = node.name
        if name in AGGREGATES and not (name in ("MIN", "MAX") and len(node.args) == 2):
            raise CalcEvalError(f"aggregate {name} outside an LOD is not supported in row-level checks")
        args = [self.eval(a, i) for a in node.args]
        if name == "ZN":
            return 0.0 if args[0] is None else args[0]
        if name == "IFNULL":
            return args[1] if args[0] is None else args[0]
        if name == "ISNULL":
            return args[0] is None
        if name == "IIF":
            return args[1] if args[0] is True else args[2] if args[0] is False else None
        if name in ("MIN", "MAX") and len(args) == 2:
            if None in args:
                return None
            return min(args) if name == "MIN" else max(args)
        fn = _NULL_PROPAGATING.get(name)
        if fn is None:
            raise CalcEvalError(f"unsupported function {name}")
        if any(a is None for a in args):
            return None
        return fn(*args)

    def _lod_value(self, node: TLod, i: int):
        table = self._lod.get(node)
        if table is None:
            groups: Dict[Tuple, List] = defaultdict(list)
            inner = node.agg.args[0] if node.agg.args else TLit(1.0)
            for j in range(len(self.rows)):
                k = tuple(self.field(d, j) for d in node.dims)
                groups[k].append(self.eval(inner, j))
            table = {k: _aggregate(node.agg.name, vals) for k, vals in groups.items()}
            self._lod[node] = table
        return table[tuple(self.field(d, i) for d in node.dims)]


def _aggregate(name: str, vals: Sequence):
    present = [v for v in vals if v is not None]
    if name == "COUNT":
        return float(len(present))
    if name == "COUNTD":
        return float(len(set(present)))
    if not present:
        return None
    if name == "SUM":
        return math.fsum(present)
    if name == "AVG":
        return math.fsum(present) / len(present)
    if name == "MIN":
        return min(present)
    if name == "MAX":
        return max(present)
    raise CalcEvalError(name)


def field_refs(node) -> List[str]:
    """All field names referenced by a calc (including LOD dimensions)."""
    out: List[str] = []

    def visit(n):
        if isinstance(n, TField):
            out.append(n.name)
        elif isinstance(n, TFunc):
            for a in n.args:
                visit(a)
        elif isinstance(n, TBin):
            visit(n.left)
            visit(n.right)
        elif isinstance(n, (TNot, TNeg)):
            visit(n.operand)
        elif isinstance(n, TIf):
            for c, v in n.branches:
                visit(c)
                visit(v)
            if n.else_ is not None:
                visit(n.else_)
        elif isinstance(n, TLod):
            out.extend(n.dims)
            visit(n.agg)

    visit(node)
    return out
