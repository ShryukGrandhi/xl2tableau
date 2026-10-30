"""Translate Excel calc-column formulas into Tableau calculated fields.

The input is the formula in the anchor (first) row of a calc column. A
reference to another column in the same row becomes a field reference, a
fixed cell outside the table becomes a Tableau parameter, and aggregates over
whole table columns become FIXED level-of-detail expressions:

    =F2/SUM($F$2:$F$500)                -> [Revenue] / {FIXED : SUM([Revenue])}
    =SUMIF($A$2:$A$500,A2,$F$2:$F$500)  -> {FIXED [Region] : SUM([Revenue])}

Anything without a faithful equivalent raises Untranslatable with a reason,
so the report says exactly what needs manual work.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from ..engine import Workbook
from ..formula.ast import MAX_ROW, Binary, Bool, Err, Func, Name, Num, Percent, Ref, Str, Unary, index_to_col
from ..tables import Column, Table
from .calc import TBin, TField, TFunc, TIf, TLit, TLod, TNeg, TNot, TParam

NUMBER, STRING, BOOLEAN = "number", "string", "boolean"


class Untranslatable(Exception):
    pass


@dataclass
class Parameter:
    name: str
    sheet: str
    row: int
    col: int
    value: object

    @property
    def dtype(self) -> str:
        if isinstance(self.value, bool):
            return "boolean"
        if isinstance(self.value, float):
            return "real"
        return "string"


@dataclass
class Translation:
    column: Column
    ok: bool
    tnode: object = None
    ttype: Optional[str] = None
    reason: Optional[str] = None
    notes: List[str] = field(default_factory=list)


class ParameterRegistry:
    """Assigns stable, readable names to the input cells that become parameters."""

    def __init__(self, wb: Workbook):
        self.wb = wb
        self.by_cell: Dict[Tuple[str, int, int], Parameter] = {}
        self._names = set()
        self._defined = {}
        for name, ref in wb.defined_names.items():
            if ref.is_cell:
                self._defined[(ref.sheet, ref.row1, ref.col1)] = name

    def get(self, sheet: str, row: int, col: int, preferred: Optional[str] = None) -> Parameter:
        key = (sheet, row, col)
        if key in self.by_cell:
            return self.by_cell[key]
        name = preferred or self._defined.get(key)
        if name is None:
            label = self.wb.get(sheet, row, col - 1) if col > 1 else None
            if isinstance(label, str) and label.strip():
                name = label.strip().rstrip(":").strip()
            else:
                name = f"{sheet}!{index_to_col(col)}{row}"
        base, k = name, 2
        while name in self._names:
            name = f"{base} ({k})"
            k += 1
        self._names.add(name)
        p = Parameter(name, sheet, row, col, self.wb.get(sheet, row, col))
        self.by_cell[key] = p
        return p


_SIMPLE_NUM = {"ABS": "ABS", "SQRT": "SQRT", "EXP": "EXP", "LN": "LN", "INT": "FLOOR"}
_SIMPLE_STR = {"UPPER": "UPPER", "LOWER": "LOWER", "TRIM": "TRIM"}
_AGG = {"SUM": "SUM", "AVERAGE": "AVG", "MIN": "MIN", "MAX": "MAX", "COUNT": "COUNT", "COUNTA": "COUNT"}
_CMP = {"=": "==", "<>": "!=", "<": "<", ">": ">", "<=": "<=", ">=": ">="}


class TableTranslator:
    def __init__(self, wb: Workbook, table: Table, params: ParameterRegistry):
        self.wb = wb
        self.table = table
        self.params = params
        self.results: Dict[str, Translation] = {}
        self._in_progress: set = set()

    # ---------------------------------------------------------- public API
    def translate_all(self) -> List[Translation]:
        for col in self.table.columns:
            if col.kind == "calc":
                self.translate_column(col)
        return [self.results[c.name] for c in self.table.columns if c.kind == "calc"]

    def translate_column(self, col: Column) -> Translation:
        if col.name in self.results:
            return self.results[col.name]
        if col.ast is None:
            t = Translation(col, False, reason="formula could not be parsed")
            self.results[col.name] = t
            return t
        self._in_progress.add(col.name)
        notes: List[str] = []
        try:
            anchor = self.table.first_row + next(
                i for i in range(self.table.n_rows)
                if (self.table.first_row + i) not in col.inconsistent_rows)
            node, ttype = self._tr(col.ast, anchor, col.col, notes)
            t = Translation(col, True, node, ttype, notes=notes)
        except Untranslatable as e:
            t = Translation(col, False, reason=str(e), notes=notes)
        finally:
            self._in_progress.discard(col.name)
        self.results[col.name] = t
        return t

    # ------------------------------------------------------------- helpers
    def _in_table(self, sheet: str, col: int) -> bool:
        return sheet == self.table.sheet and self.table.first_col <= col <= self.table.last_col

    def _column_range(self, ref: Ref) -> Optional[Column]:
        """The table column a range covers, if it is exactly one whole column of data."""
        t = self.table
        sheet = ref.sheet or t.sheet
        if sheet != t.sheet or ref.col1 != ref.col2 or not self._in_table(sheet, ref.col1):
            return None
        if ref.row1 not in (t.header_row, t.first_row):
            return None
        if ref.row2 < t.last_row:
            return None
        if ref.row2 != MAX_ROW:
            # Cells below the table inside the range must be empty.
            for r in range(t.last_row + 1, ref.row2 + 1):
                if self.wb.get(sheet, r, ref.col1) is not None:
                    return None
        return t.column_at(ref.col1)

    def _field(self, col: Column, raw: bool = False):
        """A field reference. Nullable data columns are wrapped so blanks behave like Excel."""
        if col.kind == "calc":
            if col.name in self._in_progress:
                raise Untranslatable(f"circular reference through column '{col.name}'")
            t = self.translate_column(col)
            if not t.ok:
                raise Untranslatable(f"depends on untranslatable column '{col.name}'")
            return TField(col.name), t.ttype
        ttype = {"integer": NUMBER, "real": NUMBER, "string": STRING, "boolean": BOOLEAN}[col.dtype]
        node = TField(col.name)
        if col.nullable and not raw:
            node = {NUMBER: lambda n: TFunc("ZN", (n,)),
                    STRING: lambda n: TFunc("IFNULL", (n, TLit(""))),
                    BOOLEAN: lambda n: TFunc("IFNULL", (n, TLit(False)))}[ttype](node)
        return node, ttype

    def _cell_ref(self, ref: Ref, row: Optional[int], ctx_col: int, notes, raw=False, defined=None):
        sheet = ref.sheet or self.table.sheet
        t = self.table
        r, c = ref.row1, ref.col1
        if self._in_table(sheet, c) and t.header_row <= r <= t.last_row:
            if row is None:
                raise Untranslatable("a cell outside the table refers to a single row of the table")
            if r != row or ref.abs_row1:
                raise Untranslatable(
                    f"refers to another row ({ref.a1()}); running totals and offsets need Tableau table calculations")
            col = t.column_at(c)
            return self._field(col, raw)
        if not ref.abs_row1 and row is not None:
            raise Untranslatable(f"relative reference {ref.a1()} outside the table changes per row")
        sh = self.wb.sheets.get(sheet)
        if sh is None:
            raise Untranslatable(f"reference to missing sheet {sheet!r}")
        if (r, c) in sh.formulas:
            # A computed cell outside the table (e.g. a grand total): translate its
            # formula without row context so it can only use table-wide aggregates.
            from ..formula.parser import parse

            notes.append(f"inlined computed cell {sheet}!{index_to_col(c)}{r}")
            return self._tr(parse(sh.formulas[(r, c)]), None, ctx_col, notes)
        p = self.params.get(sheet, r, c, defined)
        if p.value is None:
            raise Untranslatable(f"reference to empty cell {ref.a1()}")
        ttype = {"real": NUMBER, "string": STRING, "boolean": BOOLEAN}[p.dtype]
        return TParam(p.name), ttype

    def _as_number(self, pair):
        node, ttype = pair
        if ttype == NUMBER:
            return node
        if ttype == BOOLEAN:
            return TFunc("INT", (node,))
        raise Untranslatable("text used in arithmetic")

    def _as_string(self, pair):
        node, ttype = pair
        if ttype == STRING:
            return node
        if ttype == NUMBER:
            return TFunc("STR", (node,))
        return TIf(((node, TLit("TRUE")),), TLit("FALSE"))

    def _as_bool(self, pair):
        node, ttype = pair
        if ttype == BOOLEAN:
            return node
        if ttype == NUMBER:
            return TBin("!=", node, TLit(0.0))
        raise Untranslatable("text used as a condition")

    # ------------------------------------------------------------ translate
    def _tr(self, node, row: Optional[int], ccol: int, notes) -> Tuple[object, str]:
        tr = lambda n: self._tr(n, row, ccol, notes)  # noqa: E731
        if isinstance(node, Num):
            return TLit(node.value), NUMBER
        if isinstance(node, Str):
            return TLit(node.value), STRING
        if isinstance(node, Bool):
            return TLit(node.value), BOOLEAN
        if isinstance(node, Err):
            raise Untranslatable(f"error literal {node.code}")
        if isinstance(node, Name):
            ref = self.wb.resolve_name(node.name)
            if ref is None or not ref.is_cell:
                raise Untranslatable(f"defined name {node.name} is not a single cell")
            return self._cell_ref(ref, row, ccol, notes, defined=node.name)
        if isinstance(node, Ref):
            if not node.is_cell:
                raise Untranslatable(f"range {node.a1()} used outside an aggregate")
            return self._cell_ref(node, row, ccol, notes)
        if isinstance(node, Unary):
            inner = self._as_number(tr(node.operand))
            return (TNeg(inner) if node.op == "-" else inner), NUMBER
        if isinstance(node, Percent):
            return TBin("/", self._as_number(tr(node.operand)), TLit(100.0)), NUMBER
        if isinstance(node, Binary):
            return self._binary(node, tr)
        if isinstance(node, Func):
            return self._func(node, row, ccol, notes, tr)
        raise Untranslatable(f"unsupported syntax {type(node).__name__}")

    def _binary(self, node: Binary, tr):
        op = node.op
        if op == "&":
            return TBin("+", self._as_string(tr(node.left)), self._as_string(tr(node.right))), STRING
        if op in _CMP:
            (a, ta), (b, tb) = tr(node.left), tr(node.right)
            if ta != tb:
                if {ta, tb} == {NUMBER, BOOLEAN}:
                    raise Untranslatable("comparison between a number and TRUE/FALSE")
                raise Untranslatable("comparison between text and a number")
            return TBin(_CMP[op], a, b), BOOLEAN
        a, b = self._as_number(tr(node.left)), self._as_number(tr(node.right))
        if op == "^":
            return TFunc("POWER", (a, b)), NUMBER
        return TBin(op, a, b), NUMBER

    def _func(self, node: Func, row, ccol, notes, tr):
        name, args = node.name, node.args
        if name == "IF":
            if len(args) < 2:
                raise Untranslatable("IF with fewer than 2 arguments")
            cond = self._as_bool(tr(args[0]))
            then, tt = tr(args[1])
            if len(args) == 3:
                els, te = tr(args[2])
            else:
                els, te = TLit(False), BOOLEAN
            if tt != te:
                raise Untranslatable("IF branches return different types (Tableau needs one type)")
            # Flatten nested IFs in the else branch into ELSEIF chains.
            branches = ((cond, then),)
            if isinstance(els, TIf):
                return TIf(branches + els.branches, els.else_), tt
            return TIf(branches, els), tt
        if name == "IFERROR":
            if len(args) != 2:
                raise Untranslatable("IFERROR needs 2 arguments")
            (a, ta), (b, tb) = tr(args[0]), tr(args[1])
            if ta != tb:
                raise Untranslatable("IFERROR value and fallback have different types")
            notes.append("IFERROR -> IFNULL: only errors that Tableau represents as NULL are caught")
            return TFunc("IFNULL", (a, b)), ta
        if name in ("AND", "OR"):
            if not args:
                raise Untranslatable(f"{name}() with no arguments")
            parts = [self._as_bool(tr(a)) for a in args]
            out = parts[0]
            for p in parts[1:]:
                out = TBin(name, out, p)
            return out, BOOLEAN
        if name == "NOT":
            return TNot(self._as_bool(tr(args[0]))), BOOLEAN
        if name in _SIMPLE_NUM and len(args) == 1:
            return TFunc(_SIMPLE_NUM[name], (self._as_number(tr(args[0])),)), NUMBER
        if name == "ROUND":
            digits = self._as_number(tr(args[1])) if len(args) > 1 else TLit(0.0)
            return TFunc("ROUND", (self._as_number(tr(args[0])), digits)), NUMBER
        if name == "POWER" and len(args) == 2:
            return TFunc("POWER", (self._as_number(tr(args[0])), self._as_number(tr(args[1])))), NUMBER
        if name == "MOD" and len(args) == 2:
            # Excel MOD takes the sign of the divisor; Tableau's % takes the sign of the
            # dividend. a - b*FLOOR(a/b) reproduces Excel for negative inputs.
            a, b = self._as_number(tr(args[0])), self._as_number(tr(args[1]))
            return TBin("-", a, TBin("*", b, TFunc("FLOOR", (TBin("/", a, b),)))), NUMBER
        if name in _SIMPLE_STR and len(args) == 1:
            return TFunc(_SIMPLE_STR[name], (self._as_string(tr(args[0])),)), STRING
        if name == "LEN" and len(args) == 1:
            return TFunc("LEN", (self._as_string(tr(args[0])),)), NUMBER
        if name in ("LEFT", "RIGHT") and 1 <= len(args) <= 2:
            n = self._as_number(tr(args[1])) if len(args) == 2 else TLit(1.0)
            return TFunc(name, (self._as_string(tr(args[0])), n)), STRING
        if name == "MID" and len(args) == 3:
            return TFunc("MID", (self._as_string(tr(args[0])), self._as_number(tr(args[1])),
                                 self._as_number(tr(args[2])))), STRING
        if name == "CONCATENATE" and args:
            parts = [self._as_string(tr(a)) for a in args]
            out = parts[0]
            for p in parts[1:]:
                out = TBin("+", out, p)
            return out, STRING
        if name == "ISBLANK" and len(args) == 1 and isinstance(args[0], Ref) and args[0].is_cell:
            node_, _ = self._cell_ref(args[0], row, ccol, notes, raw=True)
            return TFunc("ISNULL", (node_,)), BOOLEAN
        if name in _AGG:
            return self._aggregate(name, args, tr)
        if name in ("SUMIF", "COUNTIF", "AVERAGEIF"):
            return self._conditional_single(name, args, row, tr)
        if name in ("SUMIFS", "COUNTIFS"):
            return self._conditional_multi(name, args, row)
        if name in ("VLOOKUP", "MATCH", "INDEX"):
            raise Untranslatable(f"{name} needs a Tableau relationship or join to the lookup range")
        raise Untranslatable(f"function {name} has no translation")

    def _aggregate(self, name, args, tr):
        if all(not (isinstance(a, Ref) and not a.is_cell) for a in args):
            parts = [self._as_number(tr(a)) for a in args]
            if name == "SUM":
                out = parts[0]
                for p in parts[1:]:
                    out = TBin("+", out, p)
                return out, NUMBER
            if name in ("MIN", "MAX") and len(parts) >= 2:
                out = parts[0]
                for p in parts[1:]:
                    out = TFunc(name, (out, p))
                return out, NUMBER
            raise Untranslatable(f"{name} over individual cells")
        if len(args) != 1:
            raise Untranslatable(f"{name} over several ranges")
        col = self._column_range(args[0])
        if col is None:
            raise Untranslatable(f"{name} over {args[0].a1()}, which is not exactly one table column")
        fld, _ = self._field(col, raw=True)
        return TLod((), TFunc(_AGG[name], (fld,))), NUMBER

    def _criteria_column(self, rng, what: str) -> Column:
        if not isinstance(rng, Ref) or rng.is_cell:
            raise Untranslatable(f"{what} is not a range")
        col = self._column_range(rng)
        if col is None:
            raise Untranslatable(f"{what} {rng.a1()} is not exactly one table column")
        return col

    def _conditional_single(self, name, args, row, tr):
        if len(args) < 2:
            raise Untranslatable(f"{name} needs at least 2 arguments")
        crit_col = self._criteria_column(args[0], "criteria range")
        value_col = crit_col
        if name != "COUNTIF" and len(args) == 3:
            value_col = self._criteria_column(args[2], "sum range")
        crit = args[1]
        same_row = (isinstance(crit, Ref) and crit.is_cell and row is not None
                    and (crit.sheet or self.table.sheet) == self.table.sheet
                    and crit.row1 == row and not crit.abs_row1)
        agg = {"SUMIF": "SUM", "COUNTIF": "COUNT", "AVERAGEIF": "AVG"}[name]
        if same_row:
            if crit.col1 != crit_col.col:
                raise Untranslatable(
                    f"{name} criterion compares column '{crit_col.name}' to a different column; "
                    "FIXED cannot express a cross-column match")
            dim, _ = self._field(crit_col, raw=True)
            val, _ = self._field(value_col, raw=True)
            return TLod((dim.name,), TFunc(agg, (val,))), NUMBER
        if isinstance(crit, (Num, Str)):
            dim, dtype = self._field(crit_col, raw=True)
            lit = TLit(crit.value)
            if (dtype == NUMBER) != isinstance(crit.value, float):
                raise Untranslatable(f"{name} criterion type does not match column '{crit_col.name}'")
            if isinstance(crit.value, str) and any(ch in crit.value for ch in "*?<>="):
                raise Untranslatable(f"{name} criterion {crit.value!r} uses operators or wildcards")
            val, _ = self._field(value_col, raw=True)
            inner = TIf(((TBin("==", dim, lit), TLit(1.0) if agg == "COUNT" else val),), None)
            return TLod((), TFunc("SUM" if agg == "COUNT" else agg, (inner,))), NUMBER
        raise Untranslatable(f"{name} criterion must be the same row of the criteria column or a constant")

    def _conditional_multi(self, name, args, row):
        rest = args[1:] if name == "SUMIFS" else args
        if not rest or len(rest) % 2:
            raise Untranslatable(f"{name} has unpaired criteria")
        dims = []
        for rng, crit in zip(rest[0::2], rest[1::2]):
            col = self._criteria_column(rng, "criteria range")
            if not (isinstance(crit, Ref) and crit.is_cell and row is not None and crit.row1 == row
                    and not crit.abs_row1 and crit.col1 == col.col
                    and (crit.sheet or self.table.sheet) == self.table.sheet):
                raise Untranslatable(f"{name} criteria must each be the same row of their own column")
            dims.append(self._field(col, raw=True)[0].name)
        if name == "SUMIFS":
            val, _ = self._field(self._criteria_column(args[0], "sum range"), raw=True)
            return TLod(tuple(dims), TFunc("SUM", (val,))), NUMBER
        return TLod(tuple(dims), TFunc("COUNT", (TField(dims[0]),))), NUMBER
