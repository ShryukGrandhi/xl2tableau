"""Find tabular regions in a workbook and classify their columns.

A column is a *data* column (exported to CSV) or a *calc* column (becomes a
Tableau calculated field). Calc columns are recognised by converting every
formula to relative R1C1 notation: ``=C2*D2`` in row 2 and ``=C3*D3`` in row 3
both become ``RC[-2]*RC[-1]``. The most common signature is the column's
pattern; rows that deviate are reported as inconsistent formulas.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .engine import Workbook
from .formula.ast import Binary, Bool, Err, Func, Name, Num, Percent, Ref, Str, Unary, col_to_index
from .formula.parser import FormulaParseError, parse
from .formula.values import XlError


@dataclass
class Column:
    name: str
    col: int
    kind: str  # "data" or "calc"
    dtype: str  # "integer" | "real" | "string" | "boolean"
    nullable: bool = False
    formula: Optional[str] = None  # formula text in the anchor (first data) row
    ast: object = None
    signature: Optional[str] = None
    inconsistent_rows: List[int] = field(default_factory=list)


@dataclass
class Table:
    sheet: str
    name: str
    header_row: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    columns: List[Column]

    def column_at(self, col: int) -> Optional[Column]:
        for c in self.columns:
            if c.col == col:
                return c
        return None

    @property
    def n_rows(self) -> int:
        return self.last_row - self.first_row + 1


def to_r1c1(node, row: int, col: int) -> str:
    """Render an AST with references relative to (row, col)."""
    def r(n):
        return to_r1c1(n, row, col)

    if isinstance(node, Num):
        return repr(node.value)
    if isinstance(node, Str):
        return '"' + node.value.replace('"', '""') + '"'
    if isinstance(node, Bool):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, Err):
        return node.code
    if isinstance(node, Name):
        return node.name.upper()
    if isinstance(node, Ref):
        def rc(rr, cc, ar, ac):
            rs = f"R{rr}" if ar else ("R" if rr == row else f"R[{rr - row}]")
            cs = f"C{cc}" if ac else ("C" if cc == col else f"C[{cc - col}]")
            return rs + cs

        prefix = f"'{node.sheet}'!" if node.sheet else ""
        if node.whole_col:
            a = f"C{node.col1}" if node.abs_col1 else f"C[{node.col1 - col}]"
            b = f"C{node.col2}" if node.abs_col2 else f"C[{node.col2 - col}]"
            return f"{prefix}{a}:{b}"
        first = rc(node.row1, node.col1, node.abs_row1, node.abs_col1)
        if node.is_cell:
            return prefix + first
        return f"{prefix}{first}:{rc(node.row2, node.col2, node.abs_row2, node.abs_col2)}"
    if isinstance(node, Func):
        return f"{node.name}({','.join(r(a) for a in node.args)})"
    if isinstance(node, Unary):
        return f"{node.op}{r(node.operand)}"
    if isinstance(node, Percent):
        return f"{r(node.operand)}%"
    if isinstance(node, Binary):
        return f"({r(node.left)}{node.op}{r(node.right)})"
    raise TypeError(node)


def infer_dtype(values) -> (str, bool):
    """Return (Tableau datatype, nullable) for a list of cell values."""
    kinds = set()
    nullable = False
    all_int = True
    for v in values:
        if v is None or v == "":
            nullable = True
        elif isinstance(v, XlError):
            nullable = True  # errors arrive in Tableau as NULL
        elif isinstance(v, bool):
            kinds.add("boolean")
        elif isinstance(v, (int, float)):
            kinds.add("number")
            if float(v) != int(v):
                all_int = False
        else:
            kinds.add("string")
    if not kinds:
        return "string", True
    if len(kinds) > 1:
        return "string", nullable
    k = kinds.pop()
    if k == "number":
        return ("integer" if all_int else "real"), nullable
    return k, nullable


_RANGE_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)")


def _cell(wb: Workbook, sheet: str, r: int, c: int):
    sh = wb.sheets[sheet]
    if (r, c) in sh.formulas:
        return sh.formulas[(r, c)]
    return sh.constants.get((r, c))


def _guess_region(wb: Workbook, sheet: str):
    """Heuristic: first row with >= 2 adjacent text cells followed by a non-empty row."""
    sh = wb.sheets[sheet]
    max_row, max_col = sh.max_row, sh.max_col
    for r in range(1, max_row):
        c = 1
        while c <= max_col:
            if isinstance(sh.constants.get((r, c)), str):
                c2 = c
                while isinstance(sh.constants.get((r, c2 + 1)), str):
                    c2 += 1
                below = any(_cell(wb, sheet, r + 1, k) is not None for k in range(c, c2 + 1))
                if c2 > c and below:
                    last = r + 1
                    while last + 1 <= max_row and any(
                            _cell(wb, sheet, last + 1, k) is not None for k in range(c, c2 + 1)):
                        last += 1
                    return r, last, c, c2
                c = c2 + 1
            else:
                c += 1
    return None


def detect_tables(wb: Workbook) -> List[Table]:
    """Find tables (Excel table objects first, else a heuristic per sheet).

    The workbook must already be recalculated so calc-column types can be inferred.
    """
    regions = []
    for sheet, name, ref in wb.excel_tables:
        m = _RANGE_RE.fullmatch(ref)
        if m:
            regions.append((sheet, name, int(m.group(2)), int(m.group(4)),
                            col_to_index(m.group(1)), col_to_index(m.group(3))))
    sheets_with_tables = {r[0] for r in regions}
    for sheet in wb.sheets:
        if sheet in sheets_with_tables:
            continue
        g = _guess_region(wb, sheet)
        if g:
            regions.append((sheet, sheet, *g))
    return [_build_table(wb, *reg) for reg in regions]


def _build_table(wb: Workbook, sheet: str, name: str, header: int, last: int, c1: int, c2: int) -> Table:
    sh = wb.sheets[sheet]
    first = header + 1
    columns: List[Column] = []
    used_names: Dict[str, int] = {}
    for c in range(c1, c2 + 1):
        raw = sh.constants.get((header, c))
        cname = str(raw).strip() if raw is not None else f"Column{c}"
        if cname in used_names:
            used_names[cname] += 1
            cname = f"{cname} ({used_names[cname]})"
        else:
            used_names[cname] = 1
        sigs = Counter()
        row_sig = {}
        for r in range(first, last + 1):
            text = sh.formulas.get((r, c))
            if text is None:
                continue
            ast = wb.asts.get((sheet, r, c))  # parsed during recalc
            if ast is None:
                try:
                    ast = parse(text)
                except FormulaParseError:
                    ast = None
            sig = to_r1c1(ast, r, c) if ast is not None else "<parse error>"
            sigs[sig] += 1
            row_sig[r] = sig
        values = [wb.get(sheet, r, c) for r in range(first, last + 1)]
        dtype, nullable = infer_dtype(values)
        if not sigs:
            columns.append(Column(cname, c, "data", dtype, nullable))
            continue
        signature, _ = sigs.most_common(1)[0]
        anchor = min(r for r, s in row_sig.items() if s == signature)
        text = sh.formulas[(anchor, c)]
        bad = [r for r in range(first, last + 1) if row_sig.get(r) != signature]
        columns.append(Column(cname, c, "calc", dtype, nullable, formula=text,
                              ast=wb.asts.get((sheet, anchor, c)) or (
                                  parse(text) if signature != "<parse error>" else None),
                              signature=signature, inconsistent_rows=bad))
    return Table(sheet, name, header, first, last, c1, c2, columns)
