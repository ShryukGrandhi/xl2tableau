"""Check translated calculated fields against the Excel engine, row by row.

The check reads the generated .twbx back from disk (XML and CSV), parses each
calculated-field formula, evaluates it over the CSV rows with the Tableau
semantics model, and compares every row with the value the Excel engine
computed for the corresponding cell.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional

from .engine import Workbook
from .formula.values import XlError
from .tables import Table
from .tableau.calc import CalcEvalError, CalcEvaluator, parse_calc
from .tableau.twb import load_twbx, read_csv_typed


def values_agree(excel, tableau) -> bool:
    if isinstance(excel, XlError):
        return tableau is None  # Tableau has no error values; failures become NULL
    if tableau is None:
        return excel == ""
    if isinstance(excel, bool) or isinstance(tableau, bool):
        return isinstance(excel, bool) and isinstance(tableau, bool) and excel == tableau
    if isinstance(excel, float) and isinstance(tableau, (int, float)):
        return math.isclose(excel, tableau, rel_tol=1e-9, abs_tol=1e-9)
    return excel == tableau


@dataclass
class FieldCheck:
    table: str
    field: str
    rows: int = 0
    matched: int = 0
    error: Optional[str] = None
    mismatches: List[tuple] = field(default_factory=list)  # (excel row, excel value, tableau value)

    @property
    def ok(self) -> bool:
        return self.error is None and self.matched == self.rows


def validate_twbx(path: str, wb: Workbook, tables: List[Table], keep: int = 5) -> List[FieldCheck]:
    parsed, csvs = load_twbx(path)
    by_name = {t.name: t for t in tables}
    checks: List[FieldCheck] = []
    for ds in parsed.datasources:
        table = by_name[ds.caption]
        rows = read_csv_typed(csvs[ds.csv_file], ds.relation_columns)
        calcs = {}
        parse_errors = {}
        for internal, formula in ds.calcs.items():
            try:
                calcs[internal] = parse_calc(formula)
            except ValueError as e:
                parse_errors[internal] = f"formula does not parse: {e}"
        ev = CalcEvaluator(rows, calcs, parsed.params)
        for internal in ds.calcs:
            caption = ds.captions[internal]
            col = next(c for c in table.columns if c.name == caption)
            fc = FieldCheck(table.name, caption, rows=len(rows))
            checks.append(fc)
            if internal in parse_errors:
                fc.error = parse_errors[internal]
                continue
            try:
                got = ev.column(internal)
            except (CalcEvalError, TypeError, KeyError) as e:
                fc.error = f"evaluation failed: {e}"
                continue
            for i, tv in enumerate(got):
                r = table.first_row + i
                xv = wb.get(table.sheet, r, col.col)
                if values_agree(xv, tv):
                    fc.matched += 1
                elif len(fc.mismatches) < keep:
                    fc.mismatches.append((r, xv, tv))
    return checks
