import math
import random

import pytest

from xl2tableau.engine import Workbook
from xl2tableau.formula.parser import parse
from xl2tableau.formula.values import XlError
from xl2tableau.sample import build_sales_workbook, expected_values


def ev(formula, cells=None):
    wb = Workbook()
    for addr, v in (cells or {}).items():
        col = ord(addr[0]) - 64
        wb.set("S", int(addr[1:]), col, v)
    wb.set("S", 100, 26, formula)
    wb.recalc()
    return wb.get("S", 100, 26)


@pytest.mark.parametrize("formula,expected", [
    ("=1+2*3", 7.0),
    ("=-2^2", 4.0),
    ("=2^3^2", 64.0),
    ("=10/4", 2.5),
    ("=50%*2", 1.0),
    ('="a"&1&TRUE', "a1TRUE"),
    ('="5"+1', 6.0),
    ("=TRUE+1", 2.0),
    ('="abc"="ABC"', True),
    ('=1<"a"', True),           # numbers sort before text
    ('="a"<TRUE', True),        # text sorts before logicals
    ("=ROUND(2.5,0)", 3.0),     # half away from zero, unlike Python round()
    ("=ROUND(-2.5,0)", -3.0),
    ("=ROUND(1.005,2)", 1.01),
    ("=ROUNDUP(1.21,1)", 1.3),
    ("=ROUNDDOWN(-1.29,1)", -1.2),
    ("=MOD(-7,3)", 2.0),        # sign of divisor
    ("=MOD(7,-3)", -2.0),
    ("=INT(-1.5)", -2.0),
    ('=MID("spreadsheet",3,4)', "read"),
    ('=TRIM("  a   b ")', "a b"),
    ('=LEFT("abc")', "a"),
    ('=RIGHT("abc",2)', "bc"),
    ('=LEN(12.5)', 4.0),
    ('=CONCATENATE("x",2,"y")', "x2y"),
    ("=AND(TRUE,1)", True),
    ("=OR(FALSE,0)", False),
    ("=NOT(0)", True),
    ("=SQRT(16)", 4.0),
    ("=POWER(2,10)", 1024.0),
])
def test_scalar_formulas(formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula,code", [
    ("=1/0", "#DIV/0!"),
    ('="x"+1', "#VALUE!"),
    ("=SQRT(-1)", "#NUM!"),
    ("=NOSUCHFN(1)", "#NAME?"),
    ("=AVERAGE(A1:A3)", "#DIV/0!"),
    ("=MOD(1,0)", "#DIV/0!"),
    ("=0^-1", "#DIV/0!"),
])
def test_errors(formula, code):
    assert ev(formula) == XlError(code)


def test_errors_propagate_and_iferror_catches():
    assert ev("=1+1/0") == XlError("#DIV/0!")
    assert ev("=IFERROR(1/0,-1)") == -1.0
    assert ev("=SUM(A1:A2)", {"A1": 1, "A2": "=1/0"}) == XlError("#DIV/0!")


def test_if_is_lazy():
    assert ev("=IF(TRUE,1,1/0)") == 1.0
    assert ev("=IF(FALSE,1)") is False


def test_blank_semantics():
    assert ev("=A1+1") == 1.0
    assert ev('=A1&"x"') == "x"
    assert ev("=A1=0") is True
    assert ev('=A1=""') is True
    assert ev("=A1") == 0.0  # a formula pointing at a blank shows 0


def test_aggregates_skip_text_and_blanks_in_ranges():
    cells = {"A1": 1, "A2": "2", "A3": True, "A4": 4}
    assert ev("=SUM(A1:A5)", cells) == 5.0
    assert ev("=COUNT(A1:A5)", cells) == 2.0
    assert ev("=COUNTA(A1:A5)", cells) == 4.0
    assert ev("=AVERAGE(A1:A5)", cells) == 2.5
    assert ev('=SUM(1,"2",TRUE)') == 4.0  # direct arguments are coerced


def test_conditional_functions():
    cells = {"A1": "East", "A2": "west", "A3": "East", "A4": "Eastern",
             "B1": 10, "B2": 20, "B3": 30, "B4": 40}
    assert ev('=SUMIF(A1:A4,"east",B1:B4)', cells) == 40.0
    assert ev('=SUMIF(A1:A4,"East*",B1:B4)', cells) == 80.0
    assert ev('=SUMIF(A1:A4,"<>East",B1:B4)', cells) == 60.0
    assert ev('=COUNTIF(B1:B4,">=20")', cells) == 3.0
    assert ev('=AVERAGEIF(A1:A4,"East",B1:B4)', cells) == 20.0
    assert ev('=SUMIFS(B1:B4,A1:A4,"East",B1:B4,">15")', cells) == 30.0
    assert ev('=COUNTIFS(A1:A4,"?ast",B1:B4,"<100")', cells) == 2.0
    assert ev('=SUMIF(A1:A4,A1,B1:B4)', cells) == 40.0


def test_criterion_numeric_text_matches_numbers():
    cells = {"A1": 5, "A2": "5", "A3": 6}
    assert ev('=COUNTIF(A1:A3,"5")', cells) == 2.0
    assert ev("=COUNTIF(A1:A3,5)", cells) == 2.0


def test_lookups():
    cells = {"A1": "apple", "A2": "Banana", "A3": "cherry",
             "B1": 1, "B2": 2, "B3": 3, "C1": 10, "C2": 20, "C3": 30}
    assert ev('=VLOOKUP("banana",A1:C3,3,FALSE)', cells) == 20.0
    assert ev('=VLOOKUP("kiwi",A1:C3,2,FALSE)', cells) == XlError("#N/A")
    assert ev("=VLOOKUP(2.5,B1:C3,2,TRUE)", cells) == 20.0
    assert ev("=VLOOKUP(0,B1:C3,2)", cells) == XlError("#N/A")
    assert ev('=MATCH("cherry",A1:A3,0)', cells) == 3.0
    assert ev("=INDEX(C1:C3,2)", cells) == 20.0
    assert ev('=INDEX(A1:C3,MATCH("cherry",A1:A3,0),2)', cells) == 3.0
    assert ev('=VLOOKUP("apple",A1:C3,4,FALSE)', cells) == XlError("#REF!")


def test_implicit_intersection():
    wb = Workbook()
    for r in range(1, 4):
        wb.set("S", r, 1, r * 10)
    wb.set("S", 2, 2, "=A1:A3*2")  # same-row cell of the range
    wb.set("S", 5, 2, "=A1:A3*2")  # no intersection
    wb.recalc()
    assert wb.get("S", 2, 2) == 40.0
    assert wb.get("S", 5, 2) == XlError("#VALUE!")


def test_cross_sheet_and_defined_names():
    wb = Workbook()
    wb.set("Inputs", 1, 1, 0.5)
    wb.set("Calc", 1, 1, "=Inputs!A1*Rate*4")
    wb.defined_names["RATE"] = parse("Inputs!$A$1")
    wb.recalc()
    assert wb.get("Calc", 1, 1) == 1.0


def test_topological_order_with_forward_references():
    wb = Workbook()
    wb.set("S", 1, 1, "=SUM(B1:B3)")  # depends on formulas defined "later"
    wb.set("S", 1, 2, "=B2*2")
    wb.set("S", 2, 2, "=B3+1")
    wb.set("S", 3, 2, 5)
    wb.recalc()
    assert wb.get("S", 1, 1) == 12.0 + 6.0 + 5.0


def test_long_dependency_chain_has_no_recursion_limit():
    wb = Workbook()
    wb.set("S", 1, 1, 1)
    n = 20000
    for r in range(2, n + 1):
        wb.set("S", r, 1, f"=A{r - 1}+1")
    wb.recalc()
    assert wb.get("S", n, 1) == float(n)


def test_cycles_are_reported_not_hung():
    wb = Workbook()
    wb.set("S", 1, 1, "=B1+1")
    wb.set("S", 1, 2, "=A1+1")
    wb.set("S", 1, 3, "=A1*2")  # downstream of the cycle
    wb.set("S", 1, 4, "=5")
    stats = wb.recalc()
    assert stats.cycle_cells == 3
    assert wb.get("S", 1, 3) == XlError("#CIRC!")
    assert wb.get("S", 1, 4) == 5.0


def test_whole_column_range_does_not_depend_on_itself():
    wb = Workbook()
    for r in range(1, 4):
        wb.set("S", r, 1, r)
    wb.set("S", 1, 2, "=SUM(A:A)")
    wb.recalc()
    assert wb.get("S", 1, 2) == 6.0


def _random_conditional_workbook(seed, n=300):
    rnd = random.Random(seed)
    wb = Workbook()
    keys = ["a", "A", "b", "5", 5, 7.5, True, None, "c d"]
    for r in range(1, n + 1):
        k = rnd.choice(keys)
        if k is not None:
            wb.set("S", r, 1, k)
        wb.set("S", r, 2, rnd.choice(["x", "y", 1, 2]))
        wb.set("S", r, 3, rnd.choice([1.5, -2, 3, "txt", None]))
        wb.set("S", r, 4, f"=SUMIF($A$1:$A${n},A{r},$C$1:$C${n})")
        wb.set("S", r, 5, f"=COUNTIFS($A$1:$A${n},A{r},$B$1:$B${n},B{r})")
        wb.set("S", r, 6, f"=AVERAGEIF($A$1:$A${n},A{r},$C$1:$C${n})")
        wb.set("S", r, 7, f"=VLOOKUP(A{r},$A$1:$C${n},3,FALSE)")
        wb.set("S", r, 8, f"=MATCH(B{r},$B$1:$B${n},0)")
    return wb


@pytest.mark.parametrize("seed", [1, 2, 3])
def test_hash_index_matches_linear_scan(seed):
    wb1 = _random_conditional_workbook(seed)
    wb2 = _random_conditional_workbook(seed)
    wb1.recalc(use_index=True)
    wb2.recalc(use_index=False)
    assert wb1.values == wb2.values


def test_sample_model_matches_independent_ground_truth():
    wb, rows = build_sales_workbook(400, seed=11, as_engine=True)
    wb.recalc()
    exp = expected_values(rows)
    from xl2tableau.sample import HEADERS

    for i, e in enumerate(exp):
        for name, want in e.items():
            got = wb.get("Sales", i + 2, HEADERS.index(name) + 1)
            if want is None:
                assert isinstance(got, XlError)
            elif isinstance(want, float):
                assert math.isclose(got, want, rel_tol=1e-9, abs_tol=1e-9), (i, name)
            else:
                assert got == want, (i, name)


def test_load_from_openpyxl_file(tmp_path):
    x, rows = build_sales_workbook(30, seed=3)
    p = tmp_path / "m.xlsx"
    x.save(p)
    wb = Workbook.load(str(p))
    stats = wb.recalc()
    assert stats.formulas == 30 * 12 and stats.parse_errors == 0 and stats.cycle_cells == 0
    assert "TAXRATE" in wb.defined_names
    exp = expected_values(rows)
    assert math.isclose(wb.get("Sales", 2, 7), exp[0]["Tax"])
