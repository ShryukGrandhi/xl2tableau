import pytest

from xl2tableau.convert import convert_workbook
from xl2tableau.engine import Workbook
from xl2tableau.formula.parser import parse
from xl2tableau.tableau.calc import render

DATA = [  # Name, Qty, Price, Disc
    ("Alpha", 3, 2.5, 0.1),
    ("beta", -7, 4.0, None),
    ("Alpha", 0, 1.25, 0.0),
    ("Gamma", 12, 3.0, None),
]


def build(formula_template, data=DATA, extra=None):
    """One table with columns A..D of data and E as a calc column."""
    wb = Workbook()
    for j, h in enumerate(["Name", "Qty", "Price", "Disc", "Calc"], 1):
        wb.set("T", 1, j, h)
    last = len(data) + 1
    for i, rec in enumerate(data):
        r = i + 2
        for j, v in enumerate(rec, 1):
            wb.set("T", r, j, v)
        wb.set("T", r, 5, formula_template.format(r=r, last=last, prev=r - 1))
    wb.set("Inputs", 1, 1, "Rate")
    wb.set("Inputs", 1, 2, 0.2)
    wb.set("Inputs", 2, 1, "Label")
    wb.set("Inputs", 2, 2, "lbl")
    for (s, r, c), v in (extra or {}).items():
        wb.set(s, r, c, v)
    return wb


def run(tmp_path, template, **kw):
    wb = build(template, **kw)
    res = convert_workbook(wb, str(tmp_path))
    tr = res.translations["T"][0]
    chk = next((c for c in res.checks if c.field == "Calc"), None)
    return tr, chk, res


@pytest.mark.parametrize("template,expected", [
    ("=B{r}*C{r}", "([Qty] * [Price])"),
    ("=B{r}*C{r}*(1-D{r})", "(([Qty] * [Price]) * (1 - ZN([Disc])))"),
    ("=B{r}*Inputs!$B$1", "([Qty] * [Parameters].[Rate])"),
    ('=A{r}&"#"&B{r}', '(([Name] + "#") + STR([Qty]))'),
    ("=B{r}^2", "POWER([Qty], 2)"),
    ("=C{r}/SUM($C$2:$C${last})", "([Price] / {FIXED : SUM([Price])})"),
    ("=SUMIF($A$2:$A${last},A{r},$B$2:$B${last})", "{FIXED [Name] : SUM([Qty])}"),
    ("=COUNTIF($A$2:$A${last},A{r})", "{FIXED [Name] : COUNT([Name])}"),
    ("=AVERAGE(B:B)", "{FIXED : AVG([Qty])}"),
    ('=IF(B{r}>5,"big",IF(B{r}>0,"small","none"))',
     'IF ([Qty] > 5) THEN "big" ELSEIF ([Qty] > 0) THEN "small" ELSE "none" END'),
    ("=IFERROR(C{r}/B{r},0)", "IFNULL(([Price] / [Qty]), 0)"),
    ("=AND(B{r}>0,C{r}<4)", "(([Qty] > 0) AND ([Price] < 4))"),
    ("=ISBLANK(D{r})", "ISNULL([Disc])"),
    ("=MAX(B{r},0,C{r})", "MAX(MAX([Qty], 0), [Price])"),
])
def test_translations_render_and_agree_with_excel(tmp_path, template, expected):
    tr, chk, _ = run(tmp_path, template)
    assert tr.ok, tr.reason
    assert render(tr.tnode) == expected
    assert chk.error is None
    assert chk.matched == chk.rows == len(DATA), chk.mismatches


def test_mod_matches_excel_for_negative_numbers(tmp_path):
    # Qty -7: Excel MOD(-7,3) = 2; Tableau's % would give -1.
    tr, chk, _ = run(tmp_path, "=MOD(B{r},3)")
    assert tr.ok and chk.matched == chk.rows


def test_int_uses_floor(tmp_path):
    tr, chk, _ = run(tmp_path, "=INT(B{r}/2)")
    assert "FLOOR" in render(tr.tnode) and chk.matched == chk.rows


def test_constant_criterion_sumif(tmp_path):
    tr, chk, _ = run(tmp_path, '=SUMIF($A$2:$A$5,"Alpha",$B$2:$B$5)')
    assert render(tr.tnode) == '{FIXED : SUM(IF ([Name] == "Alpha") THEN [Qty] END)}'
    assert chk.matched == chk.rows


def test_defined_name_becomes_named_parameter(tmp_path):
    wb = build("=B{r}*MyRate")
    wb.defined_names["MYRATE"] = parse("Inputs!$B$1")
    res = convert_workbook(wb, str(tmp_path))
    tr = res.translations["T"][0]
    assert render(tr.tnode) == "([Qty] * [Parameters].[MyRate])"
    assert [p.name for p in res.params] == ["MyRate"]


def test_grand_total_cell_is_inlined_as_lod(tmp_path):
    tr, chk, _ = run(tmp_path, "=B{r}/Inputs!$B$5", extra={("Inputs", 5, 2): "=SUM(T!B2:B5)"})
    assert render(tr.tnode) == "([Qty] / {FIXED : SUM([Qty])})"
    assert chk.matched == chk.rows


def test_case_sensitivity_difference_is_detected(tmp_path):
    # Excel's = ignores case; Tableau's == does not. The check must catch this.
    tr, chk, _ = run(tmp_path, '=A{r}="alpha"')
    assert tr.ok
    assert chk.matched == 2 and chk.rows == 4
    assert {m[0] for m in chk.mismatches} == {2, 4}


def test_trim_difference_is_detected(tmp_path):
    data = [("a  b", 1, 1.0, None), ("c", 2, 1.0, None)]
    tr, chk, _ = run(tmp_path, "=TRIM(A{r})", data=data)
    assert tr.ok and chk.matched == 1


@pytest.mark.parametrize("template,reason", [
    ("=E{prev}+B{r}", "another row"),
    ('=VLOOKUP(A{r},$A$2:$C$5,2,FALSE)', "relationship or join"),
    ("=B{r}*Inputs!B1", "changes per row"),
    ('=IF(B{r}>0,"pos",0)', "different types"),
    ('=A{r}+1', "text used in arithmetic"),
    ("=SUM($B$2:$B$4)", "not exactly one table column"),
    ("=SUMIF($A$2:$A$5,B{r},$B$2:$B$5)", "cross-column"),
    ("=NOW()", "no translation"),
    ('=B{r}=A{r}', "between text and a number"),
])
def test_untranslatable_with_reason(tmp_path, template, reason):
    tr, chk, _ = run(tmp_path, template)
    assert not tr.ok
    assert reason in tr.reason


def test_dependency_on_untranslatable_column(tmp_path):
    wb = build('=VLOOKUP(A{r},$A$2:$C$5,2,FALSE)')
    wb.set("T", 1, 6, "Next")
    for r in range(2, 6):
        wb.set("T", r, 6, f"=E{r}*2")
    res = convert_workbook(wb, str(tmp_path))
    nxt = res.translations["T"][1]
    assert not nxt.ok and "untranslatable column 'Calc'" in nxt.reason
