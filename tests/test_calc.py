import pytest

from xl2tableau.tableau.calc import (
    CalcEvalError, CalcEvaluator, TBin, TField, TFunc, TIf, TLit, TLod, TParam, parse_calc, render,
)


def test_render_parse_round_trip():
    node = TIf(((TBin(">=", TField("Sales]x"), TParam("Goal")), TLit("hi \"there\"")),
                (TBin("==", TField("A"), TLit(-3.0)), TLit("neg"))), TLit("lo"))
    text = render(node)
    # -3 renders as "(-3)" and parses back as a negation, so compare by evaluation.
    rows = [{"Sales]x": 5.0, "A": -3.0}, {"Sales]x": 1.0, "A": -3.0}, {"Sales]x": 1.0, "A": 0.0}]
    ev1 = CalcEvaluator(rows, {"c": node}, {"Goal": 2.0})
    ev2 = CalcEvaluator(rows, {"c": parse_calc(text)}, {"Goal": 2.0})
    assert ev1.column("c") == ev2.column("c") == ["hi \"there\"", "neg", "lo"]


def test_lod_render_and_parse():
    node = TLod(("Region", "Product"), TFunc("SUM", (TField("Units"),)))
    text = render(node)
    assert text == "{FIXED [Region], [Product] : SUM([Units])}"
    assert parse_calc(text) == node
    assert parse_calc("{FIXED : AVG([x])}") == TLod((), TFunc("AVG", (TField("x"),)))


def test_null_semantics():
    rows = [{"a": None, "b": 2.0}, {"a": 4.0, "b": 0.0}]
    ev = CalcEvaluator(rows, {
        "sum": parse_calc("[a] + [b]"),
        "div": parse_calc("[a] / [b]"),
        "zn": parse_calc("ZN([a]) + [b]"),
        "ifn": parse_calc("IFNULL([a] / [b], -1)"),
        "cond": parse_calc('IF [a] > 1 THEN "big" ELSE "other" END'),
        "noelse": parse_calc('IF [a] > 1 THEN "big" END'),
    })
    assert ev.column("sum") == [None, 4.0]
    assert ev.column("div") == [None, None]          # divide by zero is NULL
    assert ev.column("zn") == [2.0, 4.0]
    assert ev.column("ifn") == [-1.0, -1.0]
    assert ev.column("cond") == ["other", "big"]     # NULL condition falls to ELSE
    assert ev.column("noelse") == [None, "big"]


def test_three_valued_logic():
    ev = CalcEvaluator([{"n": None}], {
        "and_f": parse_calc("FALSE AND [n] > 1"),
        "or_t": parse_calc("TRUE OR [n] > 1"),
        "and_n": parse_calc("TRUE AND [n] > 1"),
    })
    assert ev.column("and_f") == [False]
    assert ev.column("or_t") == [True]
    assert ev.column("and_n") == [None]


def test_fixed_lod_and_calc_dependencies():
    rows = [{"r": "E", "v": 1.0}, {"r": "W", "v": 2.0}, {"r": "E", "v": None}, {"r": "E", "v": 4.0}]
    ev = CalcEvaluator(rows, {
        "dbl": parse_calc("[v] * 2"),
        "by_r": parse_calc("{FIXED [r] : SUM([dbl])}"),
        "total": parse_calc("{FIXED : SUM([v])}"),
        "cnt": parse_calc("{FIXED [r] : COUNT([v])}"),
        "share": parse_calc("[dbl] / [by_r]"),
    })
    assert ev.column("by_r") == [10.0, 4.0, 10.0, 10.0]
    assert ev.column("total") == [7.0] * 4
    assert ev.column("cnt") == [2.0, 1.0, 2.0, 2.0]
    assert ev.column("share") == [0.2, 1.0, None, 0.8]


def test_tableau_specific_functions():
    ev = CalcEvaluator([{"s": "  a  b ", "x": -7.0}], {
        "trim": parse_calc("TRIM([s])"),
        "pct": parse_calc("[x] % 3"),
        "str": parse_calc("STR(1.5) + STR(2)"),
        "rnd": parse_calc("ROUND(2.5, 0)"),
    })
    assert ev.column("trim") == ["a  b"]      # ends only, unlike Excel
    assert ev.column("pct") == [-1.0]         # sign of dividend, unlike Excel MOD
    assert ev.column("str") == ["1.52"]
    assert ev.column("rnd") == [3.0]


def test_errors():
    with pytest.raises(ValueError):
        parse_calc("IF [a] THEN 1")
    with pytest.raises(ValueError):
        parse_calc("{FIXED [a] : [b]}")
    ev = CalcEvaluator([{"a": 1.0}], {"x": parse_calc("[missing] + 1"), "c1": parse_calc("[c2]"),
                                      "c2": parse_calc("[c1]")})
    with pytest.raises(CalcEvalError):
        ev.column("x")
    with pytest.raises(CalcEvalError):
        ev.column("c1")
