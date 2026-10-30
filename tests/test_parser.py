import pytest

from xl2tableau.formula.ast import MAX_ROW, Binary, Func, Name, Num, Percent, Ref, Str, Unary
from xl2tableau.formula.parser import FormulaParseError, parse


def test_precedence_mul_over_add():
    assert parse("=1+2*3") == Binary("+", Num(1.0), Binary("*", Num(2.0), Num(3.0)))


def test_negation_binds_tighter_than_power():
    # Excel: =-2^2 is 4
    assert parse("=-2^2") == Binary("^", Unary("-", Num(2.0)), Num(2.0))


def test_power_is_left_associative():
    assert parse("=2^3^2") == Binary("^", Binary("^", Num(2.0), Num(3.0)), Num(2.0))


def test_concat_below_addition_and_comparison_lowest():
    node = parse('=A1&"x"=B1+1')
    assert node.op == "="
    assert node.left.op == "&"
    assert node.right.op == "+"


def test_percent_postfix():
    assert parse("=50%") == Percent(Num(50.0))


def test_cell_and_range_refs():
    r = parse("=$B$3")
    assert isinstance(r, Ref) and r.is_cell and (r.col1, r.row1) == (2, 3) and r.abs_col1 and r.abs_row1
    rng = parse("=A10:C2")  # corners are normalised
    assert (rng.col1, rng.row1, rng.col2, rng.row2) == (1, 2, 3, 10)


def test_whole_column_and_sheet_refs():
    r = parse("=SUM('My Sheet'!C:C)").args[0]
    assert r.sheet == "My Sheet" and r.whole_col and r.row2 == MAX_ROW
    r2 = parse("=Data!A1").sheet
    assert r2 == "Data"


def test_quoted_sheet_with_apostrophe():
    r = parse("='Bob''s'!A1")
    assert r.sheet == "Bob's"


def test_function_names_that_look_like_cells():
    # LOG10 looks like a cell reference but is followed by '('
    node = parse("=LOG10(A1)")
    assert isinstance(node, Func) and node.name == "LOG10"


def test_names_and_booleans():
    assert parse("=TaxRate*2").left == Name("TaxRate")
    assert parse("=TRUE").value is True


def test_string_escape():
    assert parse('="say ""hi"""') == Str('say "hi"')


def test_empty_function_and_nested_args():
    node = parse("=IF(AND(A1>0,B1<5),SUM(A1:A3),NOW())")
    assert node.name == "IF" and node.args[2] == Func("NOW", ())


@pytest.mark.parametrize("bad", ["=1+", "=SUM(1,2", '="abc', "=1 2", "=#BOGUS", "=)"])
def test_parse_errors(bad):
    with pytest.raises(FormulaParseError):
        parse(bad)


def test_a1_round_trip():
    for text in ["$A$1", "B2:C9", "Sheet2!$D4", "'My Sheet'!A:B"]:
        assert parse("=" + text).a1() == text
