import re
import zipfile
import xml.etree.ElementTree as ET

from xl2tableau.cli import main
from xl2tableau.convert import convert_file
from xl2tableau.engine import Workbook
from xl2tableau.sample import build_sales_workbook


def _convert(tmp_path, n=120, override_row=None):
    x, _ = build_sales_workbook(n, seed=5, override_row=override_row)
    src = tmp_path / "model.xlsx"
    x.save(src)
    return convert_file(str(src), str(tmp_path / "out"))


def test_sample_model_end_to_end(tmp_path):
    res = _convert(tmp_path)
    assert [t.name for t in res.tables] == ["Sales", "Products"]
    status = {tr.column.name: tr.ok for tr in res.translations["Sales"]}
    assert status["List Price"] is False and status["Running Total"] is False
    assert sum(status.values()) == 10
    assert res.rows_checked == 10 * 120
    assert res.rows_matched == res.rows_checked
    assert {p.name for p in res.params} == {"TaxRate", "High Threshold"}


def test_override_is_flagged_and_shows_as_mismatch(tmp_path):
    res = _convert(tmp_path, override_row=9)
    tax = next(c for c in res.tables[0].columns if c.name == "Tax")
    assert tax.inconsistent_rows == [9]
    bad = {c.field: c.mismatches for c in res.checks if not c.ok}
    assert set(bad) == {"Tax", "Margin"}  # Margin depends on Tax in Excel
    assert bad["Tax"][0][0] == 9


def test_twbx_is_well_formed_and_references_resolve(tmp_path):
    res = _convert(tmp_path)
    with zipfile.ZipFile(res.outputs["twbx"]) as z:
        names = z.namelist()
        assert "model.twb" in names and "Data/sales.csv" in names and "Data/products.csv" in names
        root = ET.fromstring(z.read("model.twb"))
    params = {c.get("name") for c in root.find("datasources/datasource[@name='Parameters']")}
    for ds in root.findall("datasources/datasource"):
        if ds.get("name") == "Parameters":
            continue
        declared = {c.get("name") for c in ds.findall("column")}
        for calc in ds.findall("column/calculation"):
            formula = calc.get("formula")
            for qualified in re.findall(r"\[Parameters\]\.(\[[^\]]+\])", formula):
                assert qualified in params
            stripped = re.sub(r"\[Parameters\]\.\[[^\]]+\]", "", formula)
            for ref in re.findall(r"\[[^\]]+\]", stripped):
                assert ref in declared, ref


def test_audit_workbook_formulas(tmp_path):
    res = _convert(tmp_path, override_row=9)
    audit = Workbook.load(res.outputs["audit"])
    audit.recalc()
    summary = {audit.get("Summary", r, 1): audit.get("Summary", r, 2) for r in range(2, 12)}
    assert summary["Calculated columns"] == 12
    assert summary["Translated and verified"] == 8
    assert summary["Translated, rows disagree"] == 2
    assert summary["Needs manual work"] == 2
    assert summary["Rows checked"] == res.rows_checked
    assert summary["Rows matching"] == res.rows_matched
    assert summary["Parameters"] == 2
    assert summary["Inconsistent formula cells"] == 2  # override + Running Total's first row


def test_cli_convert_and_eval(tmp_path, capsys):
    x, _ = build_sales_workbook(20, seed=1)
    src = tmp_path / "m.xlsx"
    x.save(src)
    assert main(["convert", str(src), "-o", str(tmp_path / "o")]) == 0
    out = capsys.readouterr().out
    assert "translated 10/12 calculated columns" in out
    assert main(["eval", str(src), "--sheet", "Assumptions"]) == 0


def test_cli_strict_fails_on_mismatch(tmp_path):
    x, _ = build_sales_workbook(20, seed=1, override_row=4)
    src = tmp_path / "m.xlsx"
    x.save(src)
    assert main(["convert", str(src), "-o", str(tmp_path / "o")]) == 0
    assert main(["convert", str(src), "-o", str(tmp_path / "o"), "--strict"]) == 1
