"""End-to-end pipeline: .xlsx -> recalculated model -> .twbx + audit workbook."""
from __future__ import annotations

import os
import time
from dataclasses import dataclass, field
from typing import Dict, List

from .audit import write_audit
from .engine import RecalcStats, Workbook
from .tables import Table, detect_tables
from .tableau.translate import Parameter, ParameterRegistry, TableTranslator, Translation
from .tableau.twb import SourceSpec, build_twb, to_xml_bytes, write_csv, write_twbx
from .validate import FieldCheck, validate_twbx


@dataclass
class ConversionResult:
    workbook: Workbook
    tables: List[Table]
    translations: Dict[str, List[Translation]]
    params: List[Parameter]
    checks: List[FieldCheck]
    recalc: RecalcStats
    outputs: Dict[str, str]
    timings: Dict[str, float] = field(default_factory=dict)

    @property
    def calc_columns(self) -> int:
        return sum(len(v) for v in self.translations.values())

    @property
    def translated(self) -> int:
        return sum(1 for v in self.translations.values() for t in v if t.ok)

    @property
    def rows_checked(self) -> int:
        return sum(c.rows for c in self.checks)

    @property
    def rows_matched(self) -> int:
        return sum(c.matched for c in self.checks)


def convert_workbook(wb: Workbook, out_dir: str, stem: str = "workbook") -> ConversionResult:
    os.makedirs(out_dir, exist_ok=True)
    timings = {}
    t0 = time.perf_counter()
    stats = wb.recalc()
    t1 = time.perf_counter()
    tables = detect_tables(wb)
    registry = ParameterRegistry(wb)
    translations = {t.name: TableTranslator(wb, t, registry).translate_all() for t in tables}
    t2 = time.perf_counter()
    specs = [SourceSpec(t, translations[t.name]) for t in tables]
    params = list(registry.by_cell.values())
    xml = to_xml_bytes(build_twb(specs, params))
    csvs = {s.csv_name: write_csv(wb, s) for s in specs}
    twbx = os.path.join(out_dir, f"{stem}.twbx")
    twb = os.path.join(out_dir, f"{stem}.twb")
    write_twbx(twbx, xml, csvs, f"{stem}.twb")
    with open(twb, "wb") as f:
        f.write(xml)
    data_dir = os.path.join(out_dir, "Data")
    os.makedirs(data_dir, exist_ok=True)
    for name, text in csvs.items():
        with open(os.path.join(data_dir, name), "w", encoding="utf-8", newline="") as f:
            f.write(text)
    t3 = time.perf_counter()
    checks = validate_twbx(twbx, wb, tables)
    t4 = time.perf_counter()
    timings.update(recalc=t1 - t0, translate=t2 - t1, write=t3 - t2, validate=t4 - t3)
    result = ConversionResult(wb, tables, translations, params, checks, stats,
                              {"twbx": twbx, "twb": twb, "data": data_dir}, timings)
    audit = os.path.join(out_dir, f"{stem}_audit.xlsx")
    write_audit(result, audit)
    result.outputs["audit"] = audit
    return result


def convert_file(path: str, out_dir: str) -> ConversionResult:
    stem = os.path.splitext(os.path.basename(path))[0]
    return convert_workbook(Workbook.load(path), out_dir, stem)
