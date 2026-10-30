"""Write and read Tableau workbook XML (.twb) and packaged workbooks (.twbx).

Each detected table becomes a federated datasource over a CSV (textscan)
connection. Data columns become physical fields, translated calc columns
become calculated fields, and referenced input cells become parameters.
"""
from __future__ import annotations

import csv
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Dict, List

from ..formula.values import XlError, number_to_text
from ..tables import Table
from .calc import TLit, bracket, parse_calc, render
from .translate import Parameter, Translation

TABLEAU_VERSION = "18.1"
_TTYPE_TO_DATATYPE = {"number": "real", "string": "string", "boolean": "boolean"}


def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_") or "table"


def _role(datatype: str):
    if datatype in ("real", "integer"):
        return "measure", "quantitative"
    return "dimension", "nominal"


@dataclass
class SourceSpec:
    table: Table
    translations: List[Translation]

    @property
    def caption(self) -> str:
        return self.table.name

    @property
    def slug(self) -> str:
        return _slug(self.table.name)

    @property
    def ds_name(self) -> str:
        return f"federated.{self.slug}"

    @property
    def csv_name(self) -> str:
        return f"{self.slug}.csv"


def write_csv(wb, spec: SourceSpec) -> str:
    """The table's data columns as CSV text (calc columns are recomputed by Tableau)."""
    t = spec.table
    cols = [c for c in t.columns if c.kind == "data"]
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow([c.name for c in cols])
    for r in range(t.first_row, t.last_row + 1):
        out = []
        for c in cols:
            v = wb.get(t.sheet, r, c.col)
            if v is None or isinstance(v, XlError):
                out.append("")
            elif isinstance(v, bool):
                out.append("true" if v else "false")
            elif isinstance(v, float):
                out.append(repr(v) if c.dtype == "real" else number_to_text(v))
            else:
                out.append(str(v))
        w.writerow(out)
    return buf.getvalue()


def _param_literal(p: Parameter) -> str:
    return render(TLit(p.value))


def build_twb(specs: List[SourceSpec], params: List[Parameter]) -> ET.ElementTree:
    root = ET.Element("workbook", {"source-build": "xl2tableau", "version": TABLEAU_VERSION,
                                   "xmlns:user": "http://www.tableausoftware.com/xml/user"})
    ET.SubElement(root, "preferences")
    dss = ET.SubElement(root, "datasources")

    param_internal: Dict[str, str] = {}
    if params:
        pds = ET.SubElement(dss, "datasource", {"hasconnection": "false", "inline": "true",
                                                "name": "Parameters", "version": TABLEAU_VERSION})
        for i, p in enumerate(params, 1):
            internal = f"Parameter {i}"
            param_internal[p.name] = internal
            role, typ = _role(p.dtype)
            col = ET.SubElement(pds, "column", {
                "caption": p.name, "datatype": p.dtype, "name": bracket(internal),
                "param-domain-type": "any", "role": role, "type": typ,
                "value": _param_literal(p)})
            ET.SubElement(col, "calculation", {"class": "tableau", "formula": _param_literal(p)})

    for spec in specs:
        t = spec.table
        ok = [tr for tr in spec.translations if tr.ok]
        calc_internal = {tr.column.name: f"Calculation_{t.first_col}{tr.column.col:03d}{i:03d}"
                         for i, tr in enumerate(ok, 1)}

        def field_name(name: str) -> str:
            return bracket(calc_internal.get(name, name))

        def param_name(name: str) -> str:
            return "[Parameters]." + bracket(param_internal[name])

        ds = ET.SubElement(dss, "datasource", {"caption": spec.caption, "inline": "true",
                                               "name": spec.ds_name, "version": TABLEAU_VERSION})
        conn = ET.SubElement(ds, "connection", {"class": "federated"})
        named = ET.SubElement(ET.SubElement(conn, "named-connections"), "named-connection",
                              {"caption": spec.csv_name, "name": f"textscan.{spec.slug}"})
        ET.SubElement(named, "connection", {"class": "textscan", "directory": "Data",
                                            "filename": spec.csv_name, "password": "", "server": ""})
        rel = ET.SubElement(conn, "relation", {"connection": f"textscan.{spec.slug}",
                                               "name": spec.csv_name,
                                               "table": f"[{spec.csv_name.replace('.', '#')}]",
                                               "type": "table"})
        rcols = ET.SubElement(rel, "columns", {"character-set": "UTF-8", "header": "yes",
                                               "locale": "en_US", "separator": ","})
        data_cols = [c for c in t.columns if c.kind == "data"]
        for i, c in enumerate(data_cols):
            ET.SubElement(rcols, "column", {"datatype": c.dtype, "name": c.name, "ordinal": str(i)})
        for c in data_cols:
            role, typ = _role(c.dtype)
            ET.SubElement(ds, "column", {"datatype": c.dtype, "name": bracket(c.name),
                                         "role": role, "type": typ})
        for tr in ok:
            dt = _TTYPE_TO_DATATYPE[tr.ttype]
            role, typ = _role(dt)
            col = ET.SubElement(ds, "column", {"caption": tr.column.name, "datatype": dt,
                                               "name": bracket(calc_internal[tr.column.name]),
                                               "role": role, "type": typ})
            ET.SubElement(col, "calculation", {
                "class": "tableau", "formula": render(tr.tnode, field_name, param_name)})

    worksheets = ET.SubElement(root, "worksheets")
    for spec in specs:
        _add_worksheet(worksheets, spec)
    return ET.ElementTree(root)


def _add_worksheet(parent, spec: SourceSpec) -> None:
    """A starter sheet: first text column on rows, record count as text."""
    t = spec.table
    ws = ET.SubElement(parent, "worksheet", {"name": spec.caption[:100]})
    table = ET.SubElement(ws, "table")
    view = ET.SubElement(table, "view")
    ET.SubElement(ET.SubElement(view, "datasources"), "datasource",
                  {"caption": spec.caption, "name": spec.ds_name})
    deps = ET.SubElement(view, "datasource-dependencies", {"datasource": spec.ds_name})
    dim = next((c for c in t.columns if c.kind == "data" and c.dtype == "string"), None)
    if dim is not None:
        ET.SubElement(deps, "column", {"datatype": "string", "name": bracket(dim.name),
                                       "role": "dimension", "type": "nominal"})
        ET.SubElement(deps, "column-instance", {"column": bracket(dim.name), "derivation": "None",
                                                "name": bracket(f"none:{dim.name}:nk"),
                                                "pivot": "key", "type": "nominal"})
    ET.SubElement(view, "aggregation", {"value": "true"})
    ET.SubElement(table, "style")
    panes = ET.SubElement(table, "panes")
    pane = ET.SubElement(panes, "pane")
    ET.SubElement(ET.SubElement(pane, "view"), "breakdown", {"value": "auto"})
    ET.SubElement(pane, "mark", {"class": "Automatic"})
    rows = ET.SubElement(table, "rows")
    if dim is not None:
        rows.text = f"[{spec.ds_name}].{bracket(f'none:{dim.name}:nk')}"
    ET.SubElement(table, "cols")


def to_xml_bytes(tree: ET.ElementTree) -> bytes:
    ET.indent(tree, space="  ")
    return b"<?xml version='1.0' encoding='utf-8' ?>\n" + ET.tostring(tree.getroot(), encoding="utf-8")


def write_twbx(path: str, twb_bytes: bytes, csvs: Dict[str, str], twb_name: str = "workbook.twb") -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(twb_name, twb_bytes)
        for name, text in csvs.items():
            z.writestr(f"Data/{name}", text)


# ------------------------------------------------------------------ reading

@dataclass
class ParsedDatasource:
    name: str
    caption: str
    csv_file: str
    relation_columns: List[tuple]  # (name, datatype)
    calcs: Dict[str, str] = field(default_factory=dict)  # internal name -> formula
    captions: Dict[str, str] = field(default_factory=dict)  # internal name -> caption
    datatypes: Dict[str, str] = field(default_factory=dict)


@dataclass
class ParsedWorkbook:
    datasources: List[ParsedDatasource]
    params: Dict[str, object]  # internal parameter name -> value
    param_captions: Dict[str, str]


def _unbracket(s: str) -> str:
    return s[1:-1].replace("]]", "]") if s.startswith("[") else s


def parse_twb(xml_bytes: bytes) -> ParsedWorkbook:
    root = ET.fromstring(xml_bytes)
    sources, params, pcap = [], {}, {}
    for ds in root.find("datasources"):
        if ds.get("name") == "Parameters":
            for col in ds.findall("column"):
                internal = _unbracket(col.get("name"))
                lit = parse_calc(col.find("calculation").get("formula"))
                params[internal] = lit.value if isinstance(lit, TLit) else None
                pcap[internal] = col.get("caption", internal)
            continue
        named = ds.find("connection/named-connections/named-connection/connection")
        rel = ds.find("connection/relation")
        pds = ParsedDatasource(ds.get("name"), ds.get("caption", ds.get("name")),
                               named.get("filename"),
                               [(c.get("name"), c.get("datatype")) for c in rel.find("columns")])
        for col in ds.findall("column"):
            internal = _unbracket(col.get("name"))
            pds.captions[internal] = col.get("caption", internal)
            pds.datatypes[internal] = col.get("datatype")
            calc = col.find("calculation")
            if calc is not None:
                pds.calcs[internal] = calc.get("formula")
        sources.append(pds)
    return ParsedWorkbook(sources, params, pcap)


def read_csv_typed(text: str, columns: List[tuple]) -> List[Dict[str, object]]:
    """Read CSV the way a typed textscan connection would: blanks become NULL."""
    reader = csv.reader(io.StringIO(text))
    header = next(reader)
    types = dict(columns)
    rows = []
    for rec in reader:
        row = {}
        for name, raw in zip(header, rec):
            dt = types.get(name, "string")
            if raw == "":
                row[name] = None
            elif dt in ("real", "integer"):
                row[name] = float(raw)
            elif dt == "boolean":
                row[name] = raw.lower() == "true"
            else:
                row[name] = raw
        rows.append(row)
    return rows


def load_twbx(path: str):
    """Return (ParsedWorkbook, {csv filename: text}) from a .twbx file."""
    with zipfile.ZipFile(path) as z:
        twb_name = next(n for n in z.namelist() if n.endswith(".twb"))
        parsed = parse_twb(z.read(twb_name))
        csvs = {n.split("/", 1)[1]: z.read(n).decode("utf-8")
                for n in z.namelist() if n.startswith("Data/")}
    return parsed, csvs
