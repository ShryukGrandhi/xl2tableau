"""Command line interface: ``xl2tableau convert`` and ``xl2tableau eval``."""
from __future__ import annotations

import argparse
import sys

from .engine import Workbook
from .formula.ast import index_to_col


def _cmd_convert(args) -> int:
    from .convert import convert_file

    res = convert_file(args.input, args.out)
    for table in res.tables:
        print(f"table {table.name!r} on sheet {table.sheet!r}: {table.n_rows} rows, "
              f"{len(table.columns)} columns")
        for tr in res.translations[table.name]:
            chk = next((c for c in res.checks if c.table == table.name and c.field == tr.column.name), None)
            if tr.ok:
                state = "translated"
                if chk:
                    tag = "ok" if chk.matched == chk.rows else "DIFFERS"
                    state = f"{tag} {chk.matched}/{chk.rows} rows agree"
                if chk and chk.error:
                    state = chk.error
            else:
                state = f"MANUAL: {tr.reason}"
            print(f"  {tr.column.name:<24} {state}")
            if tr.column.inconsistent_rows:
                print(f"  {'':<24} {len(tr.column.inconsistent_rows)} cells differ from the column formula")
    print(f"translated {res.translated}/{res.calc_columns} calculated columns; "
          f"{res.rows_matched}/{res.rows_checked} checked rows agree with Excel")
    for k, v in res.outputs.items():
        print(f"{k}: {v}")
    if args.strict and res.rows_matched != res.rows_checked:
        return 1
    return 0


def _cmd_eval(args) -> int:
    wb = Workbook.load(args.input)
    stats = wb.recalc()
    print(f"{stats.formulas} formulas, {stats.cycle_cells} in cycles, {stats.parse_errors} parse errors")
    for (sheet, r, c), v in sorted(wb.values.items()):
        if args.sheet and sheet != args.sheet:
            continue
        print(f"{sheet}!{index_to_col(c)}{r}\t{v}")
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="xl2tableau",
                                description="Migrate Excel calc columns to Tableau calculated fields.")
    sub = p.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("convert", help="write .twbx, CSV and an audit .xlsx")
    c.add_argument("input")
    c.add_argument("-o", "--out", default="out")
    c.add_argument("--strict", action="store_true",
                   help="exit with status 1 if any translated field disagrees with Excel")
    c.set_defaults(func=_cmd_convert)
    e = sub.add_parser("eval", help="recalculate every formula and print the values")
    e.add_argument("input")
    e.add_argument("--sheet")
    e.set_defaults(func=_cmd_eval)
    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
