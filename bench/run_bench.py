"""Benchmark: recalculation throughput, cache speedup, engine accuracy, conversion fidelity.

Usage:
    python bench/run_bench.py            # full run (sizes up to 50,000 rows)
    python bench/run_bench.py --quick    # small sizes, used in CI

All numbers in the README come from this script.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import platform
import sys
import tempfile
import time

from xl2tableau.convert import convert_file
from xl2tableau.formula.values import XlError
from xl2tableau.sample import HEADERS, build_sales_workbook, expected_values


def best_of(fn, repeats):
    best, out = math.inf, None
    for _ in range(repeats):
        t0 = time.perf_counter()
        out = fn()
        best = min(best, time.perf_counter() - t0)
    return best, out


def bench_recalc(n, use_index, repeats):
    wb, rows = build_sales_workbook(n, seed=42, as_engine=True)
    secs, stats = best_of(lambda: wb.recalc(use_index=use_index), repeats)
    return wb, rows, secs, stats


def accuracy(wb, rows):
    exp = expected_values(rows)
    total = ok = 0
    for i, e in enumerate(exp):
        for name, want in e.items():
            got = wb.get("Sales", i + 2, HEADERS.index(name) + 1)
            total += 1
            if want is None:
                ok += isinstance(got, XlError)
            elif isinstance(want, float):
                ok += isinstance(got, float) and math.isclose(got, want, rel_tol=1e-9, abs_tol=1e-9)
            else:
                ok += got == want
    return ok, total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--json", default=None, help="write results to this file")
    args = ap.parse_args()

    sizes = [500, 2000] if args.quick else [1000, 10000, 50000]
    naive_sizes = [250, 500] if args.quick else [500, 1000, 2000]
    convert_n = 1000 if args.quick else 10000
    repeats = 1 if args.quick else 3
    results = {"python": platform.python_version(), "machine": platform.machine(),
               "recalc": [], "cache": [], "convert": None}

    print(f"Python {results['python']} on {platform.system()} {results['machine']}\n")
    print("## Recalculation (hash indexes + aggregate memo on)\n")
    print("| rows | formulas | parse s | graph s | eval s | total s | formulas/s | accuracy vs ground truth |")
    print("|---:|---:|---:|---:|---:|---:|---:|---:|")
    for n in sizes:
        wb, rows, secs, st = bench_recalc(n, True, repeats)
        ok, total = accuracy(wb, rows)
        rec = dict(rows=n, formulas=st.formulas, parse_s=st.parse_seconds, graph_s=st.graph_seconds,
                   eval_s=st.eval_seconds, total_s=secs, per_s=st.formulas / secs,
                   correct=ok, checked=total)
        results["recalc"].append(rec)
        print(f"| {n:,} | {st.formulas:,} | {st.parse_seconds:.3f} | {st.graph_seconds:.3f} | "
              f"{st.eval_seconds:.3f} | {secs:.3f} | {st.formulas / secs:,.0f} | {ok:,}/{total:,} |")

    print("\n## Effect of criteria hash index and aggregate memo\n")
    print("| rows | caches off s | caches on s | speedup |")
    print("|---:|---:|---:|---:|")
    for n in naive_sizes:
        _, _, off, _ = bench_recalc(n, False, 1)
        _, _, on, _ = bench_recalc(n, True, repeats)
        results["cache"].append(dict(rows=n, off_s=off, on_s=on, speedup=off / on))
        print(f"| {n:,} | {off:.3f} | {on:.3f} | {off / on:.1f}x |")

    print(f"\n## Full conversion of a {convert_n:,}-row .xlsx (one hardcoded override at row 9)\n")
    with tempfile.TemporaryDirectory() as tmp:
        x, _ = build_sales_workbook(convert_n, seed=42, override_row=9)
        src = os.path.join(tmp, "model.xlsx")
        x.save(src)
        t0 = time.perf_counter()
        res = convert_file(src, os.path.join(tmp, "out"))
        wall = time.perf_counter() - t0
        twbx_kb = os.path.getsize(res.outputs["twbx"]) / 1024
    conv = dict(rows=convert_n, wall_s=wall, **{f"{k}_s": v for k, v in res.timings.items()},
                calc_columns=res.calc_columns, translated=res.translated,
                rows_checked=res.rows_checked, rows_matched=res.rows_matched,
                inconsistent_cells=sum(len(c.inconsistent_rows) for t in res.tables for c in t.columns),
                twbx_kb=twbx_kb)
    results["convert"] = conv
    print("| metric | value |")
    print("|---|---:|")
    print(f"| wall time (load + recalc + translate + write + validate + audit) | {wall:.2f} s |")
    for k, v in res.timings.items():
        print(f"| {k} stage | {v:.3f} s |")
    print(f"| calc columns translated | {res.translated}/{res.calc_columns} |")
    print(f"| row-level checks agreeing with Excel | {res.rows_matched:,}/{res.rows_checked:,} |")
    print(f"| cells flagged as differing from their column pattern | {conv['inconsistent_cells']} |")
    print(f"| .twbx size | {twbx_kb:,.0f} KiB |")

    if args.json:
        with open(args.json, "w") as f:
            json.dump(results, f, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
