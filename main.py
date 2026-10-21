"""
main.py
Entry point for the sales data pipeline.

Usage:
    python main.py --input sample_data --output output/report.xlsx
    python main.py --input ./monthly_files --config my_config.json
    python main.py --input sample_data --url https://example.com/table.html
"""
from __future__ import annotations

import argparse
import json
import sys
import time

import pandas as pd

from cleaner import convert_types, deduplicate, mark_anomalies, missing_stats, rename_fields
from data_reader import read_all, read_file
from reporter import build_summaries, write_report


def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.pop("_comment", None)
    return cfg


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser(
        description="Combine, clean, and summarize Excel/CSV/HTML files into one report."
    )
    p.add_argument("--input",  default="sample_data",        help="Folder of source files")
    p.add_argument("--output", default="output/report.xlsx", help="Output .xlsx path")
    p.add_argument("--config", default="config.json",        help="JSON config file")
    p.add_argument("--sheet",  default="0",                  help="Worksheet name or index (Excel)")
    p.add_argument("--table",  default=0, type=int,          help="HTML table index (0 = first)")
    p.add_argument("--url",    default=None,
                   help="Additional HTML page URL to include (appended to folder files)")
    p.add_argument("--partial", action="store_true",
                   help="Continue and write report even if some files fail to read")
    args = p.parse_args(argv)

    # Coerce --sheet to int when it looks like a number
    sheet: int | str = int(args.sheet) if args.sheet.isdigit() else args.sheet

    t0 = time.perf_counter()

    cfg = load_config(args.config)
    alias_map     = cfg.get("alias_map", {})
    type_map      = cfg.get("type_map", {})
    date_col      = cfg.get("date_col", "Date")
    amount_col    = cfg.get("amount_col", "Amount")
    group_cols    = cfg.get("group_cols", ["Category", "Region", "Month"])
    dedup_keys    = cfg.get("dedup_key_cols", [])
    anomaly_rules = cfg.get("anomaly_rules", [])
    iqr_cols      = cfg.get("iqr_cols", [])

    # ── 1. Read ───────────────────────────────────────────────────────────────
    print(f"\n[1/5] Reading files from '{args.input}' ...")
    try:
        raw, failed = read_all(args.input, sheet=sheet, table_index=args.table)
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    # Append a URL source if requested
    if args.url:
        try:
            url_df = read_file(args.url, table_index=args.table)
            raw = pd.concat([raw, url_df], ignore_index=True)
            print(f"  [read]  {args.url}")
        except Exception as exc:
            failed.append(args.url)
            print(f"  [FAIL]  {args.url} — {exc}")

    if failed:
        print(f"\n  WARNING: {len(failed)} source(s) could not be read:")
        for f in failed:
            print(f"    - {f}")
        if not args.partial:
            print(
                "  Aborting. Run with --partial to generate a report from the files that did load.",
                file=sys.stderr,
            )
            return 3
        print("  --partial flag set: continuing with successfully read files.")

    print(f"      {len(raw)} rows from {raw['_source'].nunique()} file(s).")

    # ── 2. Rename & convert types ─────────────────────────────────────────────
    print("[2/5] Renaming fields and converting types ...")
    df, alias_conflicts = rename_fields(raw, alias_map)
    if alias_conflicts:
        for col in alias_conflicts:
            n = int(df[f"_alias_conflict_{col}"].sum())
            print(f"      WARNING: '{col}' had value conflicts on {n} row(s) — flagged in _alias_conflict_{col}.")

    # convert_types injects _anomaly_flags and _orig_<col> for failed rows
    df, conv_failures = convert_types(df, type_map)
    if conv_failures:
        for col, n in conv_failures.items():
            print(f"      WARNING: {n} row(s) in '{col}' could not convert — original values in _orig_{col}.")

    # Derive Month grouping column from parsed date
    if date_col in df.columns and pd.api.types.is_datetime64_any_dtype(df[date_col]):
        df["Month"] = df[date_col].dt.strftime("%Y-%m")

    # ── 3. Clean ──────────────────────────────────────────────────────────────
    print("[3/5] Deduplicating and computing data quality ...")
    before = len(df)
    df = deduplicate(df, key_cols=dedup_keys or None)
    removed = before - len(df)
    if removed:
        print(f"      Removed {removed} duplicate/blank row(s).")

    quality = missing_stats(df)
    high_missing = quality[quality["Missing %"] > 20]
    if not high_missing.empty:
        print(f"      NOTE: {len(high_missing)} column(s) with >20% missing values.")

    # ── 4. Anomaly marking ────────────────────────────────────────────────────
    print("[4/5] Marking anomalies ...")
    df = mark_anomalies(df, anomaly_rules, iqr_cols=iqr_cols)
    n_flagged = int((df["_anomaly_flags"] != "").sum())
    print(f"      {n_flagged} row(s) flagged (incl. conversion failures).")

    # ── 5. Write report ───────────────────────────────────────────────────────
    print(f"[5/5] Writing report to '{args.output}' ...")
    summaries = build_summaries(df, amount_col, group_cols)
    n_files = raw["_source"].nunique()
    write_report(
        df, quality, summaries, conv_failures, failed,
        amount_col, args.output, n_files,
    )

    elapsed = time.perf_counter() - t0
    if amount_col in df.columns:
        total = df[amount_col].sum()
        print(f"\n      Total {amount_col}: {total:,.2f}")
    print(f"      Clean rows  : {len(df)}")
    print(f"      Anomaly rows: {n_flagged}")
    if failed:
        print(f"      Skipped files: {len(failed)} (partial mode)")
    print(f"\nDone in {elapsed:.1f}s  →  {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
