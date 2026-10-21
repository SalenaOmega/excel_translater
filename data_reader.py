"""
data_reader.py
Unified reader for Excel (.xlsx/.xls), CSV, and static HTML tables.
Returns a DataFrame with a _source column tagging the origin.
"""
from __future__ import annotations

import glob
import os

import pandas as pd


def _try_csv(path: str) -> pd.DataFrame:
    """Try common encodings for CSV; raise ValueError if all fail."""
    for enc in ("utf-8-sig", "utf-8", "gbk", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc, dtype=str)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode CSV with any supported encoding: {path}")


def read_file(
    path: str,
    *,
    sheet: int | str = 0,
    table_index: int = 0,
) -> pd.DataFrame:
    """
    Read a single file or URL into a DataFrame.

    Parameters
    ----------
    path        : local file path, or http/https URL pointing to an HTML page
    sheet       : worksheet name or 0-based integer index (Excel only)
    table_index : which <table> to use when the source is HTML (0 = first)

    Returns
    -------
    DataFrame with a _source column containing the filename or URL.
    """
    lower = path.lower()

    if lower.startswith("http://") or lower.startswith("https://"):
        tables = pd.read_html(path, flavor="lxml")
        if table_index >= len(tables):
            raise IndexError(
                f"Page has {len(tables)} table(s); table_index={table_index} out of range."
            )
        df = tables[table_index].astype(str)
        source_tag = path

    elif lower.endswith(".html") or lower.endswith(".htm"):
        tables = pd.read_html(path, flavor="lxml")
        if table_index >= len(tables):
            raise IndexError(
                f"File has {len(tables)} table(s); table_index={table_index} out of range."
            )
        df = tables[table_index].astype(str)
        source_tag = os.path.basename(path)

    elif lower.endswith(".csv"):
        df = _try_csv(path)
        source_tag = os.path.basename(path)

    else:  # .xlsx / .xls
        # Accept both integer index and sheet name; coerce "0" → 0
        sheet_arg: int | str = sheet
        if isinstance(sheet_arg, str) and sheet_arg.isdigit():
            sheet_arg = int(sheet_arg)
        df = pd.read_excel(path, sheet_name=sheet_arg, dtype=str)
        source_tag = os.path.basename(path)

    df["_source"] = source_tag
    return df


def read_all(
    input_dir: str,
    *,
    sheet: int | str = 0,
    table_index: int = 0,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Read all Excel/CSV/HTML files in a folder and concatenate them.
    Skips temp lock files (~ prefix).

    Returns
    -------
    (combined_df, failed_files)
    failed_files is a list of basenames that could not be read.
    Raises RuntimeError if ALL files fail.
    """
    exts = ("*.xlsx", "*.xls", "*.csv", "*.html", "*.htm")
    files: list[str] = []
    for pat in exts:
        files.extend(glob.glob(os.path.join(input_dir, pat)))
    files = sorted(f for f in files if not os.path.basename(f).startswith("~$"))

    if not files:
        raise FileNotFoundError(f"No supported files found in: {input_dir}")

    frames: list[pd.DataFrame] = []
    failed: list[str] = []
    for f in files:
        try:
            frames.append(read_file(f, sheet=sheet, table_index=table_index))
            print(f"  [read]  {os.path.basename(f)}")
        except Exception as exc:
            failed.append(os.path.basename(f))
            print(f"  [FAIL]  {os.path.basename(f)} — {exc}")

    if not frames:
        raise RuntimeError("All files failed to read.")

    return pd.concat(frames, ignore_index=True), failed
