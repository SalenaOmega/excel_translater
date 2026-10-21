"""
cleaner.py
Field renaming, type conversion, missing-rate statistics, and anomaly marking.
All functions accept and return plain DataFrames; they do not touch files.
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

# Sentinel used to distinguish "was already missing" from "failed to convert"
_MISSING = object()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_blank(val: Any) -> bool:
    """True for None, NaN, and whitespace-only strings."""
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() in ("", "nan", "NaN", "None", "NA", "N/A"):
        return True
    return False


def _normalize_nulls(df: pd.DataFrame) -> pd.DataFrame:
    """Replace blank-string sentinels with real NaN so missing_stats works correctly."""
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].apply(lambda v: None if _is_blank(v) else v)
    return df


# ---------------------------------------------------------------------------
# Field renaming
# ---------------------------------------------------------------------------

def rename_fields(
    df: pd.DataFrame,
    alias_map: dict[str, str],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Normalise column names then apply alias_map.

    Returns (df_renamed, conflict_warnings).
    conflict_warnings lists columns where two source fields mapped to the same
    target but had different non-null values on the same row (value conflict).

    When multiple sources map to the same target the first non-null value wins;
    conflicting rows are flagged with _alias_conflict_<col>.
    """
    df = df.copy()
    # Strip whitespace from all headers
    df.columns = [str(c).strip() for c in df.columns]
    # Trim object-column text while we're here
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)

    df = df.rename(columns=alias_map)

    conflicts: list[str] = []
    cols = list(df.columns)
    if len(cols) != len(set(cols)):
        seen: dict[str, pd.Series] = {}
        conflict_masks: dict[str, pd.Series] = {}
        for col in cols:
            raw_col = df[col]
            if isinstance(raw_col, pd.DataFrame):
                series_list = [raw_col.iloc[:, i] for i in range(raw_col.shape[1])]
            else:
                series_list = [raw_col]

            if col not in seen:
                seen[col] = series_list[0]
                conflict_masks[col] = pd.Series(False, index=df.index)
                for s in series_list[1:]:
                    # Conflict: both are non-null and differ
                    both_present = seen[col].notna() & s.notna()
                    differ = seen[col] != s
                    conflict_masks[col] |= both_present & differ
                    seen[col] = seen[col].combine_first(s)
            else:
                for s in series_list:
                    both_present = seen[col].notna() & s.notna()
                    differ = seen[col] != s
                    conflict_masks[col] |= both_present & differ
                    seen[col] = seen[col].combine_first(s)

        df = pd.DataFrame(seen)

        # Add a boolean flag column for each conflicting target field
        for col, mask in conflict_masks.items():
            if mask.any():
                conflicts.append(col)
                df[f"_alias_conflict_{col}"] = mask

    return df, conflicts


# ---------------------------------------------------------------------------
# Type conversion — preserving original values for failed rows
# ---------------------------------------------------------------------------

def convert_types(
    df: pd.DataFrame,
    type_map: dict[str, str],
) -> tuple[pd.DataFrame, dict[str, int]]:
    """
    Convert columns per type_map {"col": "numeric"|"date"|"str"}.

    For numeric and date columns:
      - Successful conversions replace the value in-place.
      - Failed conversions: the original value is preserved in a
        _orig_<col> column; the main column gets NaN; the row is tagged
        conv_fail_<col> in _anomaly_flags (created here if absent).

    Returns (df_converted, {col: failure_count}).
    """
    df = df.copy()
    if "_anomaly_flags" not in df.columns:
        df["_anomaly_flags"] = ""

    failures: dict[str, int] = {}

    for col, kind in type_map.items():
        if col not in df.columns:
            continue

        if kind == "str":
            df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
            continue

        orig = df[col].copy()

        if kind == "numeric":
            cleaned = orig.astype(str).str.replace(r"[¥$€,\s]", "", regex=True)
            converted = pd.to_numeric(cleaned, errors="coerce")
        else:  # date
            converted = pd.to_datetime(orig, errors="coerce")

        # Rows that were already NaN/blank before conversion are not failures
        was_blank = orig.apply(_is_blank)
        new_null = converted.isna() & ~was_blank
        n = int(new_null.sum())

        if n:
            failures[col] = n
            # Save original value so the Exceptions sheet is readable
            orig_col = f"_orig_{col}"
            if orig_col not in df.columns:
                df[orig_col] = None
            df.loc[new_null, orig_col] = orig[new_null]
            # Tag _anomaly_flags
            flag = f"conv_fail_{col}"
            df.loc[new_null, "_anomaly_flags"] = df.loc[new_null, "_anomaly_flags"].apply(
                lambda x: f"{x},{flag}" if x else flag
            )

        df[col] = converted

    return df, failures


# ---------------------------------------------------------------------------
# Missing-rate statistics
# ---------------------------------------------------------------------------

def missing_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a DataFrame (Column | Missing Count | Missing % | Total Rows).
    Counts true NaN/None after _normalize_nulls has been applied.
    Sorted descending by missing %.
    """
    work = _normalize_nulls(df)
    total = len(work)
    counts = work.isna().sum()
    pct = (counts / total * 100).round(2) if total else counts * 0
    stats = pd.DataFrame({
        "Column": counts.index,
        "Missing Count": counts.values,
        "Missing %": pct.values,
        "Total Rows": total,
    })
    return stats.sort_values("Missing %", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Anomaly marking
# ---------------------------------------------------------------------------

def _iqr_bounds(series: pd.Series) -> tuple[float, float]:
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    return q1 - 1.5 * iqr, q3 + 1.5 * iqr


def mark_anomalies(
    df: pd.DataFrame,
    rules: list[dict[str, Any]],
    *,
    iqr_cols: list[str] | None = None,
) -> pd.DataFrame:
    """
    Append to the _anomaly_flags column (created by convert_types or here).
    Rules: {"name": "neg_qty", "col": "Quantity", "condition": "< 0"}.
    iqr_cols: columns to flag with IQR outlier detection.
    """
    df = df.copy()
    if "_anomaly_flags" not in df.columns:
        df["_anomaly_flags"] = ""

    op_re = re.compile(r"^([<>=!]+)\s*([\d.eE+\-]+)$")
    for rule in rules:
        col = rule.get("col", "")
        cond = rule.get("condition", "")
        name = rule.get("name", col)
        if col not in df.columns:
            continue
        m = op_re.match(cond.strip())
        if not m:
            continue
        op, val_str = m.group(1), m.group(2)
        val = float(val_str)
        numeric = pd.to_numeric(df[col], errors="coerce")
        ops = {
            "<":  numeric < val,  "<=": numeric <= val,
            ">":  numeric > val,  ">=": numeric >= val,
            "==": numeric == val, "!=": numeric != val,
        }
        mask = ops.get(op, pd.Series(False, index=df.index))
        df.loc[mask, "_anomaly_flags"] = df.loc[mask, "_anomaly_flags"].apply(
            lambda x: f"{x},{name}" if x else name
        )

    for col in (iqr_cols or []):
        if col not in df.columns:
            continue
        numeric = pd.to_numeric(df[col], errors="coerce")
        valid = numeric.dropna()
        if len(valid) < 4:
            continue
        lo, hi = _iqr_bounds(valid)
        mask = (numeric < lo) | (numeric > hi)
        flag_name = f"iqr_{col}"
        df.loc[mask, "_anomaly_flags"] = df.loc[mask, "_anomaly_flags"].apply(
            lambda x: f"{x},{flag_name}" if x else flag_name
        )

    return df


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def deduplicate(df: pd.DataFrame, key_cols: list[str] | None = None) -> pd.DataFrame:
    """
    Drop duplicate rows, excluding internal helper columns from the key.
    Fully-blank data rows (all non-internal fields null) are dropped first.
    """
    internal = {c for c in df.columns if c.startswith("_")}
    data_cols = [c for c in df.columns if c not in internal]

    # Drop rows where every data column is blank
    df = df[~df[data_cols].apply(lambda r: r.apply(_is_blank).all(), axis=1)].copy()

    if key_cols:
        subset = [c for c in key_cols if c in df.columns]
    else:
        subset = data_cols

    return df.drop_duplicates(subset=subset or None).reset_index(drop=True)
