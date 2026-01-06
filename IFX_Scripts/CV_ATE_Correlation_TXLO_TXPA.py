"""\
CV ↔ ATE correlation (flat script, no classes).

Reads a single raw data Excel file that contains CV and ATE data (typically in
separate sheets), merges rows by device/test keys, then groups by:
    Test Number, supply corner, frequency, and temperature
and performs per-group correlation using a robust offset-only model:
    - correlation factor = median(CV - ATE)

For each group (test case), the script:
    - Computes delta = (CV - ATE)
    - Uses median(delta) as the correction factor
    - Produces corrected ATE: ATE_correlated = ATE + median(delta)
    - Computes residuals as: residual = delta - median(delta)
    - Computes σ(residual) and a guard-band (X * σ)
    - Computes an R² value for the offset-only model (ATE_pred = CV - median(delta))
    - Outputs correlation factors + derived limits to an Excel file
    - Outputs row-level corrected values + residuals to a second Excel sheet

This script is designed to be configured in-code (no CLI required).
"""

from __future__ import annotations

import math
import re
import textwrap
from pathlib import Path

import pandas as pd


# =========================
# USER CONFIG (in-code)
# =========================

INPUT_XLSX = r"C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/CV_TE_Correlation/ATE_Extracted_PA_Power_Data.xlsx"  # e.g. r"C:\path\to\raw_cv_ate.xlsx"
OUTPUT_XLSX = r"C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/CV_TE_Correlation/Correlation/Correlation_PA_Power.xlsx"  # e.g. r"C:\path\to\correlation_results.xlsx" OR a folder path
OUTPUT_PLOTS_DIR = r"C:/UserData/Infineon/TE_CTRX/CTRX8188/Data_Reviews/CV_TE_Correlation/Correlation/plots_PA_Power"  # optional; if empty uses OUTPUT_XLSX folder + "plots"

# If you want to process multiple same-layout sheets (CV+ATE columns in each sheet), list them here.
# When non-empty, this overrides CV_SHEET/ATE_SHEET and runs each sheet independently.
SHEETS_TO_RUN = ["FE_Filtered"]

CV_SHEET = r""  # used when SHEETS_TO_RUN is empty
ATE_SHEET = r""  # used when SHEETS_TO_RUN is empty

# Columns that identify the same device/test row in both sheets.
# These should exist in both CV and ATE sheets.
MERGE_KEYS = [
    "DUT Nr",
    "Wafer",
    "X",
    "Y",
    "Temperature",
    "Voltage corner",
    "Frequency_GHz",
    "Test Number",
]

# Required columns containing numeric results
# For the provided PN correlation workbook these are in the same sheet.
CV_VALUE_COL = "CV_PA_Power"    # in-sheet CV values
ATE_VALUE_COL = "ATE_PA_Power"  # in-sheet ATE values

# Optional columns (if present) for limits (ATE limits)
ATE_LOW_COL = "Low"
ATE_HIGH_COL = "High"
ATE_UNIT_COL = "Unit"

# Test-case definition
# - TXLO power: grouped by Test Number → Voltage corner → Frequency → Temperature
# - TXPA power: grouped by LUT value → Voltage corner → Frequency → Temperature
# Set to "auto" to pick TXPA grouping when a usable LUT value is present.
TEST_CASE_TYPE = "auto"  # "TXLO" | "TXPA" | "auto"

TXLO_GROUP_COLS = ["Test Number", "Voltage corner", "Frequency_GHz", "Temperature"]
TXPA_GROUP_COLS = ["LUT value", "Voltage corner", "Frequency_GHz", "Temperature"]

# Regression / guard-band settings
MIN_POINTS_PER_GROUP = 5
X_SIGMA_GUARDBAND = 3.0  # guard-band = X * σ(residuals)

# Plot settings
PLOT_DPI = 160


# =========================
# Helpers
# =========================

def _as_path_maybe_folder(path_str: str, default_filename: str) -> Path:
    p = Path(path_str)
    if p.suffix.lower() != ".xlsx":
        return p / default_filename
    return p


def _to_float_series(s: pd.Series) -> pd.Series:
    # Handles European decimals and stray whitespace.
    # Keeps NaNs as NaN.
    s2 = s.astype(str).str.strip()
    s2 = s2.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA})
    s2 = s2.str.replace(" ", "", regex=False)
    s2 = s2.str.replace(",", ".", regex=False)
    return pd.to_numeric(s2, errors="coerce")


def _safe_slug(text: str) -> str:
    t = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(text))
    return t.strip("_")[:180] or "plot"


def _find_test_name_column(columns: list[str]) -> str | None:
    # Best-effort detection of the human-readable test name column.
    if not columns:
        return None

    def _norm(s: str) -> str:
        return re.sub(r"[\s_]+", "", str(s)).lower().strip()

    norm_map = {_norm(c): c for c in columns}
    for key in ("testname", "testcasename"):
        if key in norm_map:
            return norm_map[key]
    for c in columns:
        if "testname" in _norm(c):
            return c
    return None


def _find_doe_split_column(columns: list[str]) -> str | None:
    if not columns:
        return None

    def _norm(s: str) -> str:
        return re.sub(r"[\s_\-]+", "", str(s)).lower().strip()

    norm_map = {_norm(c): c for c in columns}
    for key in ("doesplit", "doe", "doe_split"):
        if key in norm_map:
            return norm_map[key]
    for c in columns:
        if "doe" in _norm(c) and "split" in _norm(c):
            return c
    return None


_FWLU_RE = re.compile(r"FwLu(?P<lut>\d{2,3})(?!\d)", flags=re.IGNORECASE)


def _extract_lut_value(test_name: str):
    m = _FWLU_RE.search(str(test_name) if test_name is not None else "")
    if not m:
        return pd.NA
    try:
        return int(m.group("lut"))
    except Exception:
        return pd.NA


def _ensure_lut_value_column(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if "LUT value" in df.columns:
        # Normalize type (keeps missing values)
        df["LUT value"] = pd.to_numeric(df["LUT value"], errors="coerce").astype("Int64")
        return df
    if "Test Name" in df.columns:
        df["LUT value"] = df["Test Name"].astype(str).map(_extract_lut_value).astype("Int64")
        return df
    df["LUT value"] = pd.Series([pd.NA] * len(df), dtype="Int64")
    return df


def _resolve_group_cols(df: pd.DataFrame) -> list[str]:
    mode = str(TEST_CASE_TYPE).strip().upper()
    if mode == "TXLO":
        return TXLO_GROUP_COLS
    if mode == "TXPA":
        return TXPA_GROUP_COLS

    # auto
    if df is not None and not df.empty:
        if "LUT value" in df.columns:
            cand = pd.to_numeric(df["LUT value"], errors="coerce")
            if cand.notna().any():
                return TXPA_GROUP_COLS
        if "Test Name" in df.columns:
            cand = df["Test Name"].astype(str).map(_extract_lut_value)
            if pd.Series(cand).notna().any():
                return TXPA_GROUP_COLS
    return TXLO_GROUP_COLS


def _insert_column_after(df: pd.DataFrame, after_col: str, col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if col not in df.columns:
        return df
    cols = list(df.columns)
    cols.remove(col)
    if after_col in cols:
        idx = cols.index(after_col) + 1
        cols.insert(idx, col)
        return df[cols]
    # If the target column isn't present, keep as-is.
    return df


def _r2_score(y_true: pd.Series, y_pred: pd.Series) -> float:
    yt = y_true.to_numpy(dtype=float)
    yp = y_pred.to_numpy(dtype=float)
    if len(yt) < 2:
        return math.nan
    y_mean = float(yt.mean())
    ss_tot = float(((yt - y_mean) ** 2).sum())
    if ss_tot == 0.0:
        return math.nan
    ss_res = float(((yt - yp) ** 2).sum())
    return 1.0 - (ss_res / ss_tot)


def _add_txpa_por_baseline_columns(merged_local: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """Add POR-baseline columns for TXPA-style grouping only.

    Implements the Excel-like logic described as "delta error considering POR",
    but uses a robust baseline:
        delta_considering_por = (CV - ATE) - median_por(CV - ATE)

    Where the POR mean is computed per base group that excludes Voltage corner:
        (LUT value, Frequency_GHz, Temperature)

    POR selection is taken from the "DoE split" column (POR/SS/FF, etc).
    If POR rows are missing for a base group, the baseline stays NaN.
    """
    if merged_local is None or merged_local.empty:
        return merged_local

    # Only for TXPA groupings (heuristic: grouped by LUT value).
    if "LUT value" not in group_cols:
        return merged_local
    doe_col = "DoE split" if "DoE split" in merged_local.columns else _find_doe_split_column(list(merged_local.columns))
    if not doe_col:
        return merged_local

    # Build a normalized base-key (excluding Voltage corner) so POR baselines
    # match across corners even when source formatting differs (e.g. 2.4 vs 2.400).
    base_key_cols: list[str] = []

    df = merged_local.copy()

    # Compute delta once at the row level so we can baseline it.
    df["Delta(CV-ATE)"] = df["CV"] - df["ATE"]

    # Normalize base keys (TXPA expected: LUT value + frequency + temperature).
    if "LUT value" in df.columns:
        df["__base_lut"] = pd.to_numeric(df["LUT value"], errors="coerce").astype("Int64")
        base_key_cols.append("__base_lut")
    if "Frequency_GHz" in df.columns:
        df["__base_freq"] = pd.to_numeric(df["Frequency_GHz"], errors="coerce")
        base_key_cols.append("__base_freq")
    if "Temperature" in df.columns:
        df["__base_temp"] = df["Temperature"].astype(str).str.strip().str.upper()
        base_key_cols.append("__base_temp")

    if not base_key_cols:
        return merged_local

    doe = df[doe_col].astype(str).str.strip().str.upper()
    # Some data sources include POR variants (e.g. "POR_NOM"); treat those as POR.
    por_mask = doe.str.contains("POR", na=False)

    por_df = df.loc[por_mask, base_key_cols + ["Delta(CV-ATE)"]].dropna(subset=["Delta(CV-ATE)"])
    if por_df.empty:
        df["POR_MedianDelta(CV-ATE)"] = math.nan
        df["Delta(CV-ATE)-POR"] = math.nan
        return df

    por_base = (
        por_df.groupby(base_key_cols, dropna=False)["Delta(CV-ATE)"]
        .median()
        .rename("POR_MedianDelta(CV-ATE)")
        .reset_index()
    )

    df = df.merge(por_base, how="left", on=base_key_cols, validate="many_to_one")
    df["Delta(CV-ATE)-POR"] = df["Delta(CV-ATE)"] - df["POR_MedianDelta(CV-ATE)"]
    return df


# =========================
# Main
# =========================

if __name__ == "__main__":
    if not str(INPUT_XLSX).strip():
        raise SystemExit("INPUT_XLSX is empty. Set it in the USER CONFIG block.")
    if not str(OUTPUT_XLSX).strip():
        raise SystemExit("OUTPUT_XLSX is empty. Set it in the USER CONFIG block.")
    if SHEETS_TO_RUN:
        sheets_to_run = [str(s).strip() for s in SHEETS_TO_RUN if str(s).strip()]
        if not sheets_to_run:
            raise SystemExit("SHEETS_TO_RUN is set but empty after stripping. Fix USER CONFIG.")
    else:
        if not str(CV_SHEET).strip() or not str(ATE_SHEET).strip():
            raise SystemExit("CV_SHEET / ATE_SHEET are empty. Set them in the USER CONFIG block.")
        sheets_to_run = [str(CV_SHEET).strip()]
        if str(CV_SHEET).strip() != str(ATE_SHEET).strip():
            sheets_to_run = []

    input_xlsx = Path(INPUT_XLSX)
    if not input_xlsx.is_file():
        raise SystemExit(f"Input file not found: {input_xlsx}")

    output_xlsx = _as_path_maybe_folder(OUTPUT_XLSX, "CV_ATE_Correlation.xlsx")
    plots_dir = Path(OUTPUT_PLOTS_DIR) if str(OUTPUT_PLOTS_DIR).strip() else output_xlsx.parent / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)

    # Prepare combined outputs
    all_factors_rows = []
    all_data_rows = []

    # Plot dependency
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    def run_one_same_sheet(sheet_name: str):
        df = pd.read_excel(input_xlsx, sheet_name=sheet_name)
        test_name_col = _find_test_name_column(list(df.columns))
        doe_split_col = _find_doe_split_column(list(df.columns))
        missing = [c for c in (MERGE_KEYS + [CV_VALUE_COL, ATE_VALUE_COL]) if c not in df.columns]
        if missing:
            raise SystemExit(f"Sheet '{sheet_name}' missing columns: {missing}")

        keep = list(dict.fromkeys(MERGE_KEYS + [CV_VALUE_COL, ATE_VALUE_COL, ATE_LOW_COL, ATE_HIGH_COL, ATE_UNIT_COL]))
        if "LUT value" in df.columns:
            keep.append("LUT value")
        if test_name_col:
            keep.append(test_name_col)
        if doe_split_col:
            keep.append(doe_split_col)
        keep = [c for c in keep if c in df.columns]
        merged_local = df[keep].copy()

        merged_local[CV_VALUE_COL] = _to_float_series(merged_local[CV_VALUE_COL])
        merged_local[ATE_VALUE_COL] = _to_float_series(merged_local[ATE_VALUE_COL])

        for k in MERGE_KEYS:
            if k in ("X", "Y", "Test Number"):
                merged_local[k] = pd.to_numeric(merged_local[k], errors="coerce")
            else:
                merged_local[k] = merged_local[k].astype(str).str.strip()

        merged_local = merged_local.dropna(subset=[CV_VALUE_COL, ATE_VALUE_COL])
        rename_map = {CV_VALUE_COL: "CV", ATE_VALUE_COL: "ATE"}
        if test_name_col and test_name_col in merged_local.columns:
            rename_map[test_name_col] = "Test Name"
        if doe_split_col and doe_split_col in merged_local.columns and doe_split_col != "DoE split":
            rename_map[doe_split_col] = "DoE split"
        merged_local = merged_local.rename(columns=rename_map)
        merged_local = _ensure_lut_value_column(merged_local)
        return merged_local

    def run_one_two_sheets():
        cv_df = pd.read_excel(input_xlsx, sheet_name=CV_SHEET)
        ate_df = pd.read_excel(input_xlsx, sheet_name=ATE_SHEET)

        test_name_cv = _find_test_name_column(list(cv_df.columns))
        test_name_ate = _find_test_name_column(list(ate_df.columns))
        doe_split_cv = _find_doe_split_column(list(cv_df.columns))
        doe_split_ate = _find_doe_split_column(list(ate_df.columns))

        missing_cv = [c for c in (MERGE_KEYS + [CV_VALUE_COL]) if c not in cv_df.columns]
        missing_ate = [c for c in (MERGE_KEYS + [ATE_VALUE_COL]) if c not in ate_df.columns]
        if missing_cv:
            raise SystemExit(f"CV sheet missing columns: {missing_cv}")
        if missing_ate:
            raise SystemExit(f"ATE sheet missing columns: {missing_ate}")

        cv_keep = list(dict.fromkeys(MERGE_KEYS + [CV_VALUE_COL, ATE_LOW_COL, ATE_HIGH_COL, ATE_UNIT_COL]))
        if "LUT value" in cv_df.columns:
            cv_keep.append("LUT value")
        if test_name_cv:
            cv_keep.append(test_name_cv)
        if doe_split_cv:
            cv_keep.append(doe_split_cv)
        cv_keep = [c for c in cv_keep if c in cv_df.columns]
        ate_keep = list(dict.fromkeys(MERGE_KEYS + [ATE_VALUE_COL]))
        if "LUT value" in ate_df.columns:
            ate_keep.append("LUT value")
        if test_name_ate:
            ate_keep.append(test_name_ate)
        if doe_split_ate:
            ate_keep.append(doe_split_ate)

        cv_df = cv_df[cv_keep].copy()
        ate_df = ate_df[ate_keep].copy()

        cv_df[CV_VALUE_COL] = _to_float_series(cv_df[CV_VALUE_COL])
        ate_df[ATE_VALUE_COL] = _to_float_series(ate_df[ATE_VALUE_COL])

        for k in MERGE_KEYS:
            if k in ("X", "Y", "Test Number"):
                cv_df[k] = pd.to_numeric(cv_df[k], errors="coerce")
                ate_df[k] = pd.to_numeric(ate_df[k], errors="coerce")
            else:
                cv_df[k] = cv_df[k].astype(str).str.strip()
                ate_df[k] = ate_df[k].astype(str).str.strip()

        cv_df = cv_df.dropna(subset=[CV_VALUE_COL])
        ate_df = ate_df.dropna(subset=[ATE_VALUE_COL])

        merged_local = pd.merge(
            cv_df,
            ate_df,
            how="inner",
            on=MERGE_KEYS,
            suffixes=("_CV", "_ATE"),
            validate="many_to_many",
        )

        merged_local = merged_local.rename(
            columns={
                f"{CV_VALUE_COL}_CV": "CV",
                f"{ATE_VALUE_COL}_ATE": "ATE",
            }
        )

        # Normalize/choose a single Test Name column if present
        if test_name_cv and f"{test_name_cv}_CV" in merged_local.columns:
            merged_local = merged_local.rename(columns={f"{test_name_cv}_CV": "Test Name"})
        elif test_name_cv and test_name_cv in merged_local.columns:
            merged_local = merged_local.rename(columns={test_name_cv: "Test Name"})
        elif test_name_ate and f"{test_name_ate}_ATE" in merged_local.columns:
            merged_local = merged_local.rename(columns={f"{test_name_ate}_ATE": "Test Name"})
        elif test_name_ate and test_name_ate in merged_local.columns:
            merged_local = merged_local.rename(columns={test_name_ate: "Test Name"})

        # Normalize/choose a single DoE split column if present
        if "DoE split" not in merged_local.columns:
            if doe_split_cv and f"{doe_split_cv}_CV" in merged_local.columns:
                merged_local = merged_local.rename(columns={f"{doe_split_cv}_CV": "DoE split"})
            elif doe_split_cv and doe_split_cv in merged_local.columns:
                merged_local = merged_local.rename(columns={doe_split_cv: "DoE split"})
            elif doe_split_ate and f"{doe_split_ate}_ATE" in merged_local.columns:
                merged_local = merged_local.rename(columns={f"{doe_split_ate}_ATE": "DoE split"})
            elif doe_split_ate and doe_split_ate in merged_local.columns:
                merged_local = merged_local.rename(columns={doe_split_ate: "DoE split"})

        # Normalize/choose a single LUT value column if present
        if "LUT value" not in merged_local.columns:
            if "LUT value_CV" in merged_local.columns:
                merged_local = merged_local.rename(columns={"LUT value_CV": "LUT value"})
            elif "LUT value_ATE" in merged_local.columns:
                merged_local = merged_local.rename(columns={"LUT value_ATE": "LUT value"})

        merged_local = _ensure_lut_value_column(merged_local)

        if "CV" not in merged_local.columns or "ATE" not in merged_local.columns:
            raise SystemExit(
                "After merge, CV/ATE columns not found. "
                "If CV_VALUE_COL and ATE_VALUE_COL are the same name, ensure suffix mapping is correct."
            )
        return merged_local

    def correlate_merged(merged_local: pd.DataFrame, sheet_label: str):
        if merged_local.empty:
            return

        group_cols = _resolve_group_cols(merged_local)
        if "LUT value" in group_cols:
            merged_local = _ensure_lut_value_column(merged_local)

        missing_group_cols = [c for c in group_cols if c not in merged_local.columns]
        if missing_group_cols:
            raise SystemExit(f"Missing GROUP_COLS in merged data: {missing_group_cols}")

        # TXPA only: compute a POR baseline (per LUT/freq/temp) and add a
        # per-row delta relative to that baseline.
        merged_local = _add_txpa_por_baseline_columns(merged_local, group_cols)

        local_plots_dir = plots_dir / _safe_slug(sheet_label)
        local_plots_dir.mkdir(parents=True, exist_ok=True)

        grouped = merged_local.groupby(group_cols, dropna=False)

        for group_key, g in grouped:
            g = g.dropna(subset=["ATE", "CV"]).copy()
            n_points = len(g)
            if n_points < MIN_POINTS_PER_GROUP:
                continue

            test_name = ""
            if "Test Name" in g.columns:
                vals = g["Test Name"].astype(str).replace({"nan": ""}).str.strip()
                vals = [v for v in vals.tolist() if v]
                uniq = list(dict.fromkeys(vals))
                if len(uniq) == 1:
                    test_name = uniq[0]
                elif len(uniq) > 1:
                    shown = uniq[:2]
                    test_name = "; ".join(shown) + (f" (+{len(uniq)-2} more)" if len(uniq) > 2 else "")

            # Offset-only correlation factors
            # delta = CV - ATE
            g["Delta(CV-ATE)"] = g["CV"] - g["ATE"]
            median_delta = float(g["Delta(CV-ATE)"].median())
            max_delta = float(g["Delta(CV-ATE)"].max())

            # TXPA only: include POR-baseline delta stats if available.
            por_median_delta = math.nan
            median_delta_minus_por = math.nan
            if "LUT value" in group_cols and "POR_MedianDelta(CV-ATE)" in g.columns:
                # Constant within the group (same LUT/freq/temp), but safe to read from series.
                por_vals = pd.to_numeric(g["POR_MedianDelta(CV-ATE)"], errors="coerce").dropna()
                por_median_delta = float(por_vals.iloc[0]) if not por_vals.empty else math.nan
                if "Delta(CV-ATE)-POR" in g.columns:
                    mdp = pd.to_numeric(g["Delta(CV-ATE)-POR"], errors="coerce").dropna()
                    median_delta_minus_por = float(mdp.median()) if not mdp.empty else math.nan

            # Apply correction to ATE so it aligns to CV
            g["ATE_correlated"] = g["ATE"] + median_delta

            # Residuals in delta-domain
            g["Residual"] = g["Delta(CV-ATE)"] - median_delta
            residual_std = float(g["Residual"].std(ddof=1)) if len(g) > 1 else math.nan
            guardband_ate = X_SIGMA_GUARDBAND * residual_std if not math.isnan(residual_std) else math.nan

            # R² for the offset-only model in ATE-domain: ATE_pred = CV - median_delta
            r2 = _r2_score(g["ATE"], g["CV"] - median_delta)

            # Limits: only ATE limits exist
            ate_low = None
            ate_high = None
            unit = ""
            if ATE_LOW_COL in g.columns:
                low_series = _to_float_series(g[ATE_LOW_COL])
                low_vals = low_series.dropna().tolist()
                ate_low = float(low_vals[0]) if low_vals else None
            if ATE_HIGH_COL in g.columns:
                high_series = _to_float_series(g[ATE_HIGH_COL])
                high_vals = high_series.dropna().tolist()
                ate_high = float(high_vals[0]) if high_vals else None
            if ATE_UNIT_COL in g.columns:
                unit_vals = g[ATE_UNIT_COL].astype(str).replace({"nan": ""}).str.strip()
                unit_vals = [v for v in unit_vals.tolist() if v]
                unit = unit_vals[0] if unit_vals else ""

            # New limits (requested): shrink using guardband
            new_ate_low = math.nan
            new_ate_high = math.nan
            if (ate_low is not None) and (ate_high is not None) and (not math.isnan(guardband_ate)):
                new_ate_low = ate_low + guardband_ate
                new_ate_high = ate_high - guardband_ate

            # Save factors row
            group_dict = dict(zip(group_cols, group_key if isinstance(group_key, tuple) else (group_key,)))
            all_factors_rows.append(
                {
                    "DataSheet": sheet_label,
                    **group_dict,
                    "Test Name": test_name,
                    "N": n_points,
                    "MedianDelta(CV-ATE)": median_delta,
                    "MaxDelta(CV-ATE)": max_delta,
                    # TXPA only (NaN for TXLO): POR-baseline columns
                    "POR_MedianDelta(CV-ATE)": por_median_delta,
                    "MedianDelta(CV-ATE)-POR": median_delta_minus_por,
                    "R2_OffsetModel": r2,
                    "ResidualStd(Delta)": residual_std,
                    "Guardband(Delta)": guardband_ate,
                    "ATE_Low": ate_low,
                    "ATE_High": ate_high,
                    "Unit": unit,
                    "New_ATE_Low": new_ate_low,
                    "New_ATE_High": new_ate_high,
                }
            )

            # Save row-level correlated data
            for _, r in g.iterrows():
                all_data_rows.append(
                    {
                        "DataSheet": sheet_label,
                        **{k: r[k] for k in MERGE_KEYS},
                        # Extra context (helps sign-off and debugging). Present for TXPA.
                        "LUT value": (int(r["LUT value"]) if "LUT value" in g.columns and pd.notna(r.get("LUT value")) else pd.NA),
                        "DoE split": (str(r.get("DoE split", "")).strip() if "DoE split" in g.columns else ""),
                        "Test Name": (str(r.get("Test Name", "")).strip() if "Test Name" in g.columns else ""),
                        "CV": float(r["CV"]),
                        "ATE": float(r["ATE"]),
                        "ATE_correlated": float(r["ATE_correlated"]),
                        "Delta(CV-ATE)": float(r["Delta(CV-ATE)"]),
                        # TXPA only (NaN for TXLO): per-row POR-baseline deltas
                        "POR_MedianDelta(CV-ATE)": (float(r["POR_MedianDelta(CV-ATE)"]) if "POR_MedianDelta(CV-ATE)" in g.columns and pd.notna(r["POR_MedianDelta(CV-ATE)"]) else math.nan),
                        "Delta(CV-ATE)-POR": (float(r["Delta(CV-ATE)-POR"]) if "Delta(CV-ATE)-POR" in g.columns and pd.notna(r["Delta(CV-ATE)-POR"]) else math.nan),
                        "Residual": float(r["Residual"]),
                    }
                )

            # Plot (2 subplots):
            # 1) per-sample series (CV vs index, ATE vs index)
            # 2) regression view (ATE vs CV) with Theil–Sen line (TS info only)
            if "DUT Nr" in g.columns:
                g_plot = g.sort_values(by=["DUT Nr"]).reset_index(drop=True)
                x_label = "Samples (sorted by DUT Nr)"
            else:
                g_plot = g.reset_index(drop=True)
                x_label = "Samples"

            x_idx = pd.Series(range(len(g_plot)))

            fig, (ax1, ax2) = plt.subplots(
                nrows=2,
                ncols=1,
                figsize=(12.0, 9.0),
                gridspec_kw={"height_ratios": [2.0, 2.0]},
            )

            ax1.plot(
                x_idx,
                g_plot["CV"],
                marker="o",
                linestyle="-",
                linewidth=2.2,
                markersize=6,
                label="CV (individual)",
            )
            ax1.plot(
                x_idx,
                g_plot["ATE"],
                marker="s",
                linestyle="--",
                linewidth=2.2,
                markersize=6,
                label="ATE (individual)",
            )

            # Original + new ATE limits (ATE domain)
            if ate_low is not None:
                ax1.axhline(ate_low, color="black", linestyle=":", linewidth=2.0, label="ATE Low")
            if ate_high is not None:
                ax1.axhline(ate_high, color="black", linestyle=":", linewidth=2.0, label="ATE High")
            if not math.isnan(new_ate_low) and not math.isnan(new_ate_high):
                ax1.axhline(new_ate_low, color="purple", linestyle="-.", linewidth=2.2, label="New ATE Low")
                ax1.axhline(new_ate_high, color="purple", linestyle="-.", linewidth=2.2, label="New ATE High")

            title_parts = [f"{k}={v}" for k, v in group_dict.items()]
            if test_name:
                title_parts.insert(1 if title_parts else 0, f"Test Name={test_name}")
            title = " | ".join(title_parts)
            title_wrapped = "\n".join(textwrap.wrap(f"{sheet_label} | {title}", width=110))
            fig.suptitle(title_wrapped, fontsize=13, y=0.98)

            ax1.set_xlabel(x_label, fontsize=12)
            ax1.set_ylabel(f"Value{(' ['+unit+']') if unit else ''}", fontsize=12)
            ax1.tick_params(axis="both", labelsize=11)
            ax1.grid(True, alpha=0.25)

            # Tight x/y scaling to data + limits
            ax1.set_xlim(-0.5, len(g_plot) - 0.5)
            y_candidates = [
                float(g_plot["CV"].min()),
                float(g_plot["CV"].max()),
                float(g_plot["ATE"].min()),
                float(g_plot["ATE"].max()),
            ]
            if ate_low is not None:
                y_candidates.append(float(ate_low))
            if ate_high is not None:
                y_candidates.append(float(ate_high))
            if not math.isnan(new_ate_low):
                y_candidates.append(float(new_ate_low))
            if not math.isnan(new_ate_high):
                y_candidates.append(float(new_ate_high))

            y_min = min(y_candidates)
            y_max = max(y_candidates)
            y_span = y_max - y_min
            pad = (0.06 * y_span) if y_span > 0 else (abs(y_min) * 0.06 + 1.0)
            ax1.set_ylim(y_min - pad, y_max + pad)

            # Keep the top subplot free of regression text (as requested).
            note_top = f"N={n_points}  median(CV-ATE)={median_delta:.4g}  max(CV-ATE)={max_delta:.4g}"
            if not math.isnan(residual_std):
                note_top += f"  σ(delta-res)={residual_std:.4g}  GB={guardband_ate:.4g}"
            ax1.text(
                0.015,
                0.02,
                note_top,
                transform=ax1.transAxes,
                fontsize=11,
                va="bottom",
                ha="left",
                bbox={"facecolor": "white", "alpha": 0.85, "edgecolor": "none", "pad": 3.0},
            )

            ax1.legend(fontsize=11, framealpha=0.92)

            # Regression subplot: CV vs ATE scattering + offset-only line (slope=1)
            ax2.scatter(
                g_plot["CV"],
                g_plot["ATE"],
                s=40,
                alpha=0.95,
                marker="o",
                linewidths=1.1,
                facecolors="none",
                edgecolors="tab:blue",
                zorder=3,
                label="ATE raw (scatter)",
            )
            x_min = float(g_plot["CV"].min())
            x_max = float(g_plot["CV"].max())
            x_span = x_max - x_min
            x_pad = (0.06 * x_span) if x_span > 0 else (abs(x_min) * 0.06 + 1.0)
            x_line = pd.Series([x_min, x_max])
            y_line = x_line - median_delta
            ax2.plot(
                x_line,
                y_line,
                linewidth=2.6,
                linestyle="--",
                color="tab:red",
                zorder=1,
                label=f"Offset model: ATE = CV - medianΔ",
            )

            ax2.set_xlabel(f"CV{(' ['+unit+']') if unit else ''}", fontsize=12)
            ax2.set_ylabel(f"ATE{(' ['+unit+']') if unit else ''}", fontsize=12)
            ax2.tick_params(axis="both", labelsize=11)
            ax2.grid(True, alpha=0.25)
            ax2.set_xlim(x_min - x_pad, x_max + x_pad)

            y2_candidates = [
                float(g_plot["ATE"].min()),
                float(g_plot["ATE"].max()),
                float(y_line.min()),
                float(y_line.max()),
            ]
            y2_min = min(y2_candidates)
            y2_max = max(y2_candidates)
            y2_span = y2_max - y2_min
            y2_pad = (0.06 * y2_span) if y2_span > 0 else (abs(y2_min) * 0.06 + 1.0)
            ax2.set_ylim(y2_min - y2_pad, y2_max + y2_pad)

            note_reg = f"N={n_points}  R²={r2:.3f}  medianΔ={median_delta:.4g}  maxΔ={max_delta:.4g}"
            ax2.text(
                0.015,
                0.02,
                note_reg,
                transform=ax2.transAxes,
                fontsize=11,
                va="bottom",
                ha="left",
                bbox={"facecolor": "white", "alpha": 0.85, "edgecolor": "none", "pad": 3.0},
            )
            ax2.legend(fontsize=10, framealpha=0.92)

            fig.tight_layout(rect=[0, 0, 1, 0.96])

            fname = _safe_slug(title) + ".png"
            fig.savefig(local_plots_dir / fname, dpi=PLOT_DPI, bbox_inches="tight")
            plt.close(fig)

    # Run
    if sheets_to_run:
        for sh in sheets_to_run:
            merged_sheet = run_one_same_sheet(sh)
            correlate_merged(merged_sheet, sh)
    else:
        merged_2 = run_one_two_sheets()
        correlate_merged(merged_2, f"{CV_SHEET}__{ATE_SHEET}")

    if not all_factors_rows:
        raise SystemExit(
            "No groups produced results. Check MIN_POINTS_PER_GROUP, grouping columns, and whether CV/ATE rows merge correctly."
        )

    factors_df = pd.DataFrame(all_factors_rows)
    correlated_df = pd.DataFrame(all_data_rows)

    # Sign-off summary (TXPA): a compact table per (LUT, freq, temp) showing
    # - POR baseline median
    # - per-split median of (Delta - POR)
    # - per-split N
    # - overall spread metrics of (Delta - POR)
    signoff_summary_df = pd.DataFrame()
    required = {
        "LUT value",
        "Frequency_GHz",
        "Temperature",
        "DoE split",
        "Delta(CV-ATE)-POR",
        "POR_MedianDelta(CV-ATE)",
    }
    if required.issubset(set(correlated_df.columns)):
        base_cols = ["LUT value", "Frequency_GHz", "Temperature"]

        tmp = correlated_df.copy()
        tmp["LUT value"] = pd.to_numeric(tmp["LUT value"], errors="coerce").astype("Int64")
        tmp["Frequency_GHz"] = pd.to_numeric(tmp["Frequency_GHz"], errors="coerce")
        tmp["Temperature"] = tmp["Temperature"].astype(str).str.strip().str.upper()
        tmp["DoE split"] = tmp["DoE split"].astype(str).str.strip().str.upper()
        tmp["Delta(CV-ATE)-POR"] = pd.to_numeric(tmp["Delta(CV-ATE)-POR"], errors="coerce")
        tmp["POR_MedianDelta(CV-ATE)"] = pd.to_numeric(tmp["POR_MedianDelta(CV-ATE)"], errors="coerce")

        # Per-split medians and counts of POR-normalized delta
        by_split = (
            tmp.dropna(subset=["Delta(CV-ATE)-POR"])  # keep numeric only
            .groupby(base_cols + ["DoE split"], dropna=False)["Delta(CV-ATE)-POR"]
            .agg(N="count", Median="median")
            .reset_index()
        )

        median_wide = by_split.pivot(index=base_cols, columns="DoE split", values="Median")
        n_wide = by_split.pivot(index=base_cols, columns="DoE split", values="N")

        median_wide = median_wide.add_prefix("MedianDeltaMinusPOR_").reset_index()
        n_wide = n_wide.add_prefix("N_").reset_index()

        # POR baseline per base condition (should be constant); take median for safety
        por_baseline = (
            tmp.dropna(subset=["POR_MedianDelta(CV-ATE)"])
            .groupby(base_cols, dropna=False)["POR_MedianDelta(CV-ATE)"]
            .median()
            .rename("POR_MedianDelta(CV-ATE)")
            .reset_index()
        )

        # Overall spread of normalized delta (all splits mixed)
        def _q(p: float):
            return lambda s: float(s.quantile(p)) if len(s) else math.nan

        spread = (
            tmp.dropna(subset=["Delta(CV-ATE)-POR"])
            .groupby(base_cols, dropna=False)["Delta(CV-ATE)-POR"]
            .agg(
                N_All="count",
                Std_All="std",
                P05_All=_q(0.05),
                P95_All=_q(0.95),
                MaxAbs_All=lambda s: float(s.abs().max()) if len(s) else math.nan,
            )
            .reset_index()
        )

        signoff_summary_df = por_baseline.merge(median_wide, on=base_cols, how="left").merge(
            n_wide, on=base_cols, how="left"
        ).merge(spread, on=base_cols, how="left")

        # Convenience KPI: maximum absolute split-median deviation (excluding POR).
        median_cols = [c for c in signoff_summary_df.columns if c.startswith("MedianDeltaMinusPOR_")]
        nonpor_median_cols = [c for c in median_cols if c.upper() != "MEDIANDELTAMINUSPOR_POR"]
        if nonpor_median_cols:
            signoff_summary_df["MaxAbsMedian_NonPOR"] = (
                signoff_summary_df[nonpor_median_cols].abs().max(axis=1, skipna=True)
            )
        else:
            signoff_summary_df["MaxAbsMedian_NonPOR"] = math.nan

    # Put Test Name directly after the primary identifier (Test Number for TXLO, LUT value for TXPA)
    if "Test Number" in factors_df.columns:
        factors_df = _insert_column_after(factors_df, "Test Number", "Test Name")
    elif "LUT value" in factors_df.columns:
        factors_df = _insert_column_after(factors_df, "LUT value", "Test Name")

    if "Test Number" in correlated_df.columns:
        correlated_df = _insert_column_after(correlated_df, "Test Number", "Test Name")
    elif "LUT value" in correlated_df.columns:
        correlated_df = _insert_column_after(correlated_df, "LUT value", "Test Name")

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
            factors_df.to_excel(writer, index=False, sheet_name="Correlation_Factors")
            correlated_df.to_excel(writer, index=False, sheet_name="Correlated_Data")
            if not signoff_summary_df.empty:
                signoff_summary_df.to_excel(writer, index=False, sheet_name="Signoff_Summary")
    except PermissionError:
        # Most common cause: the workbook is open in Excel.
        alt_path = output_xlsx.with_name(output_xlsx.stem + "_new.xlsx")
        with pd.ExcelWriter(alt_path, engine="openpyxl") as writer:
            factors_df.to_excel(writer, index=False, sheet_name="Correlation_Factors")
            correlated_df.to_excel(writer, index=False, sheet_name="Correlated_Data")
            if not signoff_summary_df.empty:
                signoff_summary_df.to_excel(writer, index=False, sheet_name="Signoff_Summary")
        output_xlsx = alt_path

    print(f"Wrote factors: {len(factors_df)} groups")
    print(f"Wrote data: {len(correlated_df)} rows")
    print(f"Output Excel: {output_xlsx}")
    print(f"Plots folder: {plots_dir}")
