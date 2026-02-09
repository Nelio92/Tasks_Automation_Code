"""\
TXPA/TXLO correlated power CDF plots on PROD raw-data CSVs (flat execution).

Goal
----
Using existing correlation factors and correlated limits, compute:

    ATE_Correlated = ATE_data + corr_factor

and generate CDF plots that overlay:
  - uncorrelated (raw ATE) distribution
  - correlated distribution
  - correlated limits valid for the correlated data

Plots
-----
- FE: 3 subplots (S1/S2/S3)
- BE: 2 subplots (B1/B2)

Data
----
- Raw PROD CSVs in repo folder: PROD_Data
- Parsing relies on:
    Tasks_Automation_Code/Reports/TX_Supply_Compensation_PROD/analyze_tx_supply_compensation_scenarios.py

Correlation files (repo root)
-----------------------------
- CV_ATE_Correlation_TXLO_Power_FE.xlsx (TXLO factors + FE limits)
- CV_ATE_Correlation_TXPA_Power_FE.xlsx (TXPA factors + FE limits)
- CV_ATE_Correlation_TXPA_Power_BE.xlsx (TXPA factors + BE limits)

Insertion mapping from filename
-------------------------------
- S11P → S1 (FE hot 135°C)
- S21P → S2 (FE cold -40°C)
- S31P → S3 (FE ambient 25°C)
- B11P → B1 (BE hot 135°C)
- B21P → B2 (BE ambient 25°C)

Notes
-----
- No CLI by design; edit constants in the CONFIG block.
- Correlation spreadsheets in this repo only contain VMIN, so voltage-corner
  matching is currently fixed to VMIN.

"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
import math
from pathlib import Path
import re
import sys
from typing import Iterable
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd


# -----------------------------------------------------------------------------
# CONFIG (edit as needed)
# -----------------------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parents[2]
REPORT_DIR = Path(__file__).resolve().parent

INPUT_FOLDER = REPO_ROOT / "PROD_Data"
INPUT_GLOB = "*.csv"
MAX_FILES: int | None = None

TXLO_CORR_XLSX_FE = REPO_ROOT / "CV_ATE_Correlation_TXLO_Power_FE.xlsx"
TXPA_CORR_XLSX_FE = REPO_ROOT / "CV_ATE_Correlation_TXPA_Power_FE.xlsx"
TXPA_CORR_XLSX_BE = REPO_ROOT / "CV_ATE_Correlation_TXPA_Power_BE.xlsx"

PLOT_DPI = 170

# Outlier filtering (MAD)
# A point x is considered an outlier if |x - median(x)| > MAD_MULTIPLIER * MAD,
# where MAD = median(|x - median(x)|).
MAD_MULTIPLIER: float = 5.0

# Extrema highlighting (specific request)
EXTREMA_FREQS: set[int] = {81, 77, 76}
EXTREMA_TXLO_LO_IDAC: int = 112
EXTREMA_TXPA_LUT: int = 255

# If True, extrema decisions use the correlated series median when a factor exists;
# otherwise falls back to raw.
EXTREMA_USE_CORRELATED_MEDIAN: bool = True

HIGHLIGHT_MIN_COLOR = "#2ca02c"  # green
HIGHLIGHT_MAX_COLOR = "#9467bd"  # purple

# Plot styling
RAW_COLOR = "#1f77b4"  # matplotlib default blue
CORR_COLOR = "#ff7f0e"  # matplotlib default orange

TX_CHANNEL_MARKERS = {
    "TX1": "o",
    "TX2": "s",
    "TX3": "D",
    "TX4": "^",
    "TX5": "v",
    "TX6": "<",
    "TX7": ">",
    "TX8": "P",
}


# Make repo root importable even when running this script from another CWD.
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Tasks_Automation_Code.Reports.TX_Supply_Compensation_PROD.analyze_tx_supply_compensation_scenarios import (  # noqa: E402
    _read_prod_csv_chip_matrix,
)


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

_INSERTION_ORDER_FE = ["S1", "S2", "S3"]
_INSERTION_ORDER_BE = ["B1", "B2"]

_TEMP_BY_INS: dict[str, int] = {"S1": 135, "S2": -40, "S3": 25, "B1": 135, "B2": 25}
_INS_TYPE_BY_INS: dict[str, str] = {"S1": "FE", "S2": "FE", "S3": "FE", "B1": "BE", "B2": "BE"}


@dataclass(frozen=True)
class InsertionInfo:
    insertion: str  # S1/S2/S3/B1/B2
    insertion_type: str  # FE/BE
    temperature_c: int


@dataclass(frozen=True)
class TestMeta:
    test_number: int
    test_name: str
    kind: str  # TXLO or TXPA
    frequency_ghz: int | None
    lut_value: int | None
    lo_idac: int | None
    pa_channel: str | None


@dataclass(frozen=True)
class PlotGroup:
    kind: str  # TXLO/TXPA
    insertion_type: str  # FE/BE
    title: str
    test_numbers: list[int]
    lut_value: int | None = None
    frequency_ghz: int | None = None


_TX_PREFIX_RE = re.compile(r"^(?P<blk>TXLO|TXPA|TXPB|TXPC)_(?P<freq>\d+)", flags=re.IGNORECASE)
_FWLU_RE = re.compile(r"FwLu(?P<lut>\d{1,3})(?!\d)", flags=re.IGNORECASE)
_POWCO_RE = re.compile(r"PowCo(?P<idac>\d{1,3})(?!\d)", flags=re.IGNORECASE)
_TX_CH_RE = re.compile(r"Tx(?P<ch>[1-8])(?!\d)", flags=re.IGNORECASE)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]
    return df


def _read_prod_csv_test_name_map(path: Path, *, tests: Iterable[int] | None = None) -> dict[int, str]:
    wanted: set[int] | None = None
    if tests is not None:
        wanted = {int(t) for t in tests}

    try:
        with path.open("r", encoding="latin1", errors="ignore") as f:
            header_line = f.readline().rstrip("\n\r")
            test_name_line = f.readline().rstrip("\n\r")
    except Exception:
        return {}

    header_cells = header_line.split(";")
    test_name_cells = test_name_line.split(";")
    if len(test_name_cells) < len(header_cells):
        test_name_cells = test_name_cells + [""] * (len(header_cells) - len(test_name_cells))

    out: dict[int, str] = {}
    for h, n in zip(header_cells, test_name_cells, strict=False):
        hs = str(h).strip()
        if not hs.isdigit():
            continue
        t = int(hs)
        if wanted is not None and t not in wanted:
            continue
        out[t] = str(n).strip()

    return out


def _insertion_from_filename(path: Path) -> InsertionInfo | None:
    name_upper = path.name.upper()

    if "S11P" in name_upper:
        return InsertionInfo(insertion="S1", insertion_type="FE", temperature_c=135)
    if "S21P" in name_upper:
        return InsertionInfo(insertion="S2", insertion_type="FE", temperature_c=-40)
    if "S31P" in name_upper:
        return InsertionInfo(insertion="S3", insertion_type="FE", temperature_c=25)
    if "B11P" in name_upper:
        return InsertionInfo(insertion="B1", insertion_type="BE", temperature_c=135)
    if "B21P" in name_upper:
        return InsertionInfo(insertion="B2", insertion_type="BE", temperature_c=25)

    # Unknown / unsupported in this report.
    return None


def _parse_test_meta(test_number: int, test_name: str) -> TestMeta | None:
    tn = int(test_number)
    name = str(test_name or "").strip()
    m = _TX_PREFIX_RE.match(name)
    if not m:
        return None

    blk = m.group("blk").upper()
    freq = None
    try:
        freq = int(m.group("freq"))
    except Exception:
        freq = None

    lut = None
    m_lut = _FWLU_RE.search(name)
    if m_lut:
        try:
            lut = int(m_lut.group("lut"))
        except Exception:
            lut = None

    lo_idac = None
    m_idac = _POWCO_RE.search(name)
    if m_idac:
        try:
            lo_idac = int(m_idac.group("idac"))
        except Exception:
            lo_idac = None

    pa_ch = None
    m_ch = _TX_CH_RE.search(name)
    if m_ch:
        pa_ch = f"TX{m_ch.group('ch')}"

    if blk == "TXLO":
        kind = "TXLO"
    else:
        kind = "TXPA"  # treat TXPA/TXPB/TXPC as the same PA-power family

    return TestMeta(
        test_number=tn,
        test_name=name,
        kind=kind,
        frequency_ghz=freq,
        lut_value=lut,
        lo_idac=lo_idac,
        pa_channel=pa_ch,
    )


def _parse_test_number_cell(cell) -> list[int]:
    if cell is None or (isinstance(cell, float) and math.isnan(cell)):
        return []

    if isinstance(cell, (int, np.integer)):
        return [int(cell)]

    # Excel often returns floats for integer-like columns.
    if isinstance(cell, float) and float(cell).is_integer():
        return [int(cell)]

    s = str(cell).strip()
    if not s:
        return []

    m = re.fullmatch(r"(?P<a>\d+)\s*-\s*(?P<b>\d+)", s)
    if m:
        a = int(m.group("a"))
        b = int(m.group("b"))
        if a > b:
            a, b = b, a
        return list(range(a, b + 1))

    if s.isdigit():
        return [int(s)]

    return []


def _cdf_xy(values: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    v = np.asarray(values, dtype=float)
    v = v[np.isfinite(v)]
    if v.size == 0:
        return np.array([], dtype=float), np.array([], dtype=float)
    x = np.sort(v)
    y = (np.arange(1, len(x) + 1, dtype=float)) / float(len(x))
    return x, y


def _median(values: np.ndarray) -> float:
    v = np.asarray(values, dtype=float)
    v = v[np.isfinite(v)]
    if v.size == 0:
        return float("nan")
    return float(np.median(v))


def _filter_outliers_mad(values: np.ndarray, *, mad_multiplier: float) -> tuple[np.ndarray, int]:
    """Return (filtered_values, n_removed) using MAD thresholding.

    Keeps NaNs/infs out automatically.
    If MAD is 0 (or too few points), returns values unchanged.
    """

    v = np.asarray(values, dtype=float)
    v = v[np.isfinite(v)]
    if v.size < 5:
        return v, 0

    med = float(np.median(v))
    abs_dev = np.abs(v - med)
    mad = float(np.median(abs_dev))
    if not np.isfinite(mad) or mad <= 0.0:
        return v, 0

    thr = float(mad_multiplier) * mad
    keep = abs_dev <= thr
    out = v[keep]
    return out, int(v.size - out.size)


def _format_test_names_for_title(test_numbers: list[int], test_name_by_num: dict[int, str], *, max_items: int = 6, max_chars: int = 220) -> str:
    parts: list[str] = []
    for tn in list(test_numbers)[: int(max_items)]:
        name = str(test_name_by_num.get(int(tn), "") or "").strip()
        if name:
            parts.append(f"{int(tn)}: {name}")
        else:
            parts.append(str(int(tn)))

    if len(test_numbers) > int(max_items):
        parts.append(f"... (+{len(test_numbers) - int(max_items)} more)")

    s = "; ".join(parts).strip()
    if len(s) > int(max_chars):
        s = s[: int(max_chars) - 3].rstrip() + "..."
    return s


# -----------------------------------------------------------------------------
# Load correlation factors / limits
# -----------------------------------------------------------------------------


def _load_txlo_factors_and_limits() -> tuple[dict[tuple[int, int, int, int], float], dict[tuple[int, int, int], dict[str, tuple[float, float]]]]:
    """Return (factor_lookup, limits_lookup).

    factor_lookup key: (test_number, freq_ghz, temperature_c, lo_idac)
    limits_lookup key: (test_number, freq_ghz, lo_idac) -> {S1/S2/S3: (low, high)}
    """

    fac = _normalize_columns(pd.read_excel(TXLO_CORR_XLSX_FE, sheet_name="Correlation_Factors"))
    lim = _normalize_columns(pd.read_excel(TXLO_CORR_XLSX_FE, sheet_name="Correlation_Limits"))

    fac["Test Number"] = pd.to_numeric(fac["Test Number"], errors="coerce").astype("Int64")
    fac["Frequency_GHz"] = pd.to_numeric(fac["Frequency_GHz"], errors="coerce").astype("Int64")
    fac["Temperature"] = pd.to_numeric(fac["Temperature"], errors="coerce").astype("Int64")
    fac["LO IDAC"] = pd.to_numeric(fac["LO IDAC"], errors="coerce").astype("Int64")

    factor_lookup: dict[tuple[int, int, int, int], float] = {}
    for _, r in fac.iterrows():
        tn = r.get("Test Number")
        fr = r.get("Frequency_GHz")
        te = r.get("Temperature")
        lo = r.get("LO IDAC")
        delta = r.get("MedianDelta(CV-ATE)")
        if pd.isna(tn) or pd.isna(fr) or pd.isna(te) or pd.isna(lo):
            continue
        if pd.isna(delta):
            continue
        factor_lookup[(int(tn), int(fr), int(te), int(lo))] = float(delta)

    # Limits
    lim["Test Number"] = pd.to_numeric(lim["Test Number"], errors="coerce").astype("Int64")
    lim["Frequency_GHz"] = pd.to_numeric(lim["Frequency_GHz"], errors="coerce").astype("Int64")
    lim["LO IDAC"] = pd.to_numeric(lim["LO IDAC"], errors="coerce").astype("Int64")

    limits_lookup: dict[tuple[int, int, int], dict[str, tuple[float, float]]] = {}

    for _, r in lim.iterrows():
        tn = r.get("Test Number")
        fr = r.get("Frequency_GHz")
        lo = r.get("LO IDAC")
        if pd.isna(tn) or pd.isna(fr) or pd.isna(lo):
            continue

        out: dict[str, tuple[float, float]] = {}
        for ins in _INSERTION_ORDER_FE:
            low_col = f"Corr_Low ({ins})"
            high_col = f"Corr_High ({ins})"
            low = pd.to_numeric(r.get(low_col), errors="coerce")
            high = pd.to_numeric(r.get(high_col), errors="coerce")
            if pd.notna(low) and pd.notna(high):
                out[ins] = (float(low), float(high))

        limits_lookup[(int(tn), int(fr), int(lo))] = out

    return factor_lookup, limits_lookup


def _load_txpa_factors_and_limits_fe() -> tuple[
    dict[tuple[int, int, int, str], float],
    dict[tuple[int, int, int, str], dict[str, tuple[float, float, str]]],
]:
    """Return (factor_lookup, limits_lookup).

    factor_lookup key: (lut_value, freq_ghz, temperature_c, pa_channel_or_ALL)
    limits_lookup key: (test_number, freq_ghz, lut_value, pa_channel_or_ALL) -> {S1/S2/S3: (low, high, unit)}

    Note: limits are per test number (expanded from ranges in the spreadsheet).
    """

    fac = _normalize_columns(pd.read_excel(TXPA_CORR_XLSX_FE, sheet_name="Correlation_Factors"))
    lim = _normalize_columns(pd.read_excel(TXPA_CORR_XLSX_FE, sheet_name="Correlation_Limits"))

    fac["LUT value"] = pd.to_numeric(fac["LUT value"], errors="coerce").astype("Int64")
    fac["Frequency_GHz"] = pd.to_numeric(fac["Frequency_GHz"], errors="coerce").astype("Int64")
    fac["Temperature"] = pd.to_numeric(fac["Temperature"], errors="coerce").astype("Int64")
    fac["PA Channel"] = fac.get("PA Channel", "ALL").astype(str).str.strip().replace({"": "ALL"})

    factor_lookup: dict[tuple[int, int, int, str], float] = {}
    for _, r in fac.iterrows():
        lut = r.get("LUT value")
        fr = r.get("Frequency_GHz")
        te = r.get("Temperature")
        pa = str(r.get("PA Channel") or "ALL").strip() or "ALL"
        delta = r.get("MedianDelta(CV-ATE)")
        if pd.isna(lut) or pd.isna(fr) or pd.isna(te) or pd.isna(delta):
            continue
        factor_lookup[(int(lut), int(fr), int(te), pa.upper())] = float(delta)

    # Limits
    lim["LUT value"] = pd.to_numeric(lim["LUT value"], errors="coerce").astype("Int64")
    lim["Frequency_GHz"] = pd.to_numeric(lim["Frequency_GHz"], errors="coerce").astype("Int64")
    lim["PA Channel"] = lim.get("PA Channel", "ALL").astype(str).str.strip().replace({"": "ALL"})

    limits_lookup: dict[tuple[int, int, int, str], dict[str, tuple[float, float, str]]] = {}
    for _, r in lim.iterrows():
        lut = r.get("LUT value")
        fr = r.get("Frequency_GHz")
        if pd.isna(lut) or pd.isna(fr):
            continue

        pa = str(r.get("PA Channel") or "ALL").strip() or "ALL"
        unit = str(r.get("Unit") or "").strip()

        tnums = _parse_test_number_cell(r.get("Test Number"))
        if not tnums:
            continue

        out: dict[str, tuple[float, float, str]] = {}
        for ins in _INSERTION_ORDER_FE:
            low_col = f"Corr_Low ({ins})"
            high_col = f"Corr_High ({ins})"
            low = pd.to_numeric(r.get(low_col), errors="coerce")
            high = pd.to_numeric(r.get(high_col), errors="coerce")
            if pd.notna(low) and pd.notna(high):
                out[ins] = (float(low), float(high), unit)

        for tn in tnums:
            limits_lookup[(int(tn), int(fr), int(lut), pa.upper())] = out

    return factor_lookup, limits_lookup


def _load_txpa_factors_and_limits_be() -> tuple[
    dict[tuple[int, int, int, str], float],
    dict[tuple[int, int, int, str], dict[str, tuple[float, float, str]]],
]:
    """Return (factor_lookup, limits_lookup) for BE.

    factor_lookup key: (lut_value, freq_ghz, temperature_c, pa_channel_or_ALL)
    limits_lookup key: (test_number, freq_ghz, lut_value, pa_channel_or_ALL) -> {B1/B2: (low, high, unit)}
    """

    # Limits are small; load first to know which LUT/freq combos we care about.
    lim = _normalize_columns(pd.read_excel(TXPA_CORR_XLSX_BE, sheet_name="Correlation_Limits"))
    lim["LUT value"] = pd.to_numeric(lim["LUT value"], errors="coerce").astype("Int64")
    lim["Frequency_GHz"] = pd.to_numeric(lim["Frequency_GHz"], errors="coerce").astype("Int64")

    needed_luts = sorted({int(v) for v in lim["LUT value"].dropna().unique()})
    needed_freqs = sorted({int(v) for v in lim["Frequency_GHz"].dropna().unique()})

    # Factors sheet is large; read only the necessary columns.
    fac = _normalize_columns(
        pd.read_excel(
            TXPA_CORR_XLSX_BE,
            sheet_name="Correlation_Factors",
            usecols=["LUT value", "Frequency_GHz", "Temperature", "PA Channel", "MedianDelta(CV-ATE)", "Voltage corner"],
            engine="openpyxl",
        )
    )

    fac["LUT value"] = pd.to_numeric(fac["LUT value"], errors="coerce").astype("Int64")
    fac["Frequency_GHz"] = pd.to_numeric(fac["Frequency_GHz"], errors="coerce").astype("Int64")
    fac["Temperature"] = pd.to_numeric(fac["Temperature"], errors="coerce").astype("Int64")
    fac["PA Channel"] = fac.get("PA Channel", "ALL").astype(str).str.strip().replace({"": "ALL"})
    fac["Voltage corner"] = fac.get("Voltage corner", "VMIN").astype(str).str.strip().replace({"": "VMIN"})

    # Filter aggressively; spreadsheets in this repo are VMIN only.
    fac = fac.loc[
        fac["Voltage corner"].astype(str).str.upper().eq("VMIN")
        & fac["LUT value"].isin(needed_luts)
        & fac["Frequency_GHz"].isin(needed_freqs)
    ].copy()

    # Deduplicate by key (some sheets contain repeated rows).
    fac = fac.dropna(subset=["LUT value", "Frequency_GHz", "Temperature", "MedianDelta(CV-ATE)"])

    factor_lookup: dict[tuple[int, int, int, str], float] = {}
    if not fac.empty:
        grp = fac.groupby(["LUT value", "Frequency_GHz", "Temperature", "PA Channel"], dropna=True)[
            "MedianDelta(CV-ATE)"
        ].median()
        for (lut, fr, te, pa), delta in grp.items():
            try:
                factor_lookup[(int(lut), int(fr), int(te), str(pa).strip().upper() or "ALL")] = float(delta)
            except Exception:
                continue

    # Limits lookup (expanded to per test number)
    limits_lookup: dict[tuple[int, int, int, str], dict[str, tuple[float, float, str]]] = {}
    for _, r in lim.iterrows():
        lut = r.get("LUT value")
        fr = r.get("Frequency_GHz")
        if pd.isna(lut) or pd.isna(fr):
            continue

        pa = str(r.get("PA Channel") or "ALL").strip() or "ALL"
        unit = str(r.get("Unit") or "").strip()

        tnums = _parse_test_number_cell(r.get("Test Number"))
        if not tnums:
            continue

        out: dict[str, tuple[float, float, str]] = {}
        for ins in _INSERTION_ORDER_BE:
            low_col = f"Corr_Low ({ins})"
            high_col = f"Corr_High ({ins})"
            low = pd.to_numeric(r.get(low_col), errors="coerce")
            high = pd.to_numeric(r.get(high_col), errors="coerce")
            if pd.notna(low) and pd.notna(high):
                out[ins] = (float(low), float(high), unit)

        for tn in tnums:
            limits_lookup[(int(tn), int(fr), int(lut), pa.upper())] = out

    return factor_lookup, limits_lookup


def _lookup_txpa_factor(
    factor_lookup: dict[tuple[int, int, int, str], float],
    *,
    lut: int,
    freq: int,
    temp_c: int,
    pa_channel: str | None,
) -> float | None:
    pa = (pa_channel or "ALL").strip().upper() or "ALL"
    key_exact = (int(lut), int(freq), int(temp_c), pa)
    key_all = (int(lut), int(freq), int(temp_c), "ALL")
    if key_exact in factor_lookup:
        return float(factor_lookup[key_exact])
    if key_all in factor_lookup:
        return float(factor_lookup[key_all])
    # fallback: any PA channel for this (lut,freq,temp)
    for (l, f, t, p), v in factor_lookup.items():
        if l == int(lut) and f == int(freq) and t == int(temp_c):
            return float(v)
    return None


def _lookup_txpa_limits(
    limits_lookup: dict[tuple[int, int, int, str], dict[str, tuple[float, float, str]]],
    *,
    test_number: int,
    lut: int,
    freq: int,
    pa_channel: str | None,
) -> dict[str, tuple[float, float, str]]:
    pa = (pa_channel or "ALL").strip().upper() or "ALL"
    key_exact = (int(test_number), int(freq), int(lut), pa)
    key_all = (int(test_number), int(freq), int(lut), "ALL")
    if key_exact in limits_lookup:
        return limits_lookup[key_exact]
    if key_all in limits_lookup:
        return limits_lookup[key_all]
    return {}


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------


def main() -> int:
    run_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = REPORT_DIR / f"output_{run_tag}__txpa_txlo_correlated_power"
    plots_dir = out_dir / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)

    # Load correlation lookups
    txlo_factor, txlo_limits = _load_txlo_factors_and_limits()
    txpa_factor_fe, txpa_limits_fe = _load_txpa_factors_and_limits_fe()
    txpa_factor_be, txpa_limits_be = _load_txpa_factors_and_limits_be()

    # Read limit sheets (also used to build plot grouping for TXPA ranges)
    txlo_lim_df = _normalize_columns(pd.read_excel(TXLO_CORR_XLSX_FE, sheet_name="Correlation_Limits"))
    txpa_fe_lim_df = _normalize_columns(pd.read_excel(TXPA_CORR_XLSX_FE, sheet_name="Correlation_Limits"))
    txpa_be_lim_df = _normalize_columns(pd.read_excel(TXPA_CORR_XLSX_BE, sheet_name="Correlation_Limits"))

    # Build tests_needed from limits (TXLO + TXPA FE + TXPA BE)
    tests_needed: set[int] = set(pd.to_numeric(txlo_lim_df["Test Number"], errors="coerce").dropna().astype(int).tolist())
    for cell in txpa_fe_lim_df.get("Test Number", []):
        tests_needed.update(_parse_test_number_cell(cell))
    for cell in txpa_be_lim_df.get("Test Number", []):
        tests_needed.update(_parse_test_number_cell(cell))

    files = sorted(INPUT_FOLDER.glob(INPUT_GLOB))
    if MAX_FILES is not None:
        files = files[: int(MAX_FILES)]

    if not files:
        raise SystemExit(f"No input files found in: {INPUT_FOLDER}")

    # Aggregate values across files per insertion & test
    values_by_ins_test: dict[tuple[str, int], list[np.ndarray]] = {}
    test_name_by_num: dict[int, str] = {}

    missing_factor_rows: list[dict] = []
    missing_limit_rows: list[dict] = []
    outlier_rows: list[dict] = []
    corr_stats_rows: list[dict] = []

    for p in files:
        ins = _insertion_from_filename(p)
        if ins is None:
            continue

        chip_matrix, info = _read_prod_csv_chip_matrix(p, tests_needed=tests_needed)
        if info.get("error"):
            continue
        if chip_matrix.empty:
            continue

        name_map = _read_prod_csv_test_name_map(p, tests=tests_needed)

        present = [t for t in tests_needed if t in chip_matrix.columns]
        for tn in present:
            test_name = name_map.get(int(tn), "")
            if int(tn) not in test_name_by_num:
                test_name_by_num[int(tn)] = test_name

            arr = pd.to_numeric(chip_matrix[int(tn)], errors="coerce").dropna().to_numpy(dtype=float)
            if arr.size == 0:
                continue

            values_by_ins_test.setdefault((ins.insertion, int(tn)), []).append(arr)

    # Decide which tests we should plot by looking at collected values
    all_test_numbers = sorted({tn for (_, tn) in values_by_ins_test.keys()})

    # Pre-parse metadata for all tests
    meta_by_tn: dict[int, TestMeta] = {}
    for tn in all_test_numbers:
        meta = _parse_test_meta(tn, test_name_by_num.get(tn, ""))
        if meta is not None:
            meta_by_tn[tn] = meta

    # Build plot groups
    groups: list[PlotGroup] = []

    # TXLO: per test number (FE only)
    for tn in sorted(set(pd.to_numeric(txlo_lim_df["Test Number"], errors="coerce").dropna().astype(int).tolist())):
        if tn in meta_by_tn and meta_by_tn[tn].kind == "TXLO":
            groups.append(PlotGroup(kind="TXLO", insertion_type="FE", title=f"TXLO test {tn}", test_numbers=[tn]))

    def _add_txpa_groups_from_limits(lim_df: pd.DataFrame, insertion_type: str) -> None:
        if lim_df is None or lim_df.empty:
            return
        df = lim_df.copy()
        df["LUT value"] = pd.to_numeric(df.get("LUT value"), errors="coerce").astype("Int64")
        df["Frequency_GHz"] = pd.to_numeric(df.get("Frequency_GHz"), errors="coerce").astype("Int64")

        for _, r in df.iterrows():
            lut = r.get("LUT value")
            fr = r.get("Frequency_GHz")
            if pd.isna(lut) or pd.isna(fr):
                continue
            lut_i = int(lut)
            fr_i = int(fr)

            tnums = _parse_test_number_cell(r.get("Test Number"))
            if not tnums:
                continue

            # Group ranges only for LUT 12..244. LUT255 must remain per-TX channel (per test).
            if 12 <= lut_i <= 244 and len(tnums) > 1:
                groups.append(
                    PlotGroup(
                        kind="TXPA",
                        insertion_type=insertion_type,
                        title=f"TXPA LUT{lut_i} @ {fr_i}GHz tests {min(tnums)}-{max(tnums)}",
                        test_numbers=sorted({int(t) for t in tnums}),
                        lut_value=lut_i,
                        frequency_ghz=fr_i,
                    )
                )
            else:
                for tn in tnums:
                    groups.append(
                        PlotGroup(
                            kind="TXPA",
                            insertion_type=insertion_type,
                            title=f"TXPA LUT{lut_i} @ {fr_i}GHz test {int(tn)}",
                            test_numbers=[int(tn)],
                            lut_value=lut_i,
                            frequency_ghz=fr_i,
                        )
                    )

    _add_txpa_groups_from_limits(txpa_fe_lim_df, "FE")
    _add_txpa_groups_from_limits(txpa_be_lim_df, "BE")

    # Only keep groups that have at least one test present in collected data
    present_tests = set(all_test_numbers)
    groups = [g for g in groups if any(t in present_tests for t in g.test_numbers)]

    # De-duplicate identical groups (limits sheets can contain repeated rows)
    seen = set()
    uniq: list[PlotGroup] = []
    for g in groups:
        key = (g.kind, g.insertion_type, tuple(g.test_numbers), g.lut_value, g.frequency_ghz)
        if key in seen:
            continue
        seen.add(key)
        uniq.append(g)
    groups = uniq

    # ---------------------------------------------------------------------
    # Extrema analysis (min/max median across freq+temp condition-sets)
    # ---------------------------------------------------------------------
    # Keys:
    #   TXLO: ("TXLO", lo_idac)
    #   TXPA: ("TXPA", insertion_type, pa_channel)
    extrema_candidates: dict[tuple, list[dict]] = {}

    for (ins_label, tn), arrays in values_by_ins_test.items():
        meta = meta_by_tn.get(int(tn))
        if meta is None or meta.frequency_ghz is None:
            continue
        freq = int(meta.frequency_ghz)
        if freq not in EXTREMA_FREQS:
            continue
        if ins_label not in _TEMP_BY_INS:
            continue
        temp = int(_TEMP_BY_INS[ins_label])

        raw = np.concatenate(arrays) if arrays else np.array([], dtype=float)
        raw, _ = _filter_outliers_mad(raw, mad_multiplier=MAD_MULTIPLIER)
        if raw.size == 0:
            continue

        if meta.kind == "TXLO" and meta.lo_idac == int(EXTREMA_TXLO_LO_IDAC) and ins_label in _INSERTION_ORDER_FE:
            corr_factor = txlo_factor.get((int(tn), int(freq), int(temp), int(meta.lo_idac)))
            series = raw
            series_kind = "raw"
            if EXTREMA_USE_CORRELATED_MEDIAN:
                if corr_factor is None:
                    continue
                series = raw + float(corr_factor)
                series_kind = "correlated"

            extrema_candidates.setdefault(("TXLO", int(meta.lo_idac)), []).append(
                {
                    "kind": "TXLO",
                    "lo_idac": int(meta.lo_idac),
                    "insertion_type": "FE",
                    "insertion": ins_label,
                    "freq_ghz": int(freq),
                    "temp_c": int(temp),
                    "test_number": int(tn),
                    "test_name": str(meta.test_name),
                    "median": float(_median(series)),
                    "median_series": series_kind,
                }
            )

        if meta.kind == "TXPA" and meta.lut_value == int(EXTREMA_TXPA_LUT):
            insertion_type = str(_INS_TYPE_BY_INS.get(ins_label, ""))
            if insertion_type not in ("FE", "BE"):
                continue

            fac_lookup = txpa_factor_fe if insertion_type == "FE" else txpa_factor_be
            corr_factor = _lookup_txpa_factor(
                fac_lookup,
                lut=int(meta.lut_value),
                freq=int(freq),
                temp_c=int(temp),
                pa_channel=meta.pa_channel,
            )
            series = raw
            series_kind = "raw"
            if EXTREMA_USE_CORRELATED_MEDIAN:
                if corr_factor is None:
                    continue
                series = raw + float(corr_factor)
                series_kind = "correlated"

            pa = (meta.pa_channel or "ALL").strip().upper() or "ALL"
            extrema_candidates.setdefault(("TXPA", insertion_type, pa), []).append(
                {
                    "kind": "TXPA",
                    "lut": int(meta.lut_value),
                    "pa_channel": pa,
                    "insertion_type": insertion_type,
                    "insertion": ins_label,
                    "freq_ghz": int(freq),
                    "temp_c": int(temp),
                    "test_number": int(tn),
                    "test_name": str(meta.test_name),
                    "median": float(_median(series)),
                    "median_series": series_kind,
                }
            )

    extrema_by_key: dict[tuple, dict[str, dict]] = {}
    extrema_summary_rows: list[dict] = []
    for key, rows in extrema_candidates.items():
        rows2 = [r for r in rows if np.isfinite(r.get("median"))]
        if not rows2:
            continue
        r_min = min(rows2, key=lambda r: float(r["median"]))
        r_max = max(rows2, key=lambda r: float(r["median"]))
        extrema_by_key[key] = {"min": r_min, "max": r_max}

        extrema_summary_rows.append(
            {
                "key": str(key),
                "kind": r_min.get("kind"),
                "insertion_type": r_min.get("insertion_type"),
                "lo_idac": r_min.get("lo_idac"),
                "lut": r_min.get("lut"),
                "pa_channel": r_min.get("pa_channel"),
                "median_series": r_min.get("median_series"),
                "min_freq_ghz": r_min.get("freq_ghz"),
                "min_temp_c": r_min.get("temp_c"),
                "min_insertion": r_min.get("insertion"),
                "min_median": r_min.get("median"),
                "min_test_number": r_min.get("test_number"),
                "max_freq_ghz": r_max.get("freq_ghz"),
                "max_temp_c": r_max.get("temp_c"),
                "max_insertion": r_max.get("insertion"),
                "max_median": r_max.get("median"),
                "max_test_number": r_max.get("test_number"),
            }
        )

    # Plot all groups
    cmap = plt.get_cmap("tab10")

    for g in groups:
        if g.insertion_type == "FE":
            fig, axes = plt.subplots(1, 3, figsize=(16.2, 5.2), sharey=True)
            insertion_list = _INSERTION_ORDER_FE
            temp_by_ins = {"S1": 135, "S2": -40, "S3": 25}
        else:
            fig, axes = plt.subplots(1, 2, figsize=(12.8, 5.2), sharey=True)
            insertion_list = _INSERTION_ORDER_BE
            temp_by_ins = {"B1": 135, "B2": 25}

        title_extra = _format_test_names_for_title(g.test_numbers, test_name_by_num)
        if title_extra:
            fig.suptitle(f"{g.title}\n{title_extra}")
        else:
            fig.suptitle(g.title)

        for ax, ins_label in zip(axes, insertion_list, strict=False):
            temp = int(temp_by_ins[ins_label])

            ax_title_base = f"{ins_label} (T={temp}°C)"
            ax.grid(True, alpha=0.25)
            ax.set_xlabel("Power [dBm]")
            if ins_label in ("S1", "B1"):
                ax.set_ylabel("CDF")

            # Add one legend entry for limits (computed per test; if grouped, show the envelope)
            lim_lows: list[float] = []
            lim_highs: list[float] = []
            lim_unit = "dBm"

            channels_seen: dict[str, str] = {}

            sub_has_min = False
            sub_has_max = False
            min_stats: tuple[float, float] | None = None  # (median, mean)
            max_stats: tuple[float, float] | None = None

            corr_stats_lines: list[str] = []

            # Plot each test number as its own series
            for idx, tn in enumerate(g.test_numbers):
                arrays = values_by_ins_test.get((ins_label, int(tn)), [])
                raw = np.concatenate(arrays) if arrays else np.array([], dtype=float)
                if raw.size == 0:
                    continue

                raw_f, n_removed = _filter_outliers_mad(raw, mad_multiplier=MAD_MULTIPLIER)
                if n_removed:
                    outlier_rows.append(
                        {
                            "group": g.title,
                            "insertion": ins_label,
                            "test_number": int(tn),
                            "n_before": int(np.isfinite(raw).sum()),
                            "n_after": int(raw_f.size),
                            "n_removed": int(n_removed),
                            "mad_multiplier": float(MAD_MULTIPLIER),
                        }
                    )
                raw = raw_f
                if raw.size == 0:
                    continue

                meta = meta_by_tn.get(int(tn))
                test_name = (meta.test_name if meta else test_name_by_num.get(int(tn), ""))

                # Determine correlation factor + limits
                corr_factor = None
                corr_low = corr_high = None
                unit = "dBm"

                if g.kind == "TXLO":
                    if meta and meta.frequency_ghz is not None and meta.lo_idac is not None:
                        corr_factor = txlo_factor.get((int(tn), int(meta.frequency_ghz), int(temp), int(meta.lo_idac)))
                        lim_map = txlo_limits.get((int(tn), int(meta.frequency_ghz), int(meta.lo_idac)), {})
                        if ins_label in lim_map:
                            corr_low, corr_high = lim_map[ins_label]
                        unit = "dBm"
                else:
                    if meta and meta.lut_value is not None and meta.frequency_ghz is not None:
                        fac_lookup = txpa_factor_fe if g.insertion_type == "FE" else txpa_factor_be
                        lim_lookup = txpa_limits_fe if g.insertion_type == "FE" else txpa_limits_be
                        corr_factor = _lookup_txpa_factor(
                            fac_lookup,
                            lut=int(meta.lut_value),
                            freq=int(meta.frequency_ghz),
                            temp_c=int(temp),
                            pa_channel=meta.pa_channel,
                        )
                        lim_map = _lookup_txpa_limits(
                            lim_lookup,
                            test_number=int(tn),
                            lut=int(meta.lut_value),
                            freq=int(meta.frequency_ghz),
                            pa_channel=meta.pa_channel,
                        )
                        if ins_label in lim_map:
                            corr_low, corr_high, unit = lim_map[ins_label]

                # TXPA: choose per-channel marker (TX1..TX8). Raw vs corr is encoded by color.
                channel_label: str | None = None
                marker = "o"
                if g.kind == "TXPA":
                    if meta and meta.pa_channel:
                        channel_label = str(meta.pa_channel).strip().upper()
                    elif len(g.test_numbers) > 1:
                        # For grouped LUT 12..244 ranges: tests are TX1..TX8 in increasing order.
                        channel_label = f"TX{idx + 1}"
                    if channel_label in TX_CHANNEL_MARKERS:
                        marker = TX_CHANNEL_MARKERS[channel_label]
                        channels_seen[channel_label] = marker

                label_for_stats = channel_label or (f"T{int(tn)}" if len(g.test_numbers) > 1 else str(int(tn)))

                # Extrema highlighting: determine whether this (freq,temp) is min/max for the requested subsets.
                is_min = False
                is_max = False
                if meta and meta.frequency_ghz is not None:
                    cond = (int(meta.frequency_ghz), int(temp))
                    if g.kind == "TXLO" and meta.lo_idac == int(EXTREMA_TXLO_LO_IDAC):
                        k = ("TXLO", int(meta.lo_idac))
                        if k in extrema_by_key:
                            is_min = cond == (int(extrema_by_key[k]["min"]["freq_ghz"]), int(extrema_by_key[k]["min"]["temp_c"]))
                            is_max = cond == (int(extrema_by_key[k]["max"]["freq_ghz"]), int(extrema_by_key[k]["max"]["temp_c"]))
                    if g.kind == "TXPA" and meta.lut_value == int(EXTREMA_TXPA_LUT):
                        pa = (meta.pa_channel or channel_label or "ALL").strip().upper() or "ALL"
                        k = ("TXPA", str(g.insertion_type), pa)
                        if k in extrema_by_key:
                            is_min = cond == (int(extrema_by_key[k]["min"]["freq_ghz"]), int(extrema_by_key[k]["min"]["temp_c"]))
                            is_max = cond == (int(extrema_by_key[k]["max"]["freq_ghz"]), int(extrema_by_key[k]["max"]["temp_c"]))

                if is_min:
                    sub_has_min = True
                if is_max:
                    sub_has_max = True

                # Raw CDF (scatter points)
                x_raw, y_raw = _cdf_xy(raw)
                ax.scatter(
                    x_raw,
                    y_raw,
                    s=10,
                    marker=marker,
                    facecolors="none",
                    edgecolors=RAW_COLOR,
                    linewidths=0.8,
                    alpha=0.85,
                    label=None,
                )

                # Correlated CDF (scatter points)
                if corr_factor is not None:
                    corr = raw + float(corr_factor)
                    x_c, y_c = _cdf_xy(corr)
                    ax.scatter(
                        x_c,
                        y_c,
                        s=10,
                        marker=marker,
                        facecolors=CORR_COLOR,
                        edgecolors=CORR_COLOR,
                        alpha=0.75,
                        label=None,
                    )

                    c_corr = corr[np.isfinite(corr)]
                    c_med = float(_median(c_corr))
                    c_mean = float(np.mean(c_corr)) if c_corr.size else float("nan")
                    c_n = int(c_corr.size)
                    corr_stats_lines.append(f"{label_for_stats}: med={c_med:.3f} mean={c_mean:.3f} (n={c_n})")
                    corr_stats_rows.append(
                        {
                            "group": g.title,
                            "kind": g.kind,
                            "insertion_type": g.insertion_type,
                            "insertion": ins_label,
                            "temp_c": int(temp),
                            "test_number": int(tn),
                            "test_name": str(test_name),
                            "freq_ghz": int(meta.frequency_ghz) if (meta and meta.frequency_ghz is not None) else None,
                            "lut": int(meta.lut_value) if (meta and meta.lut_value is not None) else g.lut_value,
                            "lo_idac": int(meta.lo_idac) if (meta and meta.lo_idac is not None) else None,
                            "pa_channel": (meta.pa_channel if meta else channel_label),
                            "corr_median": c_med,
                            "corr_mean": c_mean,
                            "n": c_n,
                            "mad_multiplier": float(MAD_MULTIPLIER),
                        }
                    )

                    if is_min or is_max:
                        outline = HIGHLIGHT_MIN_COLOR if is_min else HIGHLIGHT_MAX_COLOR
                        ax.scatter(
                            x_c,
                            y_c,
                            s=18,
                            marker=marker,
                            facecolors="none",
                            edgecolors=outline,
                            linewidths=1.2,
                            alpha=1.0,
                            label=None,
                            zorder=6,
                        )

                        if is_min:
                            min_stats = (c_med, c_mean)
                        if is_max:
                            max_stats = (c_med, c_mean)
                else:
                    missing_factor_rows.append(
                        {
                            "test_number": int(tn),
                            "insertion": ins_label,
                            "kind": g.kind,
                            "freq": getattr(meta, "frequency_ghz", None) if meta else None,
                            "temp": temp,
                            "lut": getattr(meta, "lut_value", None) if meta else g.lut_value,
                            "lo_idac": getattr(meta, "lo_idac", None) if meta else None,
                            "pa_channel": getattr(meta, "pa_channel", None) if meta else None,
                        }
                    )

                    # Still include an explicit marker for missing correlated stats.
                    corr_stats_lines.append(f"{label_for_stats}: corr missing")

                if corr_low is not None and corr_high is not None:
                    lim_lows.append(float(corr_low))
                    lim_highs.append(float(corr_high))
                    lim_unit = unit or lim_unit
                else:
                    missing_limit_rows.append(
                        {
                            "test_number": int(tn),
                            "insertion": ins_label,
                            "kind": g.kind,
                            "freq": getattr(meta, "frequency_ghz", None) if meta else None,
                            "lut": getattr(meta, "lut_value", None) if meta else g.lut_value,
                            "lo_idac": getattr(meta, "lo_idac", None) if meta else None,
                            "pa_channel": getattr(meta, "pa_channel", None) if meta else None,
                        }
                    )

            # Finalize subplot title with extrema markers (only applicable to requested subsets).
            suffix = ""
            if sub_has_min:
                if min_stats is not None:
                    suffix += f" [MIN corr med={min_stats[0]:.3f}, mean={min_stats[1]:.3f}]"
                else:
                    suffix += " [MIN corr]"
            if sub_has_max:
                if max_stats is not None:
                    suffix += f" [MAX corr med={max_stats[0]:.3f}, mean={max_stats[1]:.3f}]"
                else:
                    suffix += " [MAX corr]"
            ax.set_title(ax_title_base + suffix)

            # Correlated stats box (keep out of legend; helps compare condition-sets).
            if corr_stats_lines:
                # Keep stable order: TX1..TX8 first, then others.
                def _sort_key(line: str) -> tuple[int, str]:
                    m = re.match(r"^TX(?P<n>[1-8])\b", line)
                    if m:
                        return (0, m.group("n").zfill(2))
                    return (1, line)

                lines_sorted = sorted(corr_stats_lines, key=_sort_key)
                stats_text = "\n".join(lines_sorted)
                ax.text(
                    0.02,
                    0.98,
                    stats_text,
                    transform=ax.transAxes,
                    ha="left",
                    va="top",
                    fontsize=7,
                    family="monospace",
                    bbox={"boxstyle": "round,pad=0.25", "facecolor": "white", "alpha": 0.75, "edgecolor": "none"},
                )

            # Correlated limits: show and always include values in the legend.
            if lim_lows and lim_highs:
                low_plot = float(min(lim_lows))
                high_plot = float(max(lim_highs))
                ax.axvline(low_plot, color="red", linestyle="--", linewidth=1)
                ax.axvline(high_plot, color="red", linestyle="--", linewidth=1)
                ax.plot(
                    [],
                    [],
                    color="red",
                    linestyle="--",
                    linewidth=1,
                    label=f"Corr limits [{low_plot:.3f}, {high_plot:.3f}] {lim_unit}",
                )
            else:
                ax.plot([], [], color="red", linestyle="--", linewidth=1, label="Corr limits [missing]")

            # Legend: only show keys (raw/corr colors), TX channel markers (TXPA), and correlated limits.
            legend_handles: list[Line2D] = [
                Line2D(
                    [],
                    [],
                    marker="o",
                    linestyle="None",
                    markerfacecolor="none",
                    markeredgecolor=RAW_COLOR,
                    label="ATE raw",
                ),
                Line2D(
                    [],
                    [],
                    marker="o",
                    linestyle="None",
                    markerfacecolor=CORR_COLOR,
                    markeredgecolor=CORR_COLOR,
                    label="ATE correlated",
                ),
            ]

            if g.kind == "TXPA" and channels_seen:
                for ch in sorted(channels_seen.keys()):
                    legend_handles.append(
                        Line2D(
                            [],
                            [],
                            marker=channels_seen[ch],
                            linestyle="None",
                            markerfacecolor="none",
                            markeredgecolor="black",
                            label=ch,
                        )
                    )

            # Include the correlated-limits handle already added via ax.plot([],[]...).
            h, l = ax.get_legend_handles_labels()
            if h and l:
                legend_handles.extend(h)

            ax.legend(handles=legend_handles, loc="lower right", fontsize=7, framealpha=0.9)

        fig.tight_layout(rect=[0, 0.02, 1, 0.92])

        # Output naming
        if g.kind == "TXLO":
            out_path = plots_dir / "FE" / f"TXLO__test_{g.test_numbers[0]}.png"
        else:
            if len(g.test_numbers) > 1:
                out_path = plots_dir / g.insertion_type / f"TXPA__lut_{g.lut_value}__tests_{min(g.test_numbers)}-{max(g.test_numbers)}.png"
            else:
                out_path = plots_dir / g.insertion_type / f"TXPA__lut_{g.lut_value}__test_{g.test_numbers[0]}.png"

        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=PLOT_DPI)
        plt.close(fig)

    # Write small summaries (useful to debug missing matches)
    if missing_factor_rows:
        pd.DataFrame(missing_factor_rows).drop_duplicates().to_csv(out_dir / "missing_factors.csv", index=False)
    if missing_limit_rows:
        pd.DataFrame(missing_limit_rows).drop_duplicates().to_csv(out_dir / "missing_limits.csv", index=False)
    if outlier_rows:
        pd.DataFrame(outlier_rows).to_csv(out_dir / "outliers_removed_mad.csv", index=False)
    if extrema_summary_rows:
        pd.DataFrame(extrema_summary_rows).to_csv(out_dir / "extrema_median_summary.csv", index=False)
    if corr_stats_rows:
        pd.DataFrame(corr_stats_rows).to_csv(out_dir / "correlated_series_stats.csv", index=False)

    print(f"Done. Output: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
