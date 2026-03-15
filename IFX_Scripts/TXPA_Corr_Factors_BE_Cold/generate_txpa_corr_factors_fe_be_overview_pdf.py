"""Generate FE+BE overview plots of TXPA correlation factors vs temperature.

Reads (repo root):
  - TXPA_Corr_Factors_FE.xlsx
  - TXPA_Corr_Factors_BE.xlsx

For each test case (row), creates a single PDF page with 2 subplots:
  - FE: measured correlation factors across temperatures (-40/25/135 when present)
  - BE: measured correlation factors across temperatures (25/135) plus
        extrapolated -40°C point from a linear model.

Both subplots show a linear fit model: f(T) = a*T + b.

Output:
  - Tasks_Automation_Code/Reports/txpa_corr_factors_fe_be_overview.pdf

Notes:
  - This script assumes one row corresponds to one test case.
  - If FE -40°C is missing/NaN for a row, the FE fit uses remaining points.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[2]
REPORTS_DIR = REPO_ROOT / "Tasks_Automation_Code" / "Reports"
FE_XLSX = REPO_ROOT / "TXPA_Corr_Factors_FE.xlsx"
BE_XLSX = REPO_ROOT / "TXPA_Corr_Factors_BE.xlsx"
OUTPUT_PDF = REPORTS_DIR / "txpa_corr_factors_fe_be_overview.pdf"


KEY_COLUMNS = [
    "DataSheet",
    "LUT value",
    "Test Name",
    "Voltage corner",
    "Frequency_GHz",
    "PA Channel",
    "N",
]


@dataclass(frozen=True)
class LinearFit:
    a: float
    b: float
    r2: float | None

    def predict(self, T: np.ndarray) -> np.ndarray:
        return self.a * T + self.b


def _find_temp_column(df: pd.DataFrame, temp_c: int) -> str | None:
    for c in df.columns:
        s = str(c)
        if "Corr_Factors" in s and str(temp_c) in s:
            return c
    return None


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _linear_fit(T: np.ndarray, y: np.ndarray) -> LinearFit:
    mask = np.isfinite(T) & np.isfinite(y)
    T = T[mask]
    y = y[mask]
    if T.size < 2:
        return LinearFit(a=float("nan"), b=float("nan"), r2=None)

    a, b = np.polyfit(T, y, deg=1)

    r2: float | None
    if T.size >= 3:
        y_hat = a * T + b
        ss_res = float(np.sum((y - y_hat) ** 2))
        ss_tot = float(np.sum((y - float(np.mean(y))) ** 2))
        r2 = None if ss_tot == 0.0 else 1.0 - (ss_res / ss_tot)
    else:
        r2 = None

    return LinearFit(a=float(a), b=float(b), r2=r2)


def _fmt_fit(fit: LinearFit) -> str:
    if not np.isfinite(fit.a) or not np.isfinite(fit.b):
        return "fit: n/a"
    base = f"f(T) = {fit.a:.6g}·T + {fit.b:.6g}"
    if fit.r2 is None:
        return base
    return base + f"  (R²={fit.r2:.3f})"


def _build_title(row: pd.Series) -> str:
    lut = row.get("LUT value", "")
    vc = row.get("Voltage corner", "")
    freq = row.get("Frequency_GHz", "")
    chan = row.get("PA Channel", "")
    n = row.get("N", "")
    return f"TXPA Corr Factors | LUT={lut} | {vc} | {freq} GHz | Chan={chan} | N={n}"


def _plot_case(ax, label: str, temps_meas: np.ndarray, y_meas: np.ndarray, fit: LinearFit, *, y_pred_m40: float | None = None) -> None:
    import matplotlib.pyplot as plt

    T_line = np.linspace(-40.0, 135.0, 200)

    # Measured points
    mask = np.isfinite(temps_meas) & np.isfinite(y_meas)
    ax.scatter(temps_meas[mask], y_meas[mask], color="black", s=36, label="Measured")

    # Extrapolated -40 point (BE)
    if y_pred_m40 is not None and np.isfinite(y_pred_m40):
        ax.scatter([-40.0], [y_pred_m40], marker="x", s=60, linewidths=2, color="tab:red", label="-40°C (model)")

    # Fit line
    if np.isfinite(fit.a) and np.isfinite(fit.b):
        ax.plot(T_line, fit.predict(T_line), color="tab:blue", linewidth=2, label="Linear fit")

    ax.set_title(label)
    ax.set_xlabel("Temperature (°C)")
    ax.set_ylabel("Correlation factor")
    ax.grid(True, alpha=0.25)

    # Fit text
    ax.text(
        0.02,
        0.98,
        _fmt_fit(fit),
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=9,
        bbox=dict(boxstyle="round,pad=0.25", facecolor="white", alpha=0.8, edgecolor="none"),
    )

    ax.set_xlim(-45, 140)

    # Legend
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(loc="best", fontsize=9)


def main() -> int:
    if not FE_XLSX.exists():
        raise FileNotFoundError(f"Missing input workbook: {FE_XLSX}")
    if not BE_XLSX.exists():
        raise FileNotFoundError(f"Missing input workbook: {BE_XLSX}")

    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    df_fe = pd.read_excel(FE_XLSX, sheet_name=0)
    df_be = pd.read_excel(BE_XLSX, sheet_name=0)

    # Identify columns
    fe_col_m40 = _find_temp_column(df_fe, -40)
    fe_col_25 = _find_temp_column(df_fe, 25)
    fe_col_135 = _find_temp_column(df_fe, 135)

    be_col_25 = _find_temp_column(df_be, 25)
    be_col_135 = _find_temp_column(df_be, 135)

    missing = [
        ("FE 25", fe_col_25),
        ("FE 135", fe_col_135),
        ("BE 25", be_col_25),
        ("BE 135", be_col_135),
    ]
    for name, col in missing:
        if col is None:
            raise ValueError(f"Could not find required column for {name}°C in workbook")

    # Merge FE + BE by test case keys
    common_keys = [c for c in KEY_COLUMNS if c in df_fe.columns and c in df_be.columns]
    if not common_keys:
        raise ValueError("No common key columns found to merge FE and BE workbooks")

    df = df_fe.merge(df_be, on=common_keys, how="inner", suffixes=("_FE", "_BE"))
    if df.empty:
        raise ValueError("FE/BE merge produced 0 rows; check key columns and workbook alignment")

    # Pull factor columns from merged df
    # After merge, factor columns from FE remain with original names; BE factor columns may have suffix.
    def pick_col(original_name: str, suffix: str) -> str:
        if original_name in df.columns:
            return original_name
        alt = f"{original_name}{suffix}"
        if alt in df.columns:
            return alt
        # Try with underscore
        alt = f"{original_name}_{suffix.strip('_')}"
        if alt in df.columns:
            return alt
        raise KeyError(f"Missing expected column after merge: {original_name}")

    fe25 = pick_col(fe_col_25, "_FE")
    fe135 = pick_col(fe_col_135, "_FE")
    fe_m40 = None
    if fe_col_m40 is not None:
        # FE -40 column exists in the FE workbook; should survive merge
        fe_m40 = pick_col(fe_col_m40, "_FE")

    be25 = pick_col(be_col_25, "_BE")
    be135 = pick_col(be_col_135, "_BE")

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    with PdfPages(OUTPUT_PDF) as pdf:
        for _, row in df.iterrows():
            test_name = str(row.get("Test Name", "")).strip()

            fe_vals = {
                25.0: float(_to_num(pd.Series([row[fe25]])).iloc[0]),
                135.0: float(_to_num(pd.Series([row[fe135]])).iloc[0]),
            }
            if fe_m40 is not None:
                fe_vals[-40.0] = float(_to_num(pd.Series([row[fe_m40]])).iloc[0])

            be_vals = {
                25.0: float(_to_num(pd.Series([row[be25]])).iloc[0]),
                135.0: float(_to_num(pd.Series([row[be135]])).iloc[0]),
            }

            # FE fit uses available points (including -40 if present)
            fe_T = np.array(sorted(fe_vals.keys()), dtype=float)
            fe_y = np.array([fe_vals[t] for t in fe_T], dtype=float)
            fe_fit = _linear_fit(fe_T, fe_y)

            # BE fit uses measured points (25, 135)
            be_T = np.array([25.0, 135.0], dtype=float)
            be_y = np.array([be_vals[25.0], be_vals[135.0]], dtype=float)
            be_fit = _linear_fit(be_T, be_y)
            be_pred_m40 = float(be_fit.predict(np.array([-40.0], dtype=float))[0]) if np.isfinite(be_fit.a) else float("nan")

            fig, (ax_fe, ax_be) = plt.subplots(1, 2, figsize=(13.5, 6.5), sharey=False)

            _plot_case(ax_fe, "FE", fe_T, fe_y, fe_fit)
            _plot_case(ax_be, "BE", be_T, be_y, be_fit, y_pred_m40=be_pred_m40)

            fig.suptitle(_build_title(row), fontsize=12)
            if test_name:
                fig.text(0.5, 0.94, test_name, ha="center", va="top", fontsize=9)

            fig.tight_layout(rect=[0, 0.0, 1, 0.92])
            pdf.savefig(fig)
            plt.close(fig)

    print("Wrote", OUTPUT_PDF)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
