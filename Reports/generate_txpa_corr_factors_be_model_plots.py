"""Generate plots of TXPA correlation factor temperature models (BE).

Creates a multi-page PDF with one page per test case (row) from:
  - TXPA_Corr_Factors_BE.xlsx

Each page contains:
  1) A main plot of factor vs temperature showing measured points (25°C, 135°C)
     and the three extrapolation models used for -40°C.
  2) A row of temperature-specific subplots (-40°C, 25°C, 135°C) showing the
     factor value at that temperature (measured where available) and the three
     model predictions.

Outputs:
  - Tasks_Automation_Code/Reports/txpa_corr_factors_be_models_per_testcase.pdf
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[2]
REPORTS_DIR = REPO_ROOT / "Tasks_Automation_Code" / "Reports"
INPUT_XLSX = REPO_ROOT / "TXPA_Corr_Factors_BE.xlsx"
OUTPUT_PDF = REPORTS_DIR / "txpa_corr_factors_be_models_per_testcase.pdf"


def _k(Tc: float) -> float:
    return Tc + 273.15


def _fit_models_two_point(f25: float, f135: float) -> dict[str, tuple[float, float]]:
    """Return model coefficients for three models, each expressed as f(x)=a*x+b.

    - linearT: x = T(°C)
    - invK:    x = 1/(T+273.15)
    - logK:    x = ln(T+273.15)

    Each model is fit exactly through (25°C, f25) and (135°C, f135).
    """

    T1, T2 = 25.0, 135.0

    # Model A: linear in T
    x1, x2 = T1, T2
    a = (f135 - f25) / (x2 - x1)
    b = f25 - a * x1
    linearT = (a, b)

    # Model B: linear in 1/K
    x1, x2 = 1.0 / _k(T1), 1.0 / _k(T2)
    a = (f135 - f25) / (x2 - x1)
    b = f25 - a * x1
    invK = (a, b)

    # Model C: linear in log(K)
    x1, x2 = float(np.log(_k(T1))), float(np.log(_k(T2)))
    a = (f135 - f25) / (x2 - x1)
    b = f25 - a * x1
    logK = (a, b)

    return {"linearT": linearT, "invK": invK, "logK": logK}


def _predict(model: str, coeffs: tuple[float, float], T: np.ndarray) -> np.ndarray:
    a, b = coeffs
    if model == "linearT":
        x = T
    elif model == "invK":
        x = 1.0 / (_k(0.0) + T)  # (T+273.15)
    elif model == "logK":
        x = np.log(_k(0.0) + T)
    else:
        raise ValueError(f"Unknown model: {model}")
    return a * x + b


def _model_label(model: str) -> str:
    if model == "linearT":
        return "Linear in T: f(T)=a·T+b"
    if model == "invK":
        return "Linear in 1/K: f(T)=a·(1/(T+273.15))+b"
    if model == "logK":
        return "Linear in ln(K): f(T)=a·ln(T+273.15)+b"
    return model


def main() -> int:
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Missing input workbook: {INPUT_XLSX}")

    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    df = pd.read_excel(INPUT_XLSX, sheet_name=0)
    col_135 = [c for c in df.columns if "135" in str(c)][0]
    col_25 = [c for c in df.columns if "25" in str(c)][0]

    f25_s = pd.to_numeric(df[col_25], errors="coerce")
    f135_s = pd.to_numeric(df[col_135], errors="coerce")
    valid = f25_s.notna() & f135_s.notna()

    df = df.loc[valid].copy().reset_index(drop=True)
    f25 = f25_s.loc[valid].to_numpy(dtype=float)
    f135 = f135_s.loc[valid].to_numpy(dtype=float)

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    temps_meas = np.array([25.0, 135.0], dtype=float)
    temps_all = np.array([-40.0, 25.0, 135.0], dtype=float)
    T_line = np.linspace(-40.0, 135.0, 200)

    with PdfPages(OUTPUT_PDF) as pdf:
        for i in range(len(df)):
            coeffs = _fit_models_two_point(f25[i], f135[i])

            # Build a clean title
            lut = df.loc[i, "LUT value"]
            vc = df.loc[i, "Voltage corner"]
            freq = df.loc[i, "Frequency_GHz"]
            chan = df.loc[i, "PA Channel"]
            n = df.loc[i, "N"]
            test_name = str(df.loc[i, "Test Name"]).strip()

            fig = plt.figure(figsize=(13.5, 7.5))
            gs = fig.add_gridspec(2, 3, height_ratios=[2.2, 1.0])

            ax_main = fig.add_subplot(gs[0, :])
            ax_t_m40 = fig.add_subplot(gs[1, 0])
            ax_t_25 = fig.add_subplot(gs[1, 1])
            ax_t_135 = fig.add_subplot(gs[1, 2])

            # Main plot: lines + measured points
            ax_main.scatter(temps_meas, [f25[i], f135[i]], color="black", s=40, label="Measured data")

            for model_key in ["linearT", "invK", "logK"]:
                y_line = _predict(model_key, coeffs[model_key], T_line)
                ax_main.plot(T_line, y_line, linewidth=2, label=_model_label(model_key))

            ax_main.set_xlabel("Temperature (°C)")
            ax_main.set_ylabel("Correlation factor")
            ax_main.grid(True, alpha=0.25)
            ax_main.legend(loc="best", fontsize=9)

            title = f"TXPA Corr Factor vs Temperature | LUT={lut} | {vc} | {freq} GHz | Chan={chan} | N={n}"
            fig.suptitle(title, fontsize=12)
            ax_main.set_title(test_name, fontsize=9)

            # Temperature subplots: show values at each temperature
            def plot_temp_subplot(ax, T: float) -> None:
                y_models = {
                    "linearT": float(_predict("linearT", coeffs["linearT"], np.array([T]))[0]),
                    "invK": float(_predict("invK", coeffs["invK"], np.array([T]))[0]),
                    "logK": float(_predict("logK", coeffs["logK"], np.array([T]))[0]),
                }

                labels = ["linearT", "invK", "logK"]
                ys = [y_models[k] for k in labels]

                ax.scatter(range(len(labels)), ys, s=30)
                ax.set_xticks(range(len(labels)), labels, rotation=0)
                ax.set_title(f"T={T:.0f}°C")
                ax.grid(True, axis="y", alpha=0.25)

                # Measured point if available
                if T == 25.0:
                    ax.axhline(f25[i], color="black", linewidth=1.5, label="measured")
                elif T == 135.0:
                    ax.axhline(f135[i], color="black", linewidth=1.5, label="measured")

                # Tighten y-range around values
                y_all = ys + ([f25[i]] if T == 25.0 else []) + ([f135[i]] if T == 135.0 else [])
                y_min, y_max = min(y_all), max(y_all)
                pad = (y_max - y_min) * 0.15 if y_max != y_min else 0.2
                ax.set_ylim(y_min - pad, y_max + pad)

            plot_temp_subplot(ax_t_m40, -40.0)
            plot_temp_subplot(ax_t_25, 25.0)
            plot_temp_subplot(ax_t_135, 135.0)

            fig.tight_layout(rect=[0, 0.02, 1, 0.93])
            pdf.savefig(fig)
            plt.close(fig)

    print("Wrote", OUTPUT_PDF)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
