"""\
DPLL ↔ TXPA power correlation on PROD raw-data CSVs (flat execution).

Goal
----
For each PROD raw-data CSV in repo folder PROD_Data, extract per-chip values for:
  - DPLL tests: 52046, 52084, 52094
  - TXPA power tests:
        53179-53290, 54146-54258, 55146-55258

Then, for each DPLL test, correlate against every TXPA test above (per file) and
produce:
  - A single Excel report with:
        - an overall summary sheet (all files)
        - one sheet per input file
  - A correlation scatter plot per (file, DPLL test, TXPA test) combination.

Notes
-----
- This script intentionally has no CLI; edit constants below if needed.
- Data extraction relies on
    Tasks_Automation_Code/Reports/tx_supply_compensation_PROD/analyze_tx_supply_compensation_scenarios.py
  because it already implements the required PROD CSV parsing.

"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import sys
from typing import Iterable
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats


# -----------------------------------------------------------------------------
# Configuration (edit as needed)
# -----------------------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parents[2]
INPUT_FOLDER = REPO_ROOT / "PROD_Data"
INPUT_GLOB = "*.csv"
MAX_FILES: int | None = None  # e.g. set to 1 for a quick dry-run

# Make repo root importable even when running this script from another CWD.
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Tasks_Automation_Code.Reports.tx_supply_compensation_PROD.analyze_tx_supply_compensation_scenarios import (  # noqa: E402
    _read_prod_csv_chip_matrix,
    _safe_sheet_name,
)

DPLL_TESTS: list[int] = [52046, 52084, 52094]
TXPA_TESTS: list[int] = (
    list(range(53179, 53291))
    + list(range(54146, 54259))
    + list(range(55146, 55259))
)

# Plot tuning
PLOTS_ENABLED = True
PLOT_DPI = 160
PLOT_MAX_POINTS = None  # None = plot all; set e.g. 5000 for faster plots

# Outlier filtering
# Requirement: use MAD with threshold = MAD_MULTIPLIER * MAD.
# Outlier handling: drop only the specific pair -> we mask outlier points per test column (set to NaN).
OUTLIER_METHOD = "MAD"  # supported: MAD
MAD_MULTIPLIER = 10.0


@dataclass(frozen=True)
class CorrStats:
    n: int
    pearson_r: float
    pearson_p: float
    spearman_rho: float
    spearman_p: float
    kendall_tau: float
    kendall_p: float
    slope: float
    intercept: float
    r_value: float
    p_value: float
    std_err: float
    r2: float
    rmse: float
    mae: float
    x_mean: float
    x_std: float
    y_mean: float
    y_std: float
    note: str


def _read_prod_csv_test_name_map(path: Path, *, tests: Iterable[int] | None = None) -> dict[int, str]:
    """Read PROD CSV header+test-name row and build {test_number: test_name} mapping."""

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


def _bh_fdr_adjust(pvalues: pd.Series) -> pd.Series:
    """Benjamini–Hochberg FDR adjustment.

    Returns a series aligned to pvalues.index.
    """

    p = pd.to_numeric(pvalues, errors="coerce")
    p_np = p.to_numpy(dtype=float, na_value=np.nan)
    mask = p.notna() & np.isfinite(p_np)
    if mask.sum() == 0:
        return pd.Series(index=p.index, dtype=float)

    p_valid = p.loc[mask].astype(float)
    m = float(len(p_valid))
    order = np.argsort(p_valid.to_numpy())
    p_sorted = p_valid.to_numpy()[order]

    # q_i = p_i * m / i
    ranks = np.arange(1, len(p_sorted) + 1, dtype=float)
    q = p_sorted * (m / ranks)
    # Enforce monotonicity from the back
    q = np.minimum.accumulate(q[::-1])[::-1]
    q = np.clip(q, 0.0, 1.0)

    out = pd.Series(index=p.index, dtype=float)
    out.loc[mask] = q[np.argsort(order)]
    return out


def _mask_outliers_per_column(chip_matrix: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Mask outlier points per test column (set to NaN).

    This preserves the chip for other tests; only the affected pair (x/y) loses that point.
    """

    if chip_matrix.empty:
        return chip_matrix, {"outlier_method": OUTLIER_METHOD, "n_outlier_points_masked": 0}

    df = chip_matrix.copy()
    if OUTLIER_METHOD.upper() != "MAD":
        return df, {"outlier_method": OUTLIER_METHOD, "n_outlier_points_masked": 0, "note": "unknown_method"}

    k = float(MAD_MULTIPLIER)

    # Median and MAD per column (ignore NaNs)
    med = df.median(axis=0, numeric_only=True)
    abs_dev = (df - med).abs()
    mad = abs_dev.median(axis=0, numeric_only=True)

    # Threshold per column: k * MAD. If MAD is 0/NaN, do not mask (avoid wiping columns).
    thr = k * mad
    valid_thr = thr.notna() & (thr > 0)
    if valid_thr.sum() == 0:
        return df, {
            "outlier_method": f"MAD(k={k:g})",
            "n_outlier_points_masked": 0,
            "note": "mad_all_zero_or_nan",
        }

    outlier_cells = pd.DataFrame(False, index=df.index, columns=df.columns)
    outlier_cells.loc[:, valid_thr.index[valid_thr]] = abs_dev.loc[:, valid_thr.index[valid_thr]].gt(
        thr.loc[valid_thr], axis=1
    )

    n_masked = int(outlier_cells.to_numpy().sum())
    df[outlier_cells] = np.nan

    return df, {
        "outlier_method": f"MAD(k={k:g})",
        "n_outlier_points_masked": n_masked,
        "mad_multiplier": k,
    }


def _compute_corr_stats(x: np.ndarray, y: np.ndarray) -> CorrStats:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if x.shape != y.shape:
        raise ValueError("x and y must have the same shape")

    n = int(x.size)
    note = ""

    if n < 3:
        return CorrStats(
            n=n,
            pearson_r=float("nan"),
            pearson_p=float("nan"),
            spearman_rho=float("nan"),
            spearman_p=float("nan"),
            kendall_tau=float("nan"),
            kendall_p=float("nan"),
            slope=float("nan"),
            intercept=float("nan"),
            r_value=float("nan"),
            p_value=float("nan"),
            std_err=float("nan"),
            r2=float("nan"),
            rmse=float("nan"),
            mae=float("nan"),
            x_mean=float(np.nanmean(x)) if n else float("nan"),
            x_std=float(np.nanstd(x)) if n else float("nan"),
            y_mean=float(np.nanmean(y)) if n else float("nan"),
            y_std=float(np.nanstd(y)) if n else float("nan"),
            note="too_few_points",
        )

    x_mean = float(np.mean(x))
    y_mean = float(np.mean(y))
    x_std = float(np.std(x, ddof=1))
    y_std = float(np.std(y, ddof=1))

    # Handle constant vectors (scipy raises ValueError for pearsonr/linregress)
    x_const = not np.isfinite(x_std) or x_std == 0.0
    y_const = not np.isfinite(y_std) or y_std == 0.0

    pearson_r = pearson_p = float("nan")
    spearman_rho = spearman_p = float("nan")
    kendall_tau = kendall_p = float("nan")
    slope = intercept = r_value = p_value = std_err = float("nan")

    if x_const or y_const:
        note = "constant_input"
    else:
        pearson_r, pearson_p = stats.pearsonr(x, y)
        spearman_rho, spearman_p = stats.spearmanr(x, y)
        kendall_tau, kendall_p = stats.kendalltau(x, y)

        lr = stats.linregress(x, y)
        slope = float(lr.slope)
        intercept = float(lr.intercept)
        r_value = float(lr.rvalue)
        p_value = float(lr.pvalue)
        std_err = float(lr.stderr) if lr.stderr is not None else float("nan")

    # Errors: based on yhat from linear regression if available, else mean
    if np.isfinite(slope) and np.isfinite(intercept):
        yhat = slope * x + intercept
    else:
        yhat = np.full_like(y, fill_value=y_mean)

    rmse = float(np.sqrt(np.mean((y - yhat) ** 2)))
    mae = float(np.mean(np.abs(y - yhat)))

    r2 = float(pearson_r**2) if np.isfinite(pearson_r) else float("nan")

    return CorrStats(
        n=n,
        pearson_r=float(pearson_r),
        pearson_p=float(pearson_p),
        spearman_rho=float(spearman_rho),
        spearman_p=float(spearman_p),
        kendall_tau=float(kendall_tau),
        kendall_p=float(kendall_p),
        slope=float(slope),
        intercept=float(intercept),
        r_value=float(r_value),
        p_value=float(p_value),
        std_err=float(std_err),
        r2=float(r2),
        rmse=rmse,
        mae=mae,
        x_mean=x_mean,
        x_std=x_std,
        y_mean=y_mean,
        y_std=y_std,
        note=note,
    )


def _plot_correlation(
    *,
    x: np.ndarray,
    y: np.ndarray,
    stats_: CorrStats,
    out_path: Path,
    title: str,
    xlabel: str,
    ylabel: str,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if PLOT_MAX_POINTS is not None and stats_.n > int(PLOT_MAX_POINTS):
        rng = np.random.default_rng(0)
        idx = rng.choice(stats_.n, size=int(PLOT_MAX_POINTS), replace=False)
        x = x[idx]
        y = y[idx]

    fig = plt.figure(figsize=(7.6, 5.4), dpi=PLOT_DPI)
    ax = fig.add_subplot(1, 1, 1)

    ms = 10 if len(x) < 3000 else 6
    ax.scatter(x, y, s=ms, alpha=0.45, edgecolors="none")

    if np.isfinite(stats_.slope) and np.isfinite(stats_.intercept):
        x_line = np.array([float(np.min(x)), float(np.max(x))])
        y_line = stats_.slope * x_line + stats_.intercept
        ax.plot(x_line, y_line, color="black", linewidth=1.4, label="linreg")

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.25)

    txt = (
        f"n={stats_.n}\n"
        f"Pearson r={stats_.pearson_r:.4g} (p={stats_.pearson_p:.3g})\n"
        f"Spearman ρ={stats_.spearman_rho:.4g} (p={stats_.spearman_p:.3g})\n"
        f"Kendall τ={stats_.kendall_tau:.4g} (p={stats_.kendall_p:.3g})\n"
        f"R²={stats_.r2:.4g}\n"
        f"slope={stats_.slope:.4g}, intercept={stats_.intercept:.4g}\n"
        f"RMSE={stats_.rmse:.4g}, MAE={stats_.mae:.4g}"
    )
    if stats_.note:
        txt += f"\nNOTE: {stats_.note}"

    ax.text(
        0.02,
        0.98,
        txt,
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=9.5,
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.85, edgecolor="#999999"),
    )

    fig.tight_layout()
    fig.savefig(out_path, dpi=PLOT_DPI)
    plt.close(fig)


def _discover_input_files() -> list[Path]:
    if not INPUT_FOLDER.exists():
        return []
    files = sorted([p for p in INPUT_FOLDER.glob(INPUT_GLOB) if p.is_file()])
    if MAX_FILES is not None:
        return files[: int(MAX_FILES)]
    return files


def main() -> int:
    run_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(__file__).resolve().parent / f"output_{run_tag}__dpll_txpa_corr"
    out_dir.mkdir(parents=True, exist_ok=True)

    files = _discover_input_files()
    if not files:
        raise RuntimeError(f"No input files found in: {INPUT_FOLDER}")

    tests_needed = set(DPLL_TESTS) | set(TXPA_TESTS)

    all_rows: list[dict] = []
    parse_fail_rows: list[dict] = []

    for path in files:
        file_tag = path.name
        print(f"Reading: {file_tag}")

        chip_matrix, info = _read_prod_csv_chip_matrix(path, tests_needed=tests_needed)
        if chip_matrix.empty:
            parse_fail_rows.append({"file": file_tag, **info})
            print(f"  -> skipped (parse error: {info.get('error','')})")
            continue

        # Read test-name mapping for nicer reporting
        test_name_map = _read_prod_csv_test_name_map(path, tests=tests_needed)

        # Mask outlier points per test column (pair-level drop)
        chip_matrix_masked, outlier_info = _mask_outliers_per_column(chip_matrix)
        if int(outlier_info.get("n_outlier_points_masked", 0)) > 0:
            print(f"  Outlier mask: masked {outlier_info['n_outlier_points_masked']} points")
        chip_matrix = chip_matrix_masked

        plots_root = out_dir / "plots" / Path(file_tag).stem

        file_rows: list[dict] = []

        # Pre-fetch numpy arrays for speed
        cols_present = set(int(c) for c in chip_matrix.columns)
        for dpll_test in DPLL_TESTS:
            if dpll_test not in cols_present:
                for txpa_test in TXPA_TESTS:
                    file_rows.append(
                        {
                            "file": file_tag,
                            "dpll_test": dpll_test,
                            "dpll_test_name": test_name_map.get(int(dpll_test), ""),
                            "txpa_test": txpa_test,
                            "txpa_test_name": test_name_map.get(int(txpa_test), ""),
                            "n": 0,
                            "note": "missing_dpll_test",
                            "outlier_method": outlier_info.get("outlier_method", ""),
                            "n_outlier_points_masked": outlier_info.get("n_outlier_points_masked", 0),
                        }
                    )
                continue

            x_all = chip_matrix[dpll_test].to_numpy(dtype=float)

            for txpa_test in TXPA_TESTS:
                if txpa_test not in cols_present:
                    file_rows.append(
                        {
                            "file": file_tag,
                            "dpll_test": dpll_test,
                            "dpll_test_name": test_name_map.get(int(dpll_test), ""),
                            "txpa_test": txpa_test,
                            "txpa_test_name": test_name_map.get(int(txpa_test), ""),
                            "n": 0,
                            "note": "missing_txpa_test",
                            "outlier_method": outlier_info.get("outlier_method", ""),
                            "n_outlier_points_masked": outlier_info.get("n_outlier_points_masked", 0),
                        }
                    )
                    continue

                y_all = chip_matrix[txpa_test].to_numpy(dtype=float)

                mask = np.isfinite(x_all) & np.isfinite(y_all)
                x = x_all[mask]
                y = y_all[mask]

                stats_ = _compute_corr_stats(x, y)

                plot_rel = ""
                if PLOTS_ENABLED and stats_.n >= 3:
                    plot_path = plots_root / f"dpll_{dpll_test}" / f"txpa_{txpa_test}.png"
                    dpll_name = test_name_map.get(int(dpll_test), "")
                    txpa_name = test_name_map.get(int(txpa_test), "")
                    title = (
                        f"{Path(file_tag).stem}: DPLL {dpll_test} ({dpll_name}) vs TXPA {txpa_test} ({txpa_name})"
                        if (dpll_name or txpa_name)
                        else f"{Path(file_tag).stem}: DPLL {dpll_test} vs TXPA {txpa_test}"
                    )
                    _plot_correlation(
                        x=x,
                        y=y,
                        stats_=stats_,
                        out_path=plot_path,
                        title=title,
                        xlabel=f"DPLL test {dpll_test}" + (f" ({dpll_name})" if dpll_name else ""),
                        ylabel=f"TXPA test {txpa_test}" + (f" ({txpa_name})" if txpa_name else ""),
                    )
                    plot_rel = str(plot_path.relative_to(out_dir)).replace("\\", "/")

                row = {
                    "file": file_tag,
                    "dpll_test": dpll_test,
                    "dpll_test_name": test_name_map.get(int(dpll_test), ""),
                    "txpa_test": txpa_test,
                    "txpa_test_name": test_name_map.get(int(txpa_test), ""),
                    "n": stats_.n,
                    "pearson_r": stats_.pearson_r,
                    "pearson_p": stats_.pearson_p,
                    "spearman_rho": stats_.spearman_rho,
                    "spearman_p": stats_.spearman_p,
                    "kendall_tau": stats_.kendall_tau,
                    "kendall_p": stats_.kendall_p,
                    "slope": stats_.slope,
                    "intercept": stats_.intercept,
                    "r2": stats_.r2,
                    "rmse": stats_.rmse,
                    "mae": stats_.mae,
                    "x_mean": stats_.x_mean,
                    "x_std": stats_.x_std,
                    "y_mean": stats_.y_mean,
                    "y_std": stats_.y_std,
                    "note": stats_.note,
                    "plot_relpath": plot_rel,
                    "outlier_method": outlier_info.get("outlier_method", ""),
                    "n_outlier_points_masked": outlier_info.get("n_outlier_points_masked", 0),
                }

                file_rows.append(row)

        all_rows.extend(file_rows)

    df_all = pd.DataFrame(all_rows)
    df_parse_fail = pd.DataFrame(parse_fail_rows)

    # Add per-file FDR-adjusted p-values (BH) for multiple comparisons.
    if not df_all.empty and "file" in df_all.columns:
        df_all["pearson_p_fdr"] = np.nan
        df_all["spearman_p_fdr"] = np.nan
        df_all["kendall_p_fdr"] = np.nan
        for file_name in df_all["file"].dropna().unique():
            m = df_all["file"] == file_name
            df_all.loc[m, "pearson_p_fdr"] = _bh_fdr_adjust(df_all.loc[m, "pearson_p"])
            df_all.loc[m, "spearman_p_fdr"] = _bh_fdr_adjust(df_all.loc[m, "spearman_p"])
            df_all.loc[m, "kendall_p_fdr"] = _bh_fdr_adjust(df_all.loc[m, "kendall_p"])

    report_path = out_dir / "dpll_txpa_correlation_report.xlsx"
    used_sheet_names: set[str] = set()

    def _preferred_column_order(df: pd.DataFrame) -> pd.DataFrame:
        preferred = [
            "file",
            "dpll_test",
            "dpll_test_name",
            "txpa_test",
            "txpa_test_name",
            "n",
            "pearson_r",
            "pearson_p",
            "pearson_p_fdr",
            "spearman_rho",
            "spearman_p",
            "spearman_p_fdr",
            "kendall_tau",
            "kendall_p",
            "kendall_p_fdr",
            "r2",
            "slope",
            "intercept",
            "rmse",
            "mae",
            "x_mean",
            "x_std",
            "y_mean",
            "y_std",
            "note",
            "outlier_method",
            "n_outlier_points_masked",
            "plot_relpath",
            "plot_link",
        ]
        cols = [c for c in preferred if c in df.columns]
        rest = [c for c in df.columns if c not in cols]
        return df[cols + rest]

    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        # Meta sheet
        meta = pd.DataFrame(
            [
                {
                    "generated_at": datetime.now().isoformat(timespec="seconds"),
                    "input_folder": str(INPUT_FOLDER),
                    "input_glob": INPUT_GLOB,
                    "n_files_found": int(len(files)),
                    "dpll_tests": ",".join(map(str, DPLL_TESTS)),
                    "txpa_tests_count": int(len(TXPA_TESTS)),
                    "txpa_tests_ranges": "53179-53290; 54146-54258; 55146-55258",
                    "outlier_filter": OUTLIER_METHOD,
                    "mad_multiplier": float(MAD_MULTIPLIER) if OUTLIER_METHOD.upper() == "MAD" else "",
                    "fdr_method": "Benjamini-Hochberg",
                }
            ]
        )
        meta.to_excel(writer, sheet_name=_safe_sheet_name("Meta", used_sheet_names), index=False)

        # Overall summary
        if not df_all.empty:
            df_out = df_all.copy()
            df_out.insert(
                df_out.columns.get_loc("plot_relpath") + 1,
                "plot_link",
                df_out["plot_relpath"].map(
                    lambda p: f'=HYPERLINK("{p}","plot")' if isinstance(p, str) and p else ""
                ),
            )
            df_out = _preferred_column_order(df_out)
            df_out.to_excel(writer, sheet_name=_safe_sheet_name("Summary_All", used_sheet_names), index=False)

        # Per-file sheets
        if not df_all.empty and "file" in df_all.columns:
            for file_name in sorted(df_all["file"].dropna().unique()):
                sheet = _safe_sheet_name(Path(file_name).stem, used_sheet_names)
                df_one = df_all.loc[df_all["file"] == file_name].copy()
                df_one.insert(
                    df_one.columns.get_loc("plot_relpath") + 1,
                    "plot_link",
                    df_one["plot_relpath"].map(
                        lambda p: f'=HYPERLINK("{p}","plot")' if isinstance(p, str) and p else ""
                    ),
                )
                df_one = _preferred_column_order(df_one)
                df_one.to_excel(writer, sheet_name=sheet, index=False)

        if not df_parse_fail.empty:
            df_parse_fail.to_excel(writer, sheet_name=_safe_sheet_name("ParseFailures", used_sheet_names), index=False)

    print(f"Processed files: {len(files)}")
    print(f"Excel report:     {report_path}")
    if not df_parse_fail.empty:
        print("Parse failures:   included in Excel report")
    print(f"Plots folder:     {out_dir / 'plots'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
