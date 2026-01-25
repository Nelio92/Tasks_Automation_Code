"""TX PA supply compensation strategy simulation.

Compares two approaches:
- Per-pair average compensation (pairs: TX1&TX5, TX2&TX6, TX3&TX7, TX4&TX8)
- One global average compensation applied to all TX channels at once

The model is intentionally simple:
- Each TX channel i has a required compensation offset o_i (mV)
- If we apply compensation a_i, the residual at the block is r_i = a_i - o_i (mV)
  * r_i > 0 : over-compensation (block sees higher voltage than target)
  * r_i < 0 : under-compensation (block sees lower voltage than target)

Outputs:
- Excel workbook with parameters, summaries, and sample raw trials
- PNG plots (distributions and sigma sweep curves)

Run:
  C:/.../.venv/Scripts/python.exe Tasks_Automation_Code/Reports/tx_supply_compensation_sim/simulate_tx_supply_compensation.py

"""

from __future__ import annotations

import argparse
import datetime as _dt
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt


PAIRS_1BASED: list[tuple[int, int]] = [(1, 5), (2, 6), (3, 7), (4, 8)]
PAIRS_0BASED: list[tuple[int, int]] = [(a - 1, b - 1) for (a, b) in PAIRS_1BASED]
CHANNEL_NAMES = [f"TX{i}" for i in range(1, 9)]
DEFAULT_ACCEPTABLE_ABS_RESIDUAL_MV = 20.0


@dataclass(frozen=True)
class SimConfig:
    seed: int
    n_trials: int
    mean_offset_mv: float
    base_sigma_mv: float
    clip_min_mv: float


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _pair_applied_offsets(offsets_mv: np.ndarray) -> np.ndarray:
    """Return per-channel applied offsets under 'per-pair average' method.

    offsets_mv shape: (n_trials, 8)
    returns shape: (n_trials, 8)
    """
    applied = np.empty_like(offsets_mv)
    for i, j in PAIRS_0BASED:
        pair_avg = 0.5 * (offsets_mv[:, i] + offsets_mv[:, j])
        applied[:, i] = pair_avg
        applied[:, j] = pair_avg
    return applied


def _global_applied_offsets(offsets_mv: np.ndarray) -> np.ndarray:
    """Return per-channel applied offsets under 'global average' method."""
    global_avg = offsets_mv.mean(axis=1, keepdims=True)
    return np.repeat(global_avg, repeats=offsets_mv.shape[1], axis=1)


def _residuals(offsets_mv: np.ndarray, applied_mv: np.ndarray) -> np.ndarray:
    return applied_mv - offsets_mv


def _trial_metrics(residuals_mv: np.ndarray) -> pd.DataFrame:
    """Compute per-trial residual metrics (each row = one trial)."""
    abs_res = np.abs(residuals_mv)
    return pd.DataFrame(
        {
            "mean_abs_residual_mv": abs_res.mean(axis=1),
            "rms_residual_mv": np.sqrt((residuals_mv**2).mean(axis=1)),
            "max_abs_residual_mv": abs_res.max(axis=1),
            "min_residual_mv": residuals_mv.min(axis=1),
            "max_residual_mv": residuals_mv.max(axis=1),
        }
    )


def _exceedance_table(residuals_mv: np.ndarray, thresholds_mv: list[float]) -> pd.Series:
    abs_res = np.abs(residuals_mv)
    series: dict[str, float] = {}
    for t in thresholds_mv:
        series[f"pct_channels_|res|>={t:g}mV"] = 100.0 * float((abs_res >= t).mean())
    return pd.Series(series)


def _acceptance_table(residuals_mv: np.ndarray, acceptable_abs_mv: float) -> pd.Series:
    abs_res = np.abs(residuals_mv)
    within = abs_res <= acceptable_abs_mv
    # Note: mean() on bool gives fraction
    return pd.Series(
        {
            f"pct_channels_|res|<={acceptable_abs_mv:g}mV": 100.0 * float(within.mean()),
            f"pct_trials_all_channels_|res|<={acceptable_abs_mv:g}mV": 100.0 * float(within.all(axis=1).mean()),
            f"pct_trials_any_channel_|res|>{acceptable_abs_mv:g}mV": 100.0 * float((~within.all(axis=1)).mean()),
        }
    )


def _gating_table(offsets_mv: np.ndarray, acceptable_abs_mv: float) -> pd.Series:
        """ATE-friendly deterministic gating metrics.

        In this simplified static model:
        - Global-average method yields residual_i = mean(offsets) - offset_i.
            Therefore ALL channels are within ±A iff max_i |offset_i - mean(offsets)| <= A.
        - Per-pair method yields residuals of ±(offset_i - offset_j)/2 within each pair.
            Therefore ALL channels are within ±A iff for each pair |offset_i - offset_j| <= 2A.

        We report per-trial pass rates for those conditions.
        """

        a = float(acceptable_abs_mv)
        means = offsets_mv.mean(axis=1, keepdims=True)
        max_dev_global = np.max(np.abs(offsets_mv - means), axis=1)
        global_pass = max_dev_global <= a

        # For per-pair, check the pair-wise difference bound
        pair_diffs = []
        for i, j in PAIRS_0BASED:
                pair_diffs.append(np.abs(offsets_mv[:, i] - offsets_mv[:, j]))
        max_pair_diff = np.max(np.stack(pair_diffs, axis=1), axis=1)
        pair_pass = max_pair_diff <= (2.0 * a)

        return pd.Series(
                {
                        f"pct_trials_gate_global_maxDev_le_{a:g}mV": 100.0 * float(global_pass.mean()),
                        f"pct_trials_gate_pair_maxPairDiff_le_{2*a:g}mV": 100.0 * float(pair_pass.mean()),
                        f"gate_global_maxDev_p95_mv": float(np.quantile(max_dev_global, 0.95)),
                        f"gate_pair_maxPairDiff_p95_mv": float(np.quantile(max_pair_diff, 0.95)),
                }
        )


def _summary_table(
    offsets_mv: np.ndarray,
    residuals_pair_mv: np.ndarray,
    residuals_global_mv: np.ndarray,
    label: str,
    acceptable_abs_mv: float,
) -> pd.DataFrame:
    thresholds = [5, 10, 20, 50, 100]
    acceptable = float(acceptable_abs_mv)

    pair_metrics = _trial_metrics(residuals_pair_mv)
    global_metrics = _trial_metrics(residuals_global_mv)

    def describe_cols(df: pd.DataFrame, prefix: str) -> dict[str, float]:
        out: dict[str, float] = {}
        for col in df.columns:
            out[f"{prefix}_{col}_p50"] = float(df[col].quantile(0.50))
            out[f"{prefix}_{col}_p90"] = float(df[col].quantile(0.90))
            out[f"{prefix}_{col}_p95"] = float(df[col].quantile(0.95))
            out[f"{prefix}_{col}_p99"] = float(df[col].quantile(0.99))
        return out

    row = {
        "scenario": label,
        "offset_mean_mv": float(offsets_mv.mean()),
        "offset_std_mv": float(offsets_mv.std(ddof=1)),
        "offset_min_mv": float(offsets_mv.min()),
        "offset_max_mv": float(offsets_mv.max()),
        **describe_cols(pair_metrics, "pair"),
        **describe_cols(global_metrics, "global"),
        **_exceedance_table(residuals_pair_mv, thresholds).add_prefix("pair_").to_dict(),
        **_exceedance_table(residuals_global_mv, thresholds).add_prefix("global_").to_dict(),
        **_acceptance_table(residuals_pair_mv, acceptable).add_prefix("pair_").to_dict(),
        **_acceptance_table(residuals_global_mv, acceptable).add_prefix("global_").to_dict(),
        **_gating_table(offsets_mv, acceptable).to_dict(),
    }

    return pd.DataFrame([row])


def _plot_hist_max_abs(
    metrics_pair: pd.DataFrame,
    metrics_global: pd.DataFrame,
    title: str,
    out_path: Path,
) -> None:
    plt.figure(figsize=(10, 5))
    bins = 60
    plt.hist(metrics_pair["max_abs_residual_mv"], bins=bins, alpha=0.6, label="Per-pair avg")
    plt.hist(metrics_global["max_abs_residual_mv"], bins=bins, alpha=0.6, label="Global avg")
    plt.xlabel("Max |residual| per trial (mV)")
    plt.ylabel("Count")
    plt.title(title)
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_path, dpi=160)
    plt.close()


def _plot_sigma_sweep(
    df: pd.DataFrame,
    y_col_pair: str,
    y_col_global: str,
    title: str,
    y_label: str,
    out_path: Path,
) -> None:
    plt.figure(figsize=(10, 5))
    plt.plot(df["sigma_mult"], df[y_col_pair], marker="o", label="Per-pair avg")
    plt.plot(df["sigma_mult"], df[y_col_global], marker="o", label="Global avg")
    plt.xlabel("Sigma multiplier (k) where std = k * base_sigma")
    plt.ylabel(y_label)
    plt.title(title)
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_path, dpi=160)
    plt.close()


def _simulate_offsets_normal(
    rng: np.random.Generator,
    n_trials: int,
    mean_mv: float,
    std_mv: float,
    clip_min_mv: float,
) -> np.ndarray:
    offsets = rng.normal(loc=mean_mv, scale=std_mv, size=(n_trials, 8))
    if clip_min_mv is not None:
        offsets = np.clip(offsets, clip_min_mv, None)
    return offsets


def _simulate_offsets_outliers(
    rng: np.random.Generator,
    n_trials: int,
    mean_mv: float,
    std_mv: float,
    n_outliers: int,
    outlier_mv: float,
    clip_min_mv: float,
) -> np.ndarray:
    offsets = _simulate_offsets_normal(rng, n_trials, mean_mv, std_mv, clip_min_mv)
    for t in range(n_trials):
        idx = rng.choice(8, size=n_outliers, replace=False)
        offsets[t, idx] = outlier_mv
    return offsets


def main() -> int:
    parser = argparse.ArgumentParser(description="Simulate TX PA supply compensation strategy tradeoffs")
    parser.add_argument("--out-dir", default="", help="Output directory (default: dated folder next to this script)")
    parser.add_argument("--seed", type=int, default=12345)
    parser.add_argument("--n-trials", type=int, default=20000)
    parser.add_argument("--mean-offset-mv", type=float, default=20.0)
    parser.add_argument("--base-sigma-mv", type=float, default=2.0, help="Base sigma for scenario 2")
    parser.add_argument(
        "--similar-sigma-mv",
        type=float,
        default=1.0,
        help="Small sigma used in scenario 1 (""all offsets similar"")",
    )
    parser.add_argument("--clip-min-mv", type=float, default=0.0, help="Clip offsets to be >= this value")
    parser.add_argument(
        "--acceptable-abs-residual-mv",
        type=float,
        default=DEFAULT_ACCEPTABLE_ABS_RESIDUAL_MV,
        help="Acceptance limit for |residual| (mV). Used for additional risk metrics.",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    if args.out_dir:
        out_dir = Path(args.out_dir)
    else:
        stamp = _dt.date.today().strftime("%Y%m%d")
        out_dir = script_dir / f"output_{stamp}"

    _ensure_dir(out_dir)

    rng = np.random.default_rng(args.seed)

    # Scenario 1: all offsets similar (around 20mV)
    s1_offsets = _simulate_offsets_normal(
        rng=rng,
        n_trials=args.n_trials,
        mean_mv=args.mean_offset_mv,
        std_mv=args.similar_sigma_mv,
        clip_min_mv=args.clip_min_mv,
    )
    s1_pair_applied = _pair_applied_offsets(s1_offsets)
    s1_global_applied = _global_applied_offsets(s1_offsets)
    s1_pair_res = _residuals(s1_offsets, s1_pair_applied)
    s1_global_res = _residuals(s1_offsets, s1_global_applied)

    s1_pair_metrics = _trial_metrics(s1_pair_res)
    s1_global_metrics = _trial_metrics(s1_global_res)

    # Scenario 2: sigma sweep 1..10
    sigma_rows: list[dict[str, float]] = []
    for k in range(1, 11):
        std = float(k * args.base_sigma_mv)
        offsets = _simulate_offsets_normal(
            rng=rng,
            n_trials=args.n_trials,
            mean_mv=args.mean_offset_mv,
            std_mv=std,
            clip_min_mv=args.clip_min_mv,
        )
        pair_res = _residuals(offsets, _pair_applied_offsets(offsets))
        global_res = _residuals(offsets, _global_applied_offsets(offsets))

        pair_metrics = _trial_metrics(pair_res)
        global_metrics = _trial_metrics(global_res)

        acc = float(args.acceptable_abs_residual_mv)
        pair_within = (np.abs(pair_res) <= acc)
        global_within = (np.abs(global_res) <= acc)

        means = offsets.mean(axis=1, keepdims=True)
        max_dev_global = np.max(np.abs(offsets - means), axis=1)
        global_gate = max_dev_global <= acc
        pair_diffs = []
        for i, j in PAIRS_0BASED:
            pair_diffs.append(np.abs(offsets[:, i] - offsets[:, j]))
        max_pair_diff = np.max(np.stack(pair_diffs, axis=1), axis=1)
        pair_gate = max_pair_diff <= (2.0 * acc)

        sigma_rows.append(
            {
                "sigma_mult": float(k),
                "std_mv": std,
                "pair_max_abs_p95": float(pair_metrics["max_abs_residual_mv"].quantile(0.95)),
                "global_max_abs_p95": float(global_metrics["max_abs_residual_mv"].quantile(0.95)),
                "pair_max_abs_p99": float(pair_metrics["max_abs_residual_mv"].quantile(0.99)),
                "global_max_abs_p99": float(global_metrics["max_abs_residual_mv"].quantile(0.99)),
                "pair_mean_abs_p95": float(pair_metrics["mean_abs_residual_mv"].quantile(0.95)),
                "global_mean_abs_p95": float(global_metrics["mean_abs_residual_mv"].quantile(0.95)),
                "pair_pct_channels_ge_20mV": 100.0 * float((np.abs(pair_res) >= 20).mean()),
                "global_pct_channels_ge_20mV": 100.0 * float((np.abs(global_res) >= 20).mean()),
                f"pair_pct_channels_within_{int(acc)}mV": 100.0 * float(pair_within.mean()),
                f"global_pct_channels_within_{int(acc)}mV": 100.0 * float(global_within.mean()),
                f"pair_pct_trials_all_within_{int(acc)}mV": 100.0 * float(pair_within.all(axis=1).mean()),
                f"global_pct_trials_all_within_{int(acc)}mV": 100.0 * float(global_within.all(axis=1).mean()),
                f"global_gate_pct_trials_maxDev_le_{int(acc)}mV": 100.0 * float(global_gate.mean()),
                f"pair_gate_pct_trials_maxPairDiff_le_{int(2*acc)}mV": 100.0 * float(pair_gate.mean()),
            }
        )

    s2_sigma_df = pd.DataFrame(sigma_rows)

    # Scenario 3: outliers
    s3_cases = [
        ("outlier_1ch_120mV", 1, 120.0),
        ("outlier_2ch_120mV", 2, 120.0),
        ("outlier_1ch_150mV", 1, 150.0),
        ("outlier_2ch_150mV", 2, 150.0),
    ]

    s3_summaries: list[pd.DataFrame] = []
    s3_samples: dict[str, tuple[np.ndarray, np.ndarray, np.ndarray]] = {}
    for label, n_out, out_mv in s3_cases:
        offsets = _simulate_offsets_outliers(
            rng=rng,
            n_trials=args.n_trials,
            mean_mv=args.mean_offset_mv,
            std_mv=args.base_sigma_mv,
            n_outliers=n_out,
            outlier_mv=out_mv,
            clip_min_mv=args.clip_min_mv,
        )
        pair_res = _residuals(offsets, _pair_applied_offsets(offsets))
        global_res = _residuals(offsets, _global_applied_offsets(offsets))
        s3_summaries.append(
            _summary_table(
                offsets,
                pair_res,
                global_res,
                label=label,
                acceptable_abs_mv=args.acceptable_abs_residual_mv,
            )
        )
        s3_samples[label] = (offsets, pair_res, global_res)

    s1_summary = _summary_table(
        s1_offsets,
        s1_pair_res,
        s1_global_res,
        label="similar_offsets",
        acceptable_abs_mv=args.acceptable_abs_residual_mv,
    )
    s3_summary = pd.concat(s3_summaries, ignore_index=True)

    # Build a sample raw table (keep XLSX size reasonable)
    sample_n = min(200, args.n_trials)
    def sample_df(offsets: np.ndarray, pair_res: np.ndarray, global_res: np.ndarray) -> pd.DataFrame:
        idx = np.arange(sample_n)
        df = pd.DataFrame(offsets[idx, :], columns=[f"{c}_offset_mv" for c in CHANNEL_NAMES])
        for i, c in enumerate(CHANNEL_NAMES):
            df[f"{c}_pair_residual_mv"] = pair_res[idx, i]
            df[f"{c}_global_residual_mv"] = global_res[idx, i]
        return df

    s1_sample = sample_df(s1_offsets, s1_pair_res, s1_global_res)

    # Excel export
    xlsx_path = out_dir / "tx_supply_compensation_simulation.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        params_df = pd.DataFrame(
            {
                "parameter": [
                    "seed",
                    "n_trials",
                    "mean_offset_mv",
                    "base_sigma_mv",
                    "similar_sigma_mv",
                    "clip_min_mv",
                    "acceptable_abs_residual_mv",
                    "pairs",
                ],
                "value": [
                    args.seed,
                    args.n_trials,
                    args.mean_offset_mv,
                    args.base_sigma_mv,
                    args.similar_sigma_mv,
                    args.clip_min_mv,
                    args.acceptable_abs_residual_mv,
                    ", ".join([f"TX{a}&TX{b}" for a, b in PAIRS_1BASED]),
                ],
            }
        )
        params_df.to_excel(writer, sheet_name="Parameters", index=False)

        s1_summary.to_excel(writer, sheet_name="Scenario1_Summary", index=False)
        s1_sample.to_excel(writer, sheet_name="Scenario1_Sample", index=False)

        s2_sigma_df.to_excel(writer, sheet_name="Scenario2_SigmaSweep", index=False)

        s3_summary.to_excel(writer, sheet_name="Scenario3_Summary", index=False)

        # Add one representative sample table per outlier case
        for label, (offsets, pair_res, global_res) in s3_samples.items():
            tab = sample_df(offsets, pair_res, global_res)
            sheet = ("S3_" + label)[:31]  # Excel sheet name limit
            tab.to_excel(writer, sheet_name=sheet, index=False)

    # Plots
    _plot_hist_max_abs(
        s1_pair_metrics,
        s1_global_metrics,
        title="Scenario 1: Similar offsets around 20mV (max |residual| per trial)",
        out_path=out_dir / "scenario1_hist_max_abs_residual.png",
    )

    _plot_sigma_sweep(
        s2_sigma_df,
        y_col_pair="pair_max_abs_p95",
        y_col_global="global_max_abs_p95",
        title="Scenario 2: Sigma sweep (95th percentile of max |residual|)",
        y_label="p95(max |residual|) [mV]",
        out_path=out_dir / "scenario2_sigmaSweep_p95_maxAbs.png",
    )

    _plot_sigma_sweep(
        s2_sigma_df,
        y_col_pair="pair_max_abs_p99",
        y_col_global="global_max_abs_p99",
        title="Scenario 2: Sigma sweep (99th percentile of max |residual|)",
        y_label="p99(max |residual|) [mV]",
        out_path=out_dir / "scenario2_sigmaSweep_p99_maxAbs.png",
    )

    _plot_sigma_sweep(
        s2_sigma_df,
        y_col_pair="pair_pct_channels_ge_20mV",
        y_col_global="global_pct_channels_ge_20mV",
        title="Scenario 2: Sigma sweep (channels with |residual| >= 20mV)",
        y_label="% of channels across all trials",
        out_path=out_dir / "scenario2_sigmaSweep_pct_channels_ge20mV.png",
    )

    acc_col_pair = f"pair_pct_trials_all_within_{int(args.acceptable_abs_residual_mv)}mV"
    acc_col_global = f"global_pct_trials_all_within_{int(args.acceptable_abs_residual_mv)}mV"
    if acc_col_pair in s2_sigma_df.columns and acc_col_global in s2_sigma_df.columns:
        _plot_sigma_sweep(
            s2_sigma_df,
            y_col_pair=acc_col_pair,
            y_col_global=acc_col_global,
            title=f"Scenario 2: Sigma sweep (% trials with all channels |residual| <= {args.acceptable_abs_residual_mv:g}mV)",
            y_label="% of trials",
            out_path=out_dir / f"scenario2_sigmaSweep_pct_trials_allWithin_{int(args.acceptable_abs_residual_mv)}mV.png",
        )

    # Scenario 3 histograms
    for label, (offsets, pair_res, global_res) in s3_samples.items():
        pair_m = _trial_metrics(pair_res)
        global_m = _trial_metrics(global_res)
        _plot_hist_max_abs(
            pair_m,
            global_m,
            title=f"Scenario 3: {label} (max |residual| per trial)",
            out_path=out_dir / f"scenario3_{label}_hist_maxAbs.png",
        )

    print(f"Wrote Excel: {xlsx_path}")
    print(f"Wrote plots into: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
