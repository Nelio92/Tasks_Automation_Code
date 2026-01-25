"""\
TX supply IR-loss compensation: avg-offset vs per-pair offset (PROD data).

What this assesses
------------------
In production, the IR-loss compensation offset used later in the PA measurement
flow is computed from *pre-compensation* voltage measurements.

This analysis checks whether we can replace the current strategy:
    - Per-channel-pair compensation offset (4 independent offsets)
with a simplified strategy:
    - One averaged compensation offset (average across the 4 channel-pairs)

Key point:
- For the common linear form offset = V_ref - V_measured, using an averaged
  offset instead of per-pair offsets introduces an "applied offset error" of:

      error_i = offset_avg - offset_i = (V_i - mean(V_0..V_3))

So we can assess the simplified strategy directly from the pre-comp measurements
without needing to know V_ref.

Outputs
-------
For each PROD raw-data CSV in an input folder, this script produces:
- One Excel workbook with:
    - Summary_Cases: per file / supply / corner / scenario summary (p99-based)
    - Summary_Pairs: long-format per-pair residual summaries
    - One sheet per input file containing the same two tables
    - Optional ChipLevel sheet (can be large)
- Plots per file/case/scenario/pair:
    - sorted residual (mV) with ±threshold lines
    - sorted max-abs residual across the 4 pairs (mV)

Assumptions about PROD CSV format
--------------------------------
Same as the scenario-comparison script in this folder:
- semicolon-delimited wide format
- "Test Nr" header row first
- rows 6..13 (1-based) are summary stats and are skipped
- per-chip rows start after the "Unit" row

"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

DEFAULT_SCALE_MV_PER_UNIT = 1000.0
DEFAULT_ACCEPTABLE_RESIDUAL_MV = 5.0


@dataclass(frozen=True)
class CaseDef:
    name: str
    supply: str
    corner: str
    scenario1_tests: list[int]
    scenario2_tests: list[int]


CASES: list[CaseDef] = [
    CaseDef(
        name="VDD1V8TX_MIN",
        supply="VDD1V8TX",
        corner="MIN",
        scenario1_tests=[51010, 51011, 51012, 51013],
        scenario2_tests=[51090, 51091, 51092, 51093],
    ),
    CaseDef(
        name="VDD1V0TX_MIN",
        supply="VDD1V0TX",
        corner="MIN",
        scenario1_tests=[51030, 51031, 51032, 51033],
        scenario2_tests=[51110, 51111, 51112, 51113],
    ),
    CaseDef(
        name="VDD1V8TX_MAX",
        supply="VDD1V8TX",
        corner="MAX",
        scenario1_tests=[51050, 51051, 51052, 51053],
        scenario2_tests=[51130, 51131, 51132, 51133],
        ),
    CaseDef(
        name="VDD1V0TX_MAX",
        supply="VDD1V0TX",
        corner="MAX",
        scenario1_tests=[51070, 51071, 51072, 51073],
        scenario2_tests=[51150, 51151, 51152, 51153],
    ),
]


def _find_header_row_index(path: Path, needle: str = "Test Nr") -> int | None:
    try:
        with path.open("r", encoding="latin1", errors="ignore") as f:
            for i, line in enumerate(f):
                if needle in line:
                    return i
    except Exception:
        return None
    return None


def _read_prod_csv_chip_matrix(path: Path, *, tests_needed: set[int] | None = None) -> tuple[pd.DataFrame, dict]:
    """Return (chip_matrix, info) for standard PROD export CSVs.

    chip_matrix:
        index   -> chip_id (LOT/WAFER/X/Y/VNr composed when possible)
        columns -> test number (int)
        values  -> measurement (float)
    """

    header_row = _find_header_row_index(path, needle="Test Nr")
    if header_row is None:
        return pd.DataFrame(), {"error": "header_not_found"}
    if header_row != 0:
        return pd.DataFrame(), {"error": "header_not_first_line", "header_row": header_row}

    def _parse_first_two_lines_csv(p: Path):
        with p.open("r", encoding="latin1", errors="ignore") as f:
            line1 = f.readline().rstrip("\n\r")
            line2 = f.readline().rstrip("\n\r")
        header_cells = line1.split(";")
        test_name_cells = line2.split(";")
        if len(test_name_cells) < len(header_cells):
            test_name_cells = test_name_cells + [""] * (len(header_cells) - len(test_name_cells))
        return header_cells, test_name_cells

    header_cells, _ = _parse_first_two_lines_csv(path)
    header_cells_norm = [str(v).strip() for v in header_cells]

    meta_cols_upper = {"TEST NR", "VNR", "LOT", "WAFER", "X", "Y"}
    usecols: list[int] | None = None
    if tests_needed:
        tests_needed_int = {int(t) for t in tests_needed}
        usecols = []
        for idx, name in enumerate(header_cells_norm):
            n = str(name).strip()
            n_upper = n.upper()
            if n_upper in meta_cols_upper:
                usecols.append(idx)
                continue
            if n.isdigit() and int(n) in tests_needed_int:
                usecols.append(idx)
        if len(usecols) < 3:
            usecols = None

    try:
        raw = pd.read_csv(
            path,
            header=None,
            sep=";",
            encoding="latin1",
            skiprows=list(range(5, 13)),
            engine="c",
            low_memory=False,
            usecols=usecols,
        )
    except ValueError:
        raw = pd.read_csv(
            path,
            header=None,
            sep=";",
            encoding="latin1",
            skiprows=list(range(5, 13)),
            engine="python",
            usecols=usecols,
        )

    if raw.empty or len(raw) < 6:
        return pd.DataFrame(), {"error": "too_few_rows"}

    if usecols is None:
        raw.columns = header_cells_norm
    else:
        raw.columns = [header_cells_norm[i] for i in usecols]

    testnr_col = next((c for c in raw.columns if str(c).strip().upper() == "TEST NR"), None)
    if testnr_col is None:
        return pd.DataFrame(), {"error": "test_nr_col_not_found"}

    s = raw[testnr_col].astype(str).str.strip().str.upper()
    unit_rows = s[s == "UNIT"]
    if not unit_rows.empty:
        data_start = int(unit_rows.index[0]) + 1
    else:
        data_start = 5

    df = raw.iloc[data_start:].copy().dropna(how="all")
    if df.empty:
        return pd.DataFrame(), {"error": "no_chip_rows"}

    def _col(name: str) -> str | None:
        return next((c for c in df.columns if str(c).strip().upper() == name.upper()), None)

    vnr_col = _col("VNr")
    lot_col = _col("LOT")
    wafer_col = _col("WAFER")
    x_col = _col("X")
    y_col = _col("Y")

    def _as_int_str(series: pd.Series) -> pd.Series:
        s2 = pd.to_numeric(series, errors="coerce")
        return s2.map(lambda v: "" if pd.isna(v) else str(int(v)))

    parts = []
    if lot_col is not None:
        parts.append(df[lot_col].astype(str).str.strip())
    if wafer_col is not None:
        parts.append(df[wafer_col].astype(str).str.strip())
    if x_col is not None:
        parts.append(_as_int_str(df[x_col]))
    if y_col is not None:
        parts.append(_as_int_str(df[y_col]))
    if vnr_col is not None:
        parts.append("V" + _as_int_str(df[vnr_col]))

    if parts:
        chip_id = parts[0]
        for p in parts[1:]:
            chip_id = chip_id + ":" + p
        chip_id = chip_id.replace("::", ":")
        chip_id = chip_id.str.strip(":")
    else:
        chip_id = pd.Series([str(i) for i in range(len(df))], index=df.index)

    test_cols = [c for c in df.columns if str(c).strip().isdigit()]
    if not test_cols:
        return pd.DataFrame(), {"error": "test_cols_not_found"}

    values = df[test_cols].copy()
    for c in test_cols:
        values[c] = pd.to_numeric(values[c], errors="coerce")

    chip_matrix = values.copy()
    chip_matrix.index = chip_id.astype(str)
    chip_matrix.columns = chip_matrix.columns.astype(int)

    return chip_matrix, {"error": "", "n_tests": int(chip_matrix.shape[1]), "n_chips": int(chip_matrix.shape[0])}


def _safe_sheet_name(name: str, used: set[str]) -> str:
    safe = "".join(ch if ch.isalnum() or ch in " _-" else "_" for ch in str(name))
    safe = safe.strip() or "Sheet"
    safe = safe[:31]
    if safe not in used:
        used.add(safe)
        return safe

    base = safe[:28]
    for i in range(1, 1000):
        candidate = f"{base}_{i}"[:31]
        if candidate not in used:
            used.add(candidate)
            return candidate

    candidate = f"{base}_X"[:31]
    used.add(candidate)
    return candidate


def _extract_4ch_values(chip_matrix: pd.DataFrame, tests: list[int], prefix: str) -> tuple[pd.DataFrame, list[int]]:
    tests = list(tests)
    missing = [t for t in tests if int(t) not in chip_matrix.columns]

    cols = [int(t) for t in tests]
    sub = chip_matrix.reindex(columns=cols)

    out = pd.DataFrame(index=sub.index)
    for i in range(len(cols)):
        out[f"{prefix}_ch{i}"] = pd.to_numeric(sub.iloc[:, i], errors="coerce")
    return out, missing


def _compute_avg_vs_pair_residuals(
    *,
    chip_matrix: pd.DataFrame,
    case: CaseDef,
    scenario_label: str,
    tests: list[int],
    file_tag: str,
    scale_mv_per_unit: float,
) -> tuple[pd.DataFrame, dict]:
    vals, missing = _extract_4ch_values(chip_matrix, tests, prefix="v")

    df = vals.copy()
    df["v_mean"] = df[[f"v_ch{i}" for i in range(4)]].mean(axis=1, skipna=True)

    # Keep only chips where we have at least one channel-pair value
    has_any = df[[f"v_ch{i}" for i in range(4)]].notna().any(axis=1)
    df = df.loc[has_any].copy()

    # Residual/error when applying one averaged offset instead of per-pair offsets.
    # Under offset = Vref - Vmeas, error_i = offset_avg - offset_i = (V_i - mean(V)).
    for i in range(4):
        df[f"residual_ch{i}"] = df[f"v_ch{i}"] - df["v_mean"]
        df[f"residual_ch{i}_mV"] = df[f"residual_ch{i}"] * float(scale_mv_per_unit)
        df[f"abs_residual_ch{i}_mV"] = df[f"residual_ch{i}_mV"].abs()

    df["max_abs_residual_mV"] = df[[f"abs_residual_ch{i}_mV" for i in range(4)]].max(axis=1, skipna=True)

    df.insert(0, "chip_id", df.index.astype(str))
    df.insert(0, "scenario", scenario_label)
    df.insert(0, "case", case.name)
    df.insert(0, "corner", case.corner)
    df.insert(0, "supply", case.supply)
    df.insert(0, "file", file_tag)

    summary = {
        "file": file_tag,
        "case": case.name,
        "supply": case.supply,
        "corner": case.corner,
        "scenario": scenario_label,
        "n_chips": int(len(df)),
        "missing_tests": ",".join(map(str, missing)) if missing else "",
    }

    if len(df) > 0:
        summary.update(
            {
                "mean_max_abs_residual_mV": float(df["max_abs_residual_mV"].mean()),
                "p95_max_abs_residual_mV": float(df["max_abs_residual_mV"].quantile(0.95)),
                "p99_max_abs_residual_mV": float(df["max_abs_residual_mV"].quantile(0.99)),
                "max_max_abs_residual_mV": float(df["max_abs_residual_mV"].max()),
            }
        )

    return df, summary


def _build_pair_summary_rows(
    *,
    df_case: pd.DataFrame,
    file_tag: str,
    case: CaseDef,
    scenario_label: str,
) -> list[dict]:
    if df_case.empty:
        return []

    rows: list[dict] = []
    for i in range(4):
        col = f"residual_ch{i}_mV"
        abs_col = f"abs_residual_ch{i}_mV"
        series = df_case[col].dropna()
        abs_series = df_case[abs_col].dropna()
        if series.empty:
            continue
        rows.append(
            {
                "file": file_tag,
                "case": case.name,
                "supply": case.supply,
                "corner": case.corner,
                "scenario": scenario_label,
                "pair": f"ch{i}",
                "n_chips": int(series.shape[0]),
                "mean_residual_mV": float(series.mean()),
                "p50_residual_mV": float(series.quantile(0.50)),
                "p95_abs_residual_mV": float(abs_series.quantile(0.95)),
                "p99_abs_residual_mV": float(abs_series.quantile(0.99)),
                "max_abs_residual_mV": float(abs_series.max()),
            }
        )

    return rows


def _plot_residuals(
    *,
    df_case: pd.DataFrame,
    out_dir: Path,
    file_plot_key: str,
    case_name: str,
    scenario_label: str,
    acceptable_residual_mv: float,
) -> None:
    if df_case.empty:
        return

    plots_dir = out_dir / "plots" / str(file_plot_key) / str(case_name) / str(scenario_label)
    plots_dir.mkdir(parents=True, exist_ok=True)

    lim = float(acceptable_residual_mv)

    # Per-pair sorted residuals
    for i in range(4):
        s = df_case[f"residual_ch{i}_mV"].dropna().sort_values().reset_index(drop=True)
        if s.empty:
            continue

        fig = plt.figure(figsize=(7.8, 4.6))
        ax = fig.add_subplot(1, 1, 1)
        ax.plot(range(1, len(s) + 1), s.values, linewidth=1.1)
        ax.axhline(+lim, color="red", linestyle="--", linewidth=1, label=f"+{lim:.1f} mV")
        ax.axhline(-lim, color="red", linestyle="--", linewidth=1, label=f"-{lim:.1f} mV")
        ax.axhline(0.0, color="black", linewidth=1)
        ax.set_title(f"{case_name} {scenario_label} ch{i}: sorted residual (avg offset error) [mV]")
        ax.set_xlabel("Chip rank (sorted)")
        ax.set_ylabel("Residual (mV)")
        ax.grid(True, alpha=0.25)
        ax.legend(loc="best")
        fig.tight_layout()
        fig.savefig(plots_dir / f"ch{i}__sorted_residual_mV.png", dpi=170)
        plt.close(fig)

    # Sorted max-abs residual across the 4 pairs
    m = df_case["max_abs_residual_mV"].dropna().sort_values().reset_index(drop=True)
    if not m.empty:
        fig = plt.figure(figsize=(7.8, 4.6))
        ax = fig.add_subplot(1, 1, 1)
        ax.plot(range(1, len(m) + 1), m.values, linewidth=1.1)
        ax.axhline(lim, color="red", linestyle="--", linewidth=1, label=f"{lim:.1f} mV")
        ax.set_title(f"{case_name} {scenario_label}: sorted max |residual| across pairs [mV]")
        ax.set_xlabel("Chip rank (sorted)")
        ax.set_ylabel("Max |residual| (mV)")
        ax.grid(True, alpha=0.25)
        ax.legend(loc="best")
        fig.tight_layout()
        fig.savefig(plots_dir / "max__sorted_abs_residual_mV.png", dpi=170)
        plt.close(fig)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[3]

    parser = argparse.ArgumentParser(
        description="Assess avg compensation offset vs per-pair offsets using PROD pre-comp voltage measurements"
    )
    parser.add_argument(
        "--input-folder",
        default=str(repo_root / "PROD_Data"),
        help="Folder containing PROD raw CSVs.",
    )
    parser.add_argument(
        "--output-dir",
        default="",
        help="Output folder. If omitted, creates a dated folder next to this script.",
    )
    parser.add_argument("--glob", default="*.csv", help="File glob inside input folder.")
    parser.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Optional limit of processed files (0 = no limit).",
    )
    parser.add_argument(
        "--scale-mv-per-unit",
        type=float,
        default=DEFAULT_SCALE_MV_PER_UNIT,
        help="Scale factor applied to residuals to express them in mV (default assumes values are in V).",
    )
    parser.add_argument(
        "--acceptable-residual-mv",
        type=float,
        default=DEFAULT_ACCEPTABLE_RESIDUAL_MV,
        help="Acceptance threshold for |avg-offset error| in mV.",
    )
    parser.add_argument(
        "--scenario",
        choices=["s1", "s2", "both"],
        default="both",
        help="Which scenario test-set to analyze.",
    )
    parser.add_argument(
        "--no-chiplevel",
        action="store_true",
        help="Do not write the ChipLevel sheet (can be large).",
    )

    args = parser.parse_args()

    input_folder = Path(args.input_folder)
    if not input_folder.is_dir():
        raise SystemExit(f"Input folder not found: {input_folder}")

    if args.output_dir:
        out_dir = Path(args.output_dir)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = Path(__file__).resolve().parent / f"output_{stamp}_avg_vs_pair_offset"
    out_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(input_folder.glob(args.glob))
    if args.max_files and args.max_files > 0:
        files = files[: args.max_files]
    if not files:
        raise SystemExit(f"No files matched {args.glob} in: {input_folder}")

    tests_needed = sorted({t for c in CASES for t in (c.scenario1_tests + c.scenario2_tests)})

    chip_rows: list[pd.DataFrame] = []
    summary_rows: list[dict] = []
    pair_summary_rows: list[dict] = []
    parse_fail_rows: list[dict] = []

    lim = float(args.acceptable_residual_mv)

    for file_index, p in enumerate(files, start=1):
        file_tag = p.name
        file_plot_key = f"{file_index:02d}_{p.stem[:20]}"  # short for Windows path limits

        chip_matrix, info = _read_prod_csv_chip_matrix(p, tests_needed=set(tests_needed))
        if info.get("error"):
            parse_fail_rows.append({"file": file_tag, **info})
            continue

        chip_matrix = chip_matrix.reindex(columns=[t for t in tests_needed if t in chip_matrix.columns])

        for case in CASES:
            scenarios: list[tuple[str, list[int]]] = []
            if args.scenario in ("s1", "both"):
                scenarios.append(("scenario1", case.scenario1_tests))
            if args.scenario in ("s2", "both"):
                scenarios.append(("scenario2", case.scenario2_tests))

            for scenario_label, tests in scenarios:
                df_case, summary = _compute_avg_vs_pair_residuals(
                    chip_matrix=chip_matrix,
                    case=case,
                    scenario_label=scenario_label,
                    tests=tests,
                    file_tag=file_tag,
                    scale_mv_per_unit=float(args.scale_mv_per_unit),
                )

                # Decision on p99 of per-chip max |residual| across the 4 pairs
                if not df_case.empty:
                    p99 = float(df_case["max_abs_residual_mV"].quantile(0.99))
                    ok = p99 <= lim
                    summary.update(
                        {
                            "decision_rule": f"PASS if p99_max_abs_residual_mV <= {lim:g} mV",
                            "decision": "PASS" if ok else "FAIL",
                            "decision_reason": "OK" if ok else f"p99_max>{lim:g}mV",
                        }
                    )

                chip_rows.append(df_case)
                summary_rows.append(summary)
                pair_summary_rows.extend(
                    _build_pair_summary_rows(
                        df_case=df_case,
                        file_tag=file_tag,
                        case=case,
                        scenario_label=scenario_label,
                    )
                )

                _plot_residuals(
                    df_case=df_case,
                    out_dir=out_dir,
                    file_plot_key=file_plot_key,
                    case_name=case.name,
                    scenario_label=scenario_label,
                    acceptable_residual_mv=float(args.acceptable_residual_mv),
                )

    chip_df = pd.concat(chip_rows, ignore_index=True) if chip_rows else pd.DataFrame()
    summary_df = pd.DataFrame(summary_rows)
    pair_summary_df = pd.DataFrame(pair_summary_rows)
    parse_fail_df = pd.DataFrame(parse_fail_rows)

    report_path = out_dir / "tx_supply_compensation_avg_vs_pair_offset.xlsx"
    used_sheet_names: set[str] = set()
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name=_safe_sheet_name("Summary_Cases", used_sheet_names), index=False)
        pair_summary_df.to_excel(writer, sheet_name=_safe_sheet_name("Summary_Pairs", used_sheet_names), index=False)

        if not summary_df.empty and "file" in summary_df.columns:
            for file_name in sorted(summary_df["file"].dropna().unique()):
                sheet = _safe_sheet_name(Path(file_name).stem, used_sheet_names)
                df_case_one = summary_df.loc[summary_df["file"] == file_name].copy()
                df_pair_one = (
                    pair_summary_df.loc[pair_summary_df["file"] == file_name].copy()
                    if not pair_summary_df.empty
                    else pd.DataFrame()
                )

                startrow = 0
                df_case_one.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow)
                startrow += len(df_case_one) + 3
                if not df_pair_one.empty:
                    df_pair_one.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow)

        if not args.no_chiplevel and not chip_df.empty:
            chip_df.to_excel(writer, sheet_name=_safe_sheet_name("ChipLevel", used_sheet_names), index=False)

        if not parse_fail_df.empty:
            parse_fail_df.to_excel(writer, sheet_name=_safe_sheet_name("ParseFailures", used_sheet_names), index=False)

    print(f"Processed files: {len(files)}")
    print(f"Excel report:     {report_path}")
    print(f"Decision rule:    PASS if p99_max_abs_residual_mV <= {lim:g} mV")
    if not parse_fail_df.empty:
        print("Parse failures:   included in Excel report")
    print(f"Plots folder:     {out_dir / 'plots'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
