"""\
TX supply IR-loss compensation scenario comparison.

Goal
----
Compare the pre-offset (raw) measurements between two scenarios:
  - Scenario 1: compensation measured after TX power calibration
  - Scenario 2: compensation measured without TX power calibration

For each PROD raw-data CSV in an input folder, this script extracts chip-level
values for the specified test numbers and produces:
    - A single Excel report (.xlsx) with:
            - global summaries (case-level + per-channel-pair)
            - one summary sheet per input file
            - an optional chip-level sheet (can be large)
    - Plots per file / case / channel showing sorted distributions:
            - scenario1 vs scenario2 values
            - delta (scenario2 - scenario1) scaled to mV

Assumptions
-----------
Raw PROD CSV structure is the standard semi-colon delimited wide format where:
    - Row 1 contains meta headers + many numeric test-number columns
    - Rows 2..5 contain Test Name / Low / High / Unit
    - Rows 6..13 contain summary stats (Min/Max/Mean/...) and must be skipped
    - Rows 14..end are per-chip rows (chips are rows; tests are columns)

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


DEFAULT_DELTA_SCALE_MV_PER_UNIT = 1000.0
DEFAULT_ACCEPTABLE_DELTA_MV = 10.0


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

    In this format:
      - Row 1 contains meta headers + many numeric test-number columns
      - Rows 2..5 contain Test Name / Low / High / Unit
      - Rows 6..13 contain summary stats (Min/Max/Mean/...) and must be skipped
      - Rows 14..end are per-chip rows

    chip_matrix:
        index   -> chip_id (LOT/WAFER/X/Y/VNr composed when possible)
        columns -> test number (int)
        values  -> measurement (float)
    """

    # Fast sanity check: ensure we are looking at the right kind of file.
    header_row = _find_header_row_index(path, needle="Test Nr")
    if header_row is None:
        return pd.DataFrame(), {"error": "header_not_found"}
    if header_row != 0:
        # This script currently supports the common export where the header is at the first line.
        # If future files have a preamble, we can extend this to skip preceding lines.
        return pd.DataFrame(), {"error": "header_not_first_line", "header_row": header_row}

    # Parse header + test-name row (kept for future extensions; we primarily need the header).
    def _parse_first_two_lines_csv(p: Path):
        with p.open("r", encoding="latin1", errors="ignore") as f:
            line1 = f.readline().rstrip("\n\r")
            line2 = f.readline().rstrip("\n\r")
        header_cells = line1.split(";")
        test_name_cells = line2.split(";")
        if len(test_name_cells) < len(header_cells):
            test_name_cells = test_name_cells + [""] * (len(header_cells) - len(test_name_cells))
        return header_cells, test_name_cells

    header_cells, _test_name_cells = _parse_first_two_lines_csv(path)

    def _normalize_header_cell(value: str) -> str:
        return str(value).strip()

    header_cells_norm = [_normalize_header_cell(v) for v in header_cells]

    # Read with header=None so we can keep the explicit header row as data for alignment,
    # then set proper columns ourselves.
    # Skip rows 6-13 (1-based) => indices 5-12 (0-based)

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

        # If something went wrong, fall back to reading all columns.
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
        # Fallback for rare cases where the C engine fails.
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

    # Identify where the per-chip data starts: after the UNIT row (row label in "Test Nr" column).
    testnr_col = next((c for c in raw.columns if str(c).strip().upper() == "TEST NR"), None)
    if testnr_col is None:
        return pd.DataFrame(), {"error": "test_nr_col_not_found"}

    s = raw[testnr_col].astype(str).str.strip().str.upper()
    unit_rows = s[s == "UNIT"]
    if not unit_rows.empty:
        data_start = int(unit_rows.index[0]) + 1
    else:
        # Fallback: after the first 5 metadata rows
        data_start = 5

    df = raw.iloc[data_start:].copy().dropna(how="all")
    if df.empty:
        return pd.DataFrame(), {"error": "no_chip_rows"}

    # Build chip_id from common meta columns.
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

    # Test columns are digit-only headers (test numbers)
    test_cols = [c for c in df.columns if str(c).strip().isdigit()]
    if not test_cols:
        return pd.DataFrame(), {"error": "test_cols_not_found"}

    values = df[test_cols].copy()
    for c in test_cols:
        values[c] = pd.to_numeric(values[c], errors="coerce")

    chip_matrix = values.copy()
    chip_matrix.index = chip_id.astype(str)
    chip_matrix.columns = chip_matrix.columns.astype(int)

    return chip_matrix, {
        "error": "",
        "n_tests": int(chip_matrix.shape[1]),
        "n_chips": int(chip_matrix.shape[0]),
    }


def _extract_case_values(
    chip_matrix: pd.DataFrame,
    tests: list[int],
    prefix: str,
) -> tuple[pd.DataFrame, list[int]]:
    """Extract 4-channel values for a given test list.

    Returns (df, missing_tests).
    df columns are: {prefix}_ch0..ch3
    """
    tests = list(tests)
    missing = [t for t in tests if int(t) not in chip_matrix.columns]

    # Preserve order; if missing, column will be all-NaN.
    cols = [int(t) for t in tests]
    sub = chip_matrix.reindex(columns=cols)

    out = pd.DataFrame(index=sub.index)
    for i in range(len(cols)):
        out[f"{prefix}_ch{i}"] = pd.to_numeric(sub.iloc[:, i], errors="coerce")
    return out, missing


def _compute_case_comparison(
    *,
    chip_matrix: pd.DataFrame,
    case: CaseDef,
    file_tag: str,
    delta_scale_mv_per_unit: float,
    acceptable_delta_mv: float,
) -> tuple[pd.DataFrame, dict]:
    s1, miss1 = _extract_case_values(chip_matrix, case.scenario1_tests, "s1")
    s2, miss2 = _extract_case_values(chip_matrix, case.scenario2_tests, "s2")

    df = pd.concat([s1, s2], axis=1)

    # Means across channels (ignore NaNs)
    df["s1_mean"] = df[[f"s1_ch{i}" for i in range(4)]].mean(axis=1, skipna=True)
    df["s2_mean"] = df[[f"s2_ch{i}" for i in range(4)]].mean(axis=1, skipna=True)

    # Per-channel deltas
    for i in range(4):
        df[f"delta_ch{i}"] = df[f"s2_ch{i}"] - df[f"s1_ch{i}"]
        df[f"abs_delta_ch{i}"] = (df[f"delta_ch{i}"]).abs()
        df[f"delta_ch{i}_mV"] = df[f"delta_ch{i}"] * float(delta_scale_mv_per_unit)
        df[f"abs_delta_ch{i}_mV"] = df[f"abs_delta_ch{i}"] * float(delta_scale_mv_per_unit)

    df["delta_mean"] = df["s2_mean"] - df["s1_mean"]
    df["abs_delta_mean"] = df["delta_mean"].abs()
    df["delta_mean_mV"] = df["delta_mean"] * float(delta_scale_mv_per_unit)
    df["abs_delta_mean_mV"] = df["abs_delta_mean"] * float(delta_scale_mv_per_unit)

    # Keep only chips where we have at least one channel value in both scenarios
    has_s1 = df[[f"s1_ch{i}" for i in range(4)]].notna().any(axis=1)
    has_s2 = df[[f"s2_ch{i}" for i in range(4)]].notna().any(axis=1)
    df = df.loc[has_s1 & has_s2].copy()

    df.insert(0, "chip_id", df.index.astype(str))
    df.insert(0, "case", case.name)
    df.insert(0, "corner", case.corner)
    df.insert(0, "supply", case.supply)
    df.insert(0, "file", file_tag)

    summary = {
        "file": file_tag,
        "case": case.name,
        "supply": case.supply,
        "corner": case.corner,
        "n_chips": int(len(df)),
        "missing_s1_tests": ",".join(map(str, miss1)) if miss1 else "",
        "missing_s2_tests": ",".join(map(str, miss2)) if miss2 else "",
    }

    if len(df) > 0:
        limit = float(acceptable_delta_mv)
        p95_abs_delta_mean_mV = float(df["abs_delta_mean_mV"].quantile(0.95))
        p99_abs_delta_mean_mV = float(df["abs_delta_mean_mV"].quantile(0.99))
        decision_ok = p99_abs_delta_mean_mV <= limit
        summary.update(
            {
                "mean_abs_delta_mean": float(df["abs_delta_mean"].mean()),
                "p95_abs_delta_mean": float(df["abs_delta_mean"].quantile(0.95)),
                "p99_abs_delta_mean": float(df["abs_delta_mean"].quantile(0.99)),
                "max_abs_delta_mean": float(df["abs_delta_mean"].max()),
                "corr_s1_s2_mean": float(df[["s1_mean", "s2_mean"]].corr().iloc[0, 1]),
                "mean_abs_delta_mean_mV": float(df["abs_delta_mean_mV"].mean()),
                "p95_abs_delta_mean_mV": p95_abs_delta_mean_mV,
                "p99_abs_delta_mean_mV": p99_abs_delta_mean_mV,
                "max_abs_delta_mean_mV": float(df["abs_delta_mean_mV"].max()),
                "decision_rule": f"PASS if p99_abs_delta_mean_mV <= {limit:g} mV",
                "decision": "PASS" if decision_ok else "FAIL",
                "decision_reason": "OK" if decision_ok else f"p99_mean>{limit:g}mV",
            }
        )
        # Note: per-channel-pair summary is emitted in a separate long-format table.

    return df, summary


def _build_pair_summary_rows(
    *,
    df_case: pd.DataFrame,
    case: CaseDef,
    file_tag: str,
) -> list[dict]:
    """Return long-format (row-based) summary rows for mean + each channel."""

    if df_case.empty:
        return []

    rows: list[dict] = []

    # Mean row
    rows.append(
        {
            "file": file_tag,
            "case": case.name,
            "supply": case.supply,
            "corner": case.corner,
            "pair": "mean",
            "n_chips": int(len(df_case)),
            "mean_abs_delta_mV": float(df_case["abs_delta_mean_mV"].mean()),
            "p95_abs_delta_mV": float(df_case["abs_delta_mean_mV"].quantile(0.95)),
            "p99_abs_delta_mV": float(df_case["abs_delta_mean_mV"].quantile(0.99)),
            "max_abs_delta_mV": float(df_case["abs_delta_mean_mV"].max()),
            "corr_s1_s2": float(df_case[["s1_mean", "s2_mean"]].corr().iloc[0, 1]),
        }
    )

    for i in range(4):
        rows.append(
            {
                "file": file_tag,
                "case": case.name,
                "supply": case.supply,
                "corner": case.corner,
                "pair": f"ch{i}",
                "n_chips": int(df_case[f"abs_delta_ch{i}_mV"].notna().sum()),
                "mean_abs_delta_mV": float(df_case[f"abs_delta_ch{i}_mV"].mean()),
                "p95_abs_delta_mV": float(df_case[f"abs_delta_ch{i}_mV"].quantile(0.95)),
                "p99_abs_delta_mV": float(df_case[f"abs_delta_ch{i}_mV"].quantile(0.99)),
                "max_abs_delta_mV": float(df_case[f"abs_delta_ch{i}_mV"].max()),
                "corr_s1_s2": float(df_case[[f"s1_ch{i}", f"s2_ch{i}"]].corr().iloc[0, 1]),
            }
        )

    return rows


def _safe_sheet_name(name: str, used: set[str]) -> str:
    """Excel sheet names must be <=31 chars and unique within a workbook."""

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


def _plot_case_distributions(
    *,
    df_case: pd.DataFrame,
    case: CaseDef,
    out_dir: Path,
    file_plot_key: str,
    acceptable_delta_mv: float,
) -> None:
    if df_case.empty:
        return

    plots_dir = out_dir / "plots" / str(file_plot_key) / case.name
    plots_dir.mkdir(parents=True, exist_ok=True)

    # Per-channel sorted distributions
    for i in range(4):
        s1 = df_case[f"s1_ch{i}"].dropna().sort_values().reset_index(drop=True)
        s2 = df_case[f"s2_ch{i}"].dropna().sort_values().reset_index(drop=True)

        fig = plt.figure(figsize=(7.8, 4.6))
        ax = fig.add_subplot(1, 1, 1)
        if not s1.empty:
            ax.plot(range(1, len(s1) + 1), s1.values, linewidth=1.1, label="scenario1")
        if not s2.empty:
            ax.plot(range(1, len(s2) + 1), s2.values, linewidth=1.1, label="scenario2")
        ax.set_title(f"{case.name} ch{i}: sorted distribution")
        ax.set_xlabel("Chip rank (sorted)")
        ax.set_ylabel("Value")
        ax.grid(True, alpha=0.25)
        ax.legend(loc="best")
        fig.tight_layout()
        fig.savefig(plots_dir / f"ch{i}__sorted_s1_vs_s2.png", dpi=170)
        plt.close(fig)

        d = df_case[f"delta_ch{i}_mV"].dropna().sort_values().reset_index(drop=True)
        fig = plt.figure(figsize=(7.8, 4.6))
        ax = fig.add_subplot(1, 1, 1)
        if not d.empty:
            ax.plot(range(1, len(d) + 1), d.values, linewidth=1.1, color="#1f77b4")
        lim = float(acceptable_delta_mv)
        ax.axhline(+lim, color="red", linestyle="--", linewidth=1, label=f"+{lim:.1f} mV")
        ax.axhline(-lim, color="red", linestyle="--", linewidth=1, label=f"-{lim:.1f} mV")
        ax.axhline(0.0, color="black", linewidth=1)
        ax.set_title(f"{case.name} ch{i}: sorted delta (scenario2 - scenario1) [mV]")
        ax.set_xlabel("Chip rank (sorted)")
        ax.set_ylabel("Delta (mV)")
        ax.grid(True, alpha=0.25)
        ax.legend(loc="best")
        fig.tight_layout()
        fig.savefig(plots_dir / f"ch{i}__sorted_delta_mV.png", dpi=170)
        plt.close(fig)

    dmean = df_case["delta_mean_mV"].dropna().sort_values().reset_index(drop=True)
    fig = plt.figure(figsize=(7.8, 4.6))
    ax = fig.add_subplot(1, 1, 1)
    if not dmean.empty:
        ax.plot(range(1, len(dmean) + 1), dmean.values, linewidth=1.1)
    lim = float(acceptable_delta_mv)
    ax.axhline(+lim, color="red", linestyle="--", linewidth=1, label=f"+{lim:.1f} mV")
    ax.axhline(-lim, color="red", linestyle="--", linewidth=1, label=f"-{lim:.1f} mV")
    ax.axhline(0.0, color="black", linewidth=1)
    ax.set_title(f"{case.name}: sorted delta mean (scenario2 - scenario1) [mV]")
    ax.set_xlabel("Chip rank (sorted)")
    ax.set_ylabel("Delta mean (mV)")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="best")
    fig.tight_layout()
    fig.savefig(plots_dir / "mean__sorted_delta_mV.png", dpi=170)
    plt.close(fig)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[3]

    parser = argparse.ArgumentParser(description="Compare TX supply compensation scenario 1 vs 2 on PROD raw data")
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
    parser.add_argument(
        "--glob",
        default="*.csv",
        help="File glob inside input folder.",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Optional limit of processed files (0 = no limit).",
    )

    parser.add_argument(
        "--delta-scale-mv-per-unit",
        type=float,
        default=DEFAULT_DELTA_SCALE_MV_PER_UNIT,
        help="Scale factor applied to deltas to express them in mV (default assumes values are in V).",
    )
    parser.add_argument(
        "--acceptable-delta-mv",
        type=float,
        default=DEFAULT_ACCEPTABLE_DELTA_MV,
        help="Acceptance threshold for abs(delta) in mV.",
    )

    args = parser.parse_args()

    input_folder = Path(args.input_folder)
    if not input_folder.is_dir():
        raise SystemExit(f"Input folder not found: {input_folder}")

    if args.output_dir:
        out_dir = Path(args.output_dir)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = Path(__file__).resolve().parent / f"output_{stamp}_scenario_compare"
    out_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(input_folder.glob(args.glob))
    if args.max_files and args.max_files > 0:
        files = files[: args.max_files]

    if not files:
        raise SystemExit(f"No files matched {args.glob} in: {input_folder}")

    per_chip_rows: list[pd.DataFrame] = []
    per_file_summary_rows: list[dict] = []
    per_pair_summary_rows: list[dict] = []
    parse_fail_rows: list[dict] = []

    tests_needed = sorted({t for c in CASES for t in (c.scenario1_tests + c.scenario2_tests)})

    for file_index, p in enumerate(files, start=1):
        file_tag = p.name
        # Keep plot paths short (Windows path length constraints)
        file_plot_key = f"{file_index:02d}_{p.stem[:20]}"
        chip_matrix, info = _read_prod_csv_chip_matrix(p, tests_needed=set(tests_needed))
        if info.get("error"):
            parse_fail_rows.append({"file": file_tag, **info})
            continue

        # Subset to required tests early (saves memory and speeds computations)
        chip_matrix = chip_matrix.reindex(columns=[t for t in tests_needed if t in chip_matrix.columns])

        for case in CASES:
            df_case, summary = _compute_case_comparison(
                chip_matrix=chip_matrix,
                case=case,
                file_tag=file_tag,
                delta_scale_mv_per_unit=float(args.delta_scale_mv_per_unit),
                acceptable_delta_mv=float(args.acceptable_delta_mv),
            )
            per_chip_rows.append(df_case)
            per_file_summary_rows.append(summary)
            per_pair_summary_rows.extend(_build_pair_summary_rows(df_case=df_case, case=case, file_tag=file_tag))

            _plot_case_distributions(
                df_case=df_case,
                case=case,
                out_dir=out_dir,
                file_plot_key=file_plot_key,
                acceptable_delta_mv=float(args.acceptable_delta_mv),
            )

    per_chip = pd.concat(per_chip_rows, ignore_index=True) if per_chip_rows else pd.DataFrame()
    per_file_summary = pd.DataFrame(per_file_summary_rows)
    per_pair_summary = pd.DataFrame(per_pair_summary_rows)
    parse_fail = pd.DataFrame(parse_fail_rows)

    report_path = out_dir / "tx_supply_compensation_scenario_compare.xlsx"
    used_sheet_names: set[str] = set()
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        per_file_summary.to_excel(writer, sheet_name=_safe_sheet_name("Summary_Cases", used_sheet_names), index=False)
        per_pair_summary.to_excel(writer, sheet_name=_safe_sheet_name("Summary_Pairs", used_sheet_names), index=False)

        # Per-file sheets: case summary table + pair summary table
        if not per_file_summary.empty and "file" in per_file_summary.columns:
            for file_name in sorted(per_file_summary["file"].dropna().unique()):
                sheet = _safe_sheet_name(Path(file_name).stem, used_sheet_names)
                df_case_one = per_file_summary.loc[per_file_summary["file"] == file_name].copy()
                df_pair_one = per_pair_summary.loc[per_pair_summary["file"] == file_name].copy() if not per_pair_summary.empty else pd.DataFrame()

                startrow = 0
                df_case_one.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow)
                startrow += len(df_case_one) + 3

                if not df_pair_one.empty:
                    df_pair_one.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow)

        if not per_chip.empty:
            per_chip.to_excel(writer, sheet_name=_safe_sheet_name("ChipLevel", used_sheet_names), index=False)

        if not parse_fail.empty:
            parse_fail.to_excel(writer, sheet_name=_safe_sheet_name("ParseFailures", used_sheet_names), index=False)

    print(f"Processed files: {len(files)}")
    print(f"Excel report:     {report_path}")
    if not parse_fail.empty:
        print("Parse failures:   included in Excel report")
    print(f"Plots folder:     {out_dir / 'plots'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
