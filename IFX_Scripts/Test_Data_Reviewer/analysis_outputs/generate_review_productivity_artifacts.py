import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


IMPORTANT_TESTS = 1000
ENGINEER_DAY_HOURS = 8
DETECTION_CONSERVATIVE = 0.8501
DETECTION_AUTOMATED = 0.9735
DEFAULT_FE_INSERTIONS = 3

FULL_MANUAL_SECONDS_PER_TEST = 14
FULL_MANUAL_PROBLEMATIC_TEST_SECONDS = 80
FULL_MANUAL_FIXED_OVERHEAD_MINUTES = 25

CONSERVATIVE_LOCAL_COPY_MINUTES = 10
CONSERVATIVE_WRAP_UP_MINUTES = 15
CONSERVATIVE_SECONDS_PER_REVIEWED_TEST = 3
CONSERVATIVE_REVIEWED_TESTS_MIN = 700
CONSERVATIVE_REVIEWED_TESTS_NOMINAL = 800
CONSERVATIVE_REVIEWED_TESTS_MAX = 900
CONSERVATIVE_PROBLEMATIC_TEST_MIN_SECONDS = 30
CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS = 60
CONSERVATIVE_PROBLEMATIC_TEST_MAX_SECONDS = 120

AUTOMATED_LOCAL_COPY_MINUTES = 25
AUTOMATED_WRAP_UP_MINUTES = 15
AUTOMATED_PROBLEMATIC_TEST_MIN_SECONDS = 15
AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS = 30
AUTOMATED_PROBLEMATIC_TEST_MAX_SECONDS = 60

MONTE_CARLO_SAMPLES = 5000
MONTE_CARLO_SEED = 42

SCENARIO_ORDER = ("Best", "Nominal", "Worst")
SCENARIO_COLORS = {
    "Conservative reviewer": "#b14623",
    "Automated review": "#0f6cbd",
}

CONSERVATIVE_SCENARIOS = {
    "Best": {
        "reviewed_tests": CONSERVATIVE_REVIEWED_TESTS_MIN,
        "problematic_seconds": CONSERVATIVE_PROBLEMATIC_TEST_MIN_SECONDS,
    },
    "Nominal": {
        "reviewed_tests": CONSERVATIVE_REVIEWED_TESTS_NOMINAL,
        "problematic_seconds": CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS,
    },
    "Worst": {
        "reviewed_tests": CONSERVATIVE_REVIEWED_TESTS_MAX,
        "problematic_seconds": CONSERVATIVE_PROBLEMATIC_TEST_MAX_SECONDS,
    },
}

AUTOMATED_SCENARIOS = {
    "Best": {"problematic_seconds": AUTOMATED_PROBLEMATIC_TEST_MIN_SECONDS},
    "Nominal": {"problematic_seconds": AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS},
    "Worst": {"problematic_seconds": AUTOMATED_PROBLEMATIC_TEST_MAX_SECONDS},
}


def build_assumptions_df(fe_insertions: int) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Model": "Shared",
                "Assumption": "FE insertions reviewed",
                "Value": fe_insertions,
                "Unit": "insertions/review",
                "Notes": "Explicit input parameter. Insertion-sensitive review effort scales with this count.",
            },
            {
                "Model": "Conservative reviewer",
                "Assumption": "Local data copy and load overhead",
                "Value": CONSERVATIVE_LOCAL_COPY_MINUTES,
                "Unit": "min/review",
                "Notes": "User-provided fixed overhead before review.",
            },
            {
                "Model": "Conservative reviewer",
                "Assumption": "Key parameter tests reviewed",
                "Value": f"{CONSERVATIVE_REVIEWED_TESTS_MIN}-{CONSERVATIVE_REVIEWED_TESTS_NOMINAL}-{CONSERVATIVE_REVIEWED_TESTS_MAX}",
                "Unit": "tests/review",
                "Notes": "Best-nominal-worst sensitivity band centered on the user's in-the-range-of-800 statement.",
            },
            {
                "Model": "Conservative reviewer",
                "Assumption": "Baseline scan time per reviewed test",
                "Value": CONSERVATIVE_SECONDS_PER_REVIEWED_TEST,
                "Unit": "s/test",
                "Notes": "Nominal scan effort for each reviewed key parameter test.",
            },
            {
                "Model": "Conservative reviewer",
                "Assumption": "Problematic test investigation time",
                "Value": f"{CONSERVATIVE_PROBLEMATIC_TEST_MIN_SECONDS}-{CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS}-{CONSERVATIVE_PROBLEMATIC_TEST_MAX_SECONDS}",
                "Unit": "s/problematic test",
                "Notes": "User-provided 1 to 3 min range, expressed as best-nominal-worst.",
            },
            {
                "Model": "Conservative reviewer",
                "Assumption": "Wrap-up and reporting overhead",
                "Value": CONSERVATIVE_WRAP_UP_MINUTES,
                "Unit": "min/review",
                "Notes": "User-provided fixed overhead after review.",
            },
            {
                "Model": "Automated review",
                "Assumption": "Local data copy overhead",
                "Value": AUTOMATED_LOCAL_COPY_MINUTES,
                "Unit": "min/review",
                "Notes": "User-provided fixed overhead before running the automated flow.",
            },
            {
                "Model": "Automated review",
                "Assumption": "Problematic test handling time",
                "Value": f"{AUTOMATED_PROBLEMATIC_TEST_MIN_SECONDS}-{AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS}-{AUTOMATED_PROBLEMATIC_TEST_MAX_SECONDS}",
                "Unit": "s/problematic test",
                "Notes": "20 s nominal from the user with a narrow best-worst sensitivity band for scenario and Monte Carlo analysis.",
            },
            {
                "Model": "Automated review",
                "Assumption": "Wrap-up and reporting overhead",
                "Value": AUTOMATED_WRAP_UP_MINUTES,
                "Unit": "min/review",
                "Notes": "User-provided fixed overhead after review.",
            },
            {
                "Model": "Monte Carlo",
                "Assumption": "Distribution form",
                "Value": "triangular",
                "Unit": "distribution",
                "Notes": "Uses left-mode-right sampling for uncertain timing inputs.",
            },
            {
                "Model": "Monte Carlo",
                "Assumption": "Sample count",
                "Value": MONTE_CARLO_SAMPLES,
                "Unit": "samples per problematic-test point",
                "Notes": "Higher sample count stabilizes the uncertainty bands.",
            },
            {
                "Model": "Shared",
                "Assumption": "Important tests in scope",
                "Value": IMPORTANT_TESTS,
                "Unit": "tests",
                "Notes": "Total important tests considered for each scenario.",
            },
            {
                "Model": "Shared",
                "Assumption": "Engineer workday",
                "Value": ENGINEER_DAY_HOURS,
                "Unit": "h/day",
                "Notes": "Used to convert total review time into reviews per 8 h day.",
            },
            {
                "Model": "Shared",
                "Assumption": "Conservative detection rate",
                "Value": DETECTION_CONSERVATIVE * 100.0,
                "Unit": "%",
                "Notes": "Detection-rate assumption preserved from the prior model.",
            },
            {
                "Model": "Shared",
                "Assumption": "Automated detection rate",
                "Value": DETECTION_AUTOMATED * 100.0,
                "Unit": "%",
                "Notes": "Detection-rate assumption preserved from the prior model.",
            },
        ]
    )


def compute_full_manual_hours(issue_tests: float) -> float:
    total_seconds = (
        IMPORTANT_TESTS * FULL_MANUAL_SECONDS_PER_TEST
        + issue_tests * FULL_MANUAL_PROBLEMATIC_TEST_SECONDS
        + FULL_MANUAL_FIXED_OVERHEAD_MINUTES * 60
    )
    return total_seconds / 3600


def compute_conservative_hours(
    issue_tests: float,
    reviewed_tests: float,
    problematic_seconds: float,
    fe_insertions: float,
) -> float:
    total_seconds = (
        CONSERVATIVE_LOCAL_COPY_MINUTES * 60
        + fe_insertions * reviewed_tests * CONSERVATIVE_SECONDS_PER_REVIEWED_TEST
        + fe_insertions * issue_tests * problematic_seconds
        + CONSERVATIVE_WRAP_UP_MINUTES * 60
    )
    return total_seconds / 3600


def compute_automated_hours(issue_tests: float, problematic_seconds: float, fe_insertions: float) -> float:
    total_seconds = (
        AUTOMATED_LOCAL_COPY_MINUTES * 60
        + fe_insertions * issue_tests * problematic_seconds
        + AUTOMATED_WRAP_UP_MINUTES * 60
    )
    return total_seconds / 3600


def build_nominal_summary_df(problematic_pcts: tuple[float, ...], fe_insertions: int) -> pd.DataFrame:
    rows: list[dict[str, float]] = []
    for problematic_pct in problematic_pcts:
        issue_tests = IMPORTANT_TESTS * (problematic_pct / 100.0)
        full_manual_hours = compute_full_manual_hours(issue_tests)
        conservative_hours = compute_conservative_hours(
            issue_tests,
            CONSERVATIVE_REVIEWED_TESTS_NOMINAL,
            CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS,
            fe_insertions,
        )
        automated_hours = compute_automated_hours(
            issue_tests,
            AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS,
            fe_insertions,
        )
        throughput_gain_vs_conservative = conservative_hours / automated_hours
        throughput_gain_vs_full_manual = full_manual_hours / automated_hours
        rows.append(
            {
                "Problematic Tests (%)": problematic_pct,
                "Problematic Tests Count": issue_tests,
                "Full Manual Review Time (h)": full_manual_hours,
                "Conservative Review Time (h)": conservative_hours,
                "Automated Review Time (h)": automated_hours,
                "Full Manual Productivity (reviews/8h day)": ENGINEER_DAY_HOURS / full_manual_hours,
                "Conservative Productivity (reviews/8h day)": ENGINEER_DAY_HOURS / conservative_hours,
                "Automated Productivity (reviews/8h day)": ENGINEER_DAY_HOURS / automated_hours,
                "Automation Time Saved vs Full Manual (h)": full_manual_hours - automated_hours,
                "Automation Time Saved vs Conservative (h)": conservative_hours - automated_hours,
                "Automation Throughput Gain vs Full Manual (x)": throughput_gain_vs_full_manual,
                "Automation Throughput Gain vs Conservative (x)": throughput_gain_vs_conservative,
                "Automation Productivity Improvement vs Full Manual (%)": (throughput_gain_vs_full_manual - 1.0) * 100.0,
                "Automation Productivity Improvement vs Conservative (%)": (throughput_gain_vs_conservative - 1.0) * 100.0,
                "Conservative Detection (%)": DETECTION_CONSERVATIVE * 100.0,
                "Automated Detection (%)": DETECTION_AUTOMATED * 100.0,
                "Conservative Issues Found": issue_tests * DETECTION_CONSERVATIVE,
                "Automated Issues Found": issue_tests * DETECTION_AUTOMATED,
                "Conservative Issues Missed": issue_tests * (1.0 - DETECTION_CONSERVATIVE),
                "Automated Issues Missed": issue_tests * (1.0 - DETECTION_AUTOMATED),
            }
        )
    return pd.DataFrame(rows)


def build_scenario_curve_df(problematic_pcts: range, fe_insertions: int) -> pd.DataFrame:
    rows: list[dict[str, float | str]] = []
    for problematic_pct in problematic_pcts:
        issue_tests = IMPORTANT_TESTS * (problematic_pct / 100.0)
        for scenario_name in SCENARIO_ORDER:
            conservative_hours = compute_conservative_hours(
                issue_tests,
                CONSERVATIVE_SCENARIOS[scenario_name]["reviewed_tests"],
                CONSERVATIVE_SCENARIOS[scenario_name]["problematic_seconds"],
                fe_insertions,
            )
            automated_hours = compute_automated_hours(
                issue_tests,
                AUTOMATED_SCENARIOS[scenario_name]["problematic_seconds"],
                fe_insertions,
            )
            rows.append(
                {
                    "Problematic Tests (%)": problematic_pct,
                    "Problematic Tests Count": issue_tests,
                    "Scenario": scenario_name,
                    "Conservative Reviewed Tests": CONSERVATIVE_SCENARIOS[scenario_name]["reviewed_tests"],
                    "Conservative Problematic Time (s/test)": CONSERVATIVE_SCENARIOS[scenario_name]["problematic_seconds"],
                    "Automated Problematic Time (s/test)": AUTOMATED_SCENARIOS[scenario_name]["problematic_seconds"],
                    "Conservative Review Time (h)": conservative_hours,
                    "Automated Review Time (h)": automated_hours,
                    "Automation Productivity Improvement vs Conservative (%)": (conservative_hours / automated_hours - 1.0) * 100.0,
                    "Conservative Detection (%)": DETECTION_CONSERVATIVE * 100.0,
                    "Automated Detection (%)": DETECTION_AUTOMATED * 100.0,
                    "Conservative Issues Found": issue_tests * DETECTION_CONSERVATIVE,
                    "Automated Issues Found": issue_tests * DETECTION_AUTOMATED,
                    "Conservative Issues Missed": issue_tests * (1.0 - DETECTION_CONSERVATIVE),
                    "Automated Issues Missed": issue_tests * (1.0 - DETECTION_AUTOMATED),
                }
            )
    return pd.DataFrame(rows)


def build_monte_carlo_summary_df(problematic_pcts: range, fe_insertions: int) -> pd.DataFrame:
    rng = np.random.default_rng(MONTE_CARLO_SEED)
    rows: list[dict[str, float]] = []
    for problematic_pct in problematic_pcts:
        issue_tests = IMPORTANT_TESTS * (problematic_pct / 100.0)
        conservative_reviewed_tests = rng.triangular(
            CONSERVATIVE_REVIEWED_TESTS_MIN,
            CONSERVATIVE_REVIEWED_TESTS_NOMINAL,
            CONSERVATIVE_REVIEWED_TESTS_MAX,
            size=MONTE_CARLO_SAMPLES,
        )
        conservative_problematic_seconds = rng.triangular(
            CONSERVATIVE_PROBLEMATIC_TEST_MIN_SECONDS,
            CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS,
            CONSERVATIVE_PROBLEMATIC_TEST_MAX_SECONDS,
            size=MONTE_CARLO_SAMPLES,
        )
        automated_problematic_seconds = rng.triangular(
            AUTOMATED_PROBLEMATIC_TEST_MIN_SECONDS,
            AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS,
            AUTOMATED_PROBLEMATIC_TEST_MAX_SECONDS,
            size=MONTE_CARLO_SAMPLES,
        )

        conservative_hours = compute_conservative_hours(
            issue_tests,
            conservative_reviewed_tests,
            conservative_problematic_seconds,
            fe_insertions,
        )
        automated_hours = compute_automated_hours(
            issue_tests,
            automated_problematic_seconds,
            fe_insertions,
        )
        productivity_improvement = (conservative_hours / automated_hours - 1.0) * 100.0

        rows.append(
            {
                "Problematic Tests (%)": problematic_pct,
                "Problematic Tests Count": issue_tests,
                "Conservative Review Time Mean (h)": float(np.mean(conservative_hours)),
                "Conservative Review Time P5 (h)": float(np.percentile(conservative_hours, 5)),
                "Conservative Review Time P10 (h)": float(np.percentile(conservative_hours, 10)),
                "Conservative Review Time P50 (h)": float(np.percentile(conservative_hours, 50)),
                "Conservative Review Time P90 (h)": float(np.percentile(conservative_hours, 90)),
                "Conservative Review Time P95 (h)": float(np.percentile(conservative_hours, 95)),
                "Automated Review Time Mean (h)": float(np.mean(automated_hours)),
                "Automated Review Time P5 (h)": float(np.percentile(automated_hours, 5)),
                "Automated Review Time P10 (h)": float(np.percentile(automated_hours, 10)),
                "Automated Review Time P50 (h)": float(np.percentile(automated_hours, 50)),
                "Automated Review Time P90 (h)": float(np.percentile(automated_hours, 90)),
                "Automated Review Time P95 (h)": float(np.percentile(automated_hours, 95)),
                "Productivity Improvement Mean (%)": float(np.mean(productivity_improvement)),
                "Productivity Improvement P5 (%)": float(np.percentile(productivity_improvement, 5)),
                "Productivity Improvement P10 (%)": float(np.percentile(productivity_improvement, 10)),
                "Productivity Improvement P50 (%)": float(np.percentile(productivity_improvement, 50)),
                "Productivity Improvement P90 (%)": float(np.percentile(productivity_improvement, 90)),
                "Productivity Improvement P95 (%)": float(np.percentile(productivity_improvement, 95)),
                "Conservative Detection (%)": DETECTION_CONSERVATIVE * 100.0,
                "Automated Detection (%)": DETECTION_AUTOMATED * 100.0,
                "Conservative Issues Found": issue_tests * DETECTION_CONSERVATIVE,
                "Automated Issues Found": issue_tests * DETECTION_AUTOMATED,
                "Conservative Issues Missed": issue_tests * (1.0 - DETECTION_CONSERVATIVE),
                "Automated Issues Missed": issue_tests * (1.0 - DETECTION_AUTOMATED),
            }
        )
    return pd.DataFrame(rows)


def autosize_worksheet(worksheet) -> None:
    for column in worksheet.columns:
        values = ["" if cell.value is None else str(cell.value) for cell in column]
        width = max(len(value) for value in values) + 2
        worksheet.column_dimensions[get_column_letter(column[0].column)].width = min(width, 40)


def build_timing_assumptions_table_rows(fe_insertions: int) -> list[list[str]]:
    return [
        ["FE insertions", str(fe_insertions), str(fe_insertions)],
        ["Local copy/load overhead", f"{CONSERVATIVE_LOCAL_COPY_MINUTES} min", f"{AUTOMATED_LOCAL_COPY_MINUTES} min"],
        [
            "Reviewed tests",
            f"{CONSERVATIVE_REVIEWED_TESTS_MIN} / {CONSERVATIVE_REVIEWED_TESTS_NOMINAL} / {CONSERVATIVE_REVIEWED_TESTS_MAX}",
            str(IMPORTANT_TESTS),
        ],
        ["Baseline scan time", f"{CONSERVATIVE_SECONDS_PER_REVIEWED_TEST} s/test", "-"],
        [
            "Problematic test handling",
            f"{CONSERVATIVE_PROBLEMATIC_TEST_MIN_SECONDS} / {CONSERVATIVE_PROBLEMATIC_TEST_NOMINAL_SECONDS} / {CONSERVATIVE_PROBLEMATIC_TEST_MAX_SECONDS} s",
            f"{AUTOMATED_PROBLEMATIC_TEST_MIN_SECONDS} / {AUTOMATED_PROBLEMATIC_TEST_NOMINAL_SECONDS} / {AUTOMATED_PROBLEMATIC_TEST_MAX_SECONDS} s",
        ],
        ["Wrap-up/reporting overhead", f"{CONSERVATIVE_WRAP_UP_MINUTES} min", f"{AUTOMATED_WRAP_UP_MINUTES} min"],
    ]


def compute_table_col_widths(rows: list[list[str]], headers: list[str], padding: float = 2.0) -> list[float]:
    max_lengths = [len(header) for header in headers]
    for row in rows:
        for index, value in enumerate(row):
            max_lengths[index] = max(max_lengths[index], len(str(value)))

    weighted_lengths = [length + padding for length in max_lengths]
    total = sum(weighted_lengths)
    return [length / total for length in weighted_lengths]


def build_scenario_plot(scenario_curve_df: pd.DataFrame, plot_path: Path) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5), constrained_layout=True)
    x_values = sorted(scenario_curve_df["Problematic Tests (%)"].unique())

    for metric, axis, ylabel in (
        ("Conservative Review Time (h)", axes[0], "Review time (h)"),
        ("Automated Review Time (h)", axes[0], "Review time (h)"),
    ):
        model_name = "Conservative reviewer" if "Conservative" in metric else "Automated review"
        pivot_df = scenario_curve_df.pivot(index="Problematic Tests (%)", columns="Scenario", values=metric)
        axis.fill_between(
            x_values,
            pivot_df["Best"].reindex(x_values),
            pivot_df["Worst"].reindex(x_values),
            color=SCENARIO_COLORS[model_name],
            alpha=0.15,
        )
        axis.plot(
            x_values,
            pivot_df["Nominal"].reindex(x_values),
            linewidth=2,
            label=f"{model_name} nominal",
            color=SCENARIO_COLORS[model_name],
        )

    productivity_pivot_df = scenario_curve_df.pivot(
        index="Problematic Tests (%)",
        columns="Scenario",
        values="Automation Productivity Improvement vs Conservative (%)",
    )
    axes[1].fill_between(
        x_values,
        productivity_pivot_df.min(axis=1).reindex(x_values),
        productivity_pivot_df.max(axis=1).reindex(x_values),
        color=SCENARIO_COLORS["Automated review"],
        alpha=0.15,
        label="Automated improvement band",
    )
    axes[1].plot(
        x_values,
        productivity_pivot_df["Nominal"].reindex(x_values),
        linewidth=2,
        color=SCENARIO_COLORS["Automated review"],
        label="Automated improvement nominal",
    )
    axes[1].axhline(0.0, color="#666666", linewidth=1, linestyle="--", alpha=0.7)

    axes[0].set_title("Scenario Review Time vs Problematic Tests")
    axes[0].set_xlabel("Problematic tests (%)")
    axes[0].set_ylabel("Review time (h)")
    axes[0].grid(alpha=0.3)
    axes[0].legend()

    axes[1].set_title("Scenario Productivity Improvement vs Conservative")
    axes[1].set_xlabel("Problematic tests (%)")
    axes[1].set_ylabel("Productivity improvement (%)")
    axes[1].grid(alpha=0.3)
    axes[1].legend()

    fig.suptitle("Scenario Bands: Automated vs Conservative Reviewer", fontsize=14, fontweight="bold")
    fig.savefig(plot_path, dpi=200)
    plt.close(fig)


def build_monte_carlo_plot(monte_carlo_df: pd.DataFrame, plot_path: Path, fe_insertions: int) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5), constrained_layout=True)
    x_values = monte_carlo_df["Problematic Tests (%)"]
    table_headers = ["Timing", "Conservative", "Automated"]
    table_rows = build_timing_assumptions_table_rows(fe_insertions)

    for model_name, color, prefix in (
        ("Conservative reviewer", SCENARIO_COLORS["Conservative reviewer"], "Conservative Review Time"),
        ("Automated review", SCENARIO_COLORS["Automated review"], "Automated Review Time"),
    ):
        axes[0].fill_between(
            x_values,
            monte_carlo_df[f"{prefix} P5 (h)"],
            monte_carlo_df[f"{prefix} P95 (h)"],
            color=color,
            alpha=0.12,
        )
        axes[0].fill_between(
            x_values,
            monte_carlo_df[f"{prefix} P10 (h)"],
            monte_carlo_df[f"{prefix} P90 (h)"],
            color=color,
            alpha=0.22,
        )
        axes[0].plot(
            x_values,
            monte_carlo_df[f"{prefix} P50 (h)"],
            linewidth=2,
            label=f"{model_name} median",
            color=color,
        )

    axes[1].fill_between(
        x_values,
        monte_carlo_df["Productivity Improvement P5 (%)"],
        monte_carlo_df["Productivity Improvement P95 (%)"],
        color=SCENARIO_COLORS["Automated review"],
        alpha=0.12,
    )
    axes[1].fill_between(
        x_values,
        monte_carlo_df["Productivity Improvement P10 (%)"],
        monte_carlo_df["Productivity Improvement P90 (%)"],
        color=SCENARIO_COLORS["Automated review"],
        alpha=0.22,
    )
    axes[1].plot(
        x_values,
        monte_carlo_df["Productivity Improvement P50 (%)"],
        linewidth=2,
        label="Automated improvement median",
        color=SCENARIO_COLORS["Automated review"],
    )
    axes[1].axhline(0.0, color="#666666", linewidth=1, linestyle="--", alpha=0.7)

    axes[0].set_title("Monte Carlo Review Time Bands")
    axes[0].set_xlabel("Problematic tests (%)")
    axes[0].set_ylabel("Review time (h)")
    axes[0].grid(alpha=0.3)
    axes[0].legend()

    axes[1].set_title("Monte Carlo Productivity Improvement Bands")
    axes[1].set_xlabel("Problematic tests (%)")
    axes[1].set_ylabel("Productivity improvement (%)")
    axes[1].grid(alpha=0.3)
    axes[1].legend()

    assumptions_table = axes[0].table(
        cellText=table_rows,
        colLabels=table_headers,
        cellLoc="left",
        colLoc="left",
        bbox=[0.01, 0.58, 0.53, 0.27],
        colWidths=compute_table_col_widths(table_rows, table_headers),
    )
    assumptions_table.auto_set_font_size(False)
    assumptions_table.set_fontsize(5.2)

    for (row, col), cell in assumptions_table.get_celld().items():
        cell.set_linewidth(0.4)
        if row == 0:
            cell.set_text_props(weight="bold", color="white")
            cell.set_facecolor("#1f1f1f")
        elif col == 0:
            cell.set_text_props(weight="bold")
            cell.set_facecolor((0.95, 0.95, 0.95, 0.92))
        else:
            cell.set_facecolor((1.0, 0.98, 0.96, 0.92) if col == 1 else (0.96, 0.98, 1.0, 0.92))

    fig.suptitle("Monte Carlo Uncertainty: Automated vs Conservative Reviewer", fontsize=14, fontweight="bold")
    fig.savefig(plot_path, dpi=200)
    plt.close(fig)


def format_workbook(workbook_path: Path) -> None:
    workbook = load_workbook(workbook_path)
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        autosize_worksheet(worksheet)
    workbook.save(workbook_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate productivity comparison artifacts.")
    parser.add_argument(
        "--fe-insertions",
        type=int,
        default=DEFAULT_FE_INSERTIONS,
        help="Number of FE insertions reviewed per analysis run.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    fe_insertions = max(1, args.fe_insertions)
    output_dir = Path(__file__).resolve().parent

    assumptions_df = build_assumptions_df(fe_insertions)
    nominal_summary_df = build_nominal_summary_df((5.0, 2.0, 1.0), fe_insertions)
    scenario_curve_df = build_scenario_curve_df(range(0, 11), fe_insertions)
    monte_carlo_df = build_monte_carlo_summary_df(range(0, 11), fe_insertions)

    workbook_path = output_dir / "review_productivity_summary_1000_tests.xlsx"
    scenario_plot_path = output_dir / "review_productivity_comparison_1000_tests.png"
    monte_carlo_plot_path = output_dir / "review_productivity_monte_carlo_1000_tests.png"

    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        assumptions_df.to_excel(writer, sheet_name="Assumptions", index=False)
        nominal_summary_df.to_excel(writer, sheet_name="Nominal_Summary", index=False)
        scenario_curve_df.to_excel(writer, sheet_name="Scenario_Curve_0_to_10pct", index=False)
        monte_carlo_df.to_excel(writer, sheet_name="MonteCarlo_0_to_10pct", index=False)

    format_workbook(workbook_path)
    build_scenario_plot(scenario_curve_df, scenario_plot_path)
    build_monte_carlo_plot(monte_carlo_df, monte_carlo_plot_path, fe_insertions)

    print(f"Workbook: {workbook_path}")
    print(f"Scenario plot: {scenario_plot_path}")
    print(f"Monte Carlo plot: {monte_carlo_plot_path}")
    print(f"FE insertions: {fe_insertions}")


if __name__ == "__main__":
    main()
