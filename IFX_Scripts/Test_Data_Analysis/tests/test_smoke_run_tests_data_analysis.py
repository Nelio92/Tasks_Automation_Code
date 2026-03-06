from __future__ import annotations

import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

from openpyxl import load_workbook


REPO_ROOT = Path(__file__).resolve().parents[3]
IFX_SCRIPTS_DIR = REPO_ROOT / "Tasks_Automation_Code" / "IFX_Scripts"
SAMPLE_INPUT_DIR = IFX_SCRIPTS_DIR / "tests" / "smoke_input"
SAMPLE_FILE_NAME = "smoke_Q2_sample.csv"
SAMPLE_SHEET_NAME = "smoke_Q2_sample"
SAMPLE_PLOTS_SHEET_NAME = "smoke_Q2_sample_PLOTS"


class TestsDataAnalysisSmokeTest(unittest.TestCase):
    def test_cli_generates_expected_reports(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            output_dir = tmp_path / "outputs"
            output_dir.mkdir(parents=True, exist_ok=True)
            config_path = tmp_path / "config_smoke.yaml"
            config_path.write_text(
                textwrap.dedent(
                    f"""\
                    input_folder: {SAMPLE_INPUT_DIR.as_posix()}
                    output_folder: {output_dir.as_posix()}
                    modules:
                      - TXPA
                      - DPLL
                    yield_threshold: 100.0
                    cpk_low: 1.67
                    cpk_high: 20.0
                    outlier_mad_multiplier: 6.0
                    max_files: 1
                    single_file: {SAMPLE_FILE_NAME}
                    encoding: latin1
                    generate_correlation_report: true
                    correlation_methods:
                      - pearson
                    pearson_abs_min_for_report: 0.5
                    wafermap_circle_area_mult: 1.0
                    """
                ),
                encoding="utf-8",
            )

            result = subprocess.run(
                [sys.executable, "run_tests_data_analysis.py", "--config", str(config_path)],
                cwd=IFX_SCRIPTS_DIR,
                capture_output=True,
                text=True,
                check=False,
            )

            self.assertEqual(
                result.returncode,
                0,
                msg=(
                    "Smoke test launcher failed.\n"
                    f"STDOUT:\n{result.stdout}\n\n"
                    f"STDERR:\n{result.stderr}"
                ),
            )

            yield_report = output_dir / "Yield_Cpk_Report.xlsx"
            correlation_report = output_dir / "Correlation_Report.xlsx"
            self.assertTrue(yield_report.exists(), "Yield/Cpk report was not created")
            self.assertTrue(correlation_report.exists(), "Correlation report was not created")

            png_files = list((output_dir / "cdf_plots").rglob("*.png"))
            self.assertGreaterEqual(len(png_files), 2, "Expected CDF and wafer map PNG outputs")

            workbook = load_workbook(yield_report, read_only=True, data_only=True)
            try:
                self.assertIn(SAMPLE_SHEET_NAME, workbook.sheetnames)
                self.assertIn(SAMPLE_PLOTS_SHEET_NAME, workbook.sheetnames)
                worksheet = workbook[SAMPLE_SHEET_NAME]
                data_rows = list(worksheet.iter_rows(min_row=2, values_only=True))
                self.assertEqual(len(data_rows), 1, "Expected exactly one affected test row in the smoke sample")

                row = data_rows[0]
                self.assertEqual(row[0], SAMPLE_FILE_NAME)
                self.assertEqual(row[1], "TXPA")
                self.assertEqual(row[2], 520123)
                self.assertEqual(row[3], "TXPA_OUTPUT_PWR")
                self.assertEqual(row[7], "FAILS")
                self.assertEqual(row[9], 2)
                self.assertEqual(row[18], "View")
            finally:
                workbook.close()

            corr_workbook = load_workbook(correlation_report, read_only=True, data_only=True)
            try:
                self.assertIn(SAMPLE_SHEET_NAME, corr_workbook.sheetnames)
                corr_sheet = corr_workbook[SAMPLE_SHEET_NAME]
                corr_rows = list(corr_sheet.iter_rows(min_row=2, values_only=True))
                self.assertGreaterEqual(len(corr_rows), 2, "Expected correlation rows in smoke report")
            finally:
                corr_workbook.close()


if __name__ == "__main__":
    unittest.main()
