from __future__ import annotations

import csv
import json
import io
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock


TEST_DATA_ANALYSIS_DIR = Path(__file__).resolve().parents[1]
if str(TEST_DATA_ANALYSIS_DIR) not in sys.path:
    sys.path.insert(0, str(TEST_DATA_ANALYSIS_DIR))

import Tests_Data_Analysis as analysis
import run_tests_data_analysis as launcher
import stdf_to_flat_csv as converter


class FakeRecord:
    def __init__(self, record_id: str, **fields: object) -> None:
        self.id = record_id
        self._fields = fields

    def to_dict(self) -> dict[str, object]:
        return dict(self._fields)


class StdfToFlatCsvUnitTests(unittest.TestCase):
    def test_convert_records_to_csv_generates_analysis_compatible_layout(self) -> None:
        records = [
            FakeRecord("MIR", LOT_ID="LOT_A", SBLOT_ID="SUB1"),
            FakeRecord("WIR", WAFER_ID="W1"),
            FakeRecord("DTR", TEXT_DAT="ALARM DVI:0014 : Overrange alarm"),
            FakeRecord("PIR", SITE_NUM=0),
            FakeRecord(
                "PTR",
                TEST_NUM=520123,
                TEST_TXT="TXPA_OUTPUT_PWR",
                RESULT=9.2,
                LO_LIMIT=9.5,
                HI_LIMIT=10.5,
                UNITS="dBm",
            ),
            FakeRecord(
                "PTR",
                TEST_NUM=530045,
                TEST_TXT="DPLL_LOCK_TIME",
                RESULT=1.0,
                LO_LIMIT=0.0,
                HI_LIMIT=5.0,
                UNITS="us",
            ),
            FakeRecord("PRR", SITE_NUM=0, X_COORD=1, Y_COORD=1, PART_ID="CHIP_A"),
            FakeRecord("PIR", SITE_NUM=1),
            FakeRecord(
                "PTR",
                TEST_NUM=520123,
                TEST_TXT="TXPA_OUTPUT_PWR",
                RESULT=10.0,
                LO_LIMIT=9.5,
                HI_LIMIT=10.5,
                UNITS="dBm",
            ),
            FakeRecord(
                "PTR",
                TEST_NUM=530045,
                TEST_TXT="DPLL_LOCK_TIME",
                RESULT=2.0,
                LO_LIMIT=0.0,
                HI_LIMIT=5.0,
                UNITS="us",
            ),
            FakeRecord("PRR", SITE_NUM=1, X_COORD=1, Y_COORD=2, PART_ID="CHIP_B"),
        ]

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_csv = Path(tmp_dir) / "converted.csv"
            artifacts_dir = Path(tmp_dir) / "Artifacts"
            summary = converter.convert_records_to_csv(records, output_csv, artifacts_output_folder=artifacts_dir)

            self.assertEqual(summary.converted_files, 1)
            self.assertEqual(summary.converted_parts, 2)
            self.assertEqual(summary.converted_tests, 2)
            self.assertTrue(output_csv.exists())
            self.assertEqual(len(summary.dtr_files), 1)
            self.assertEqual(len(summary.consistency_files), 1)
            self.assertTrue(summary.dtr_files[0].exists())
            self.assertTrue(summary.consistency_files[0].exists())
            self.assertEqual(summary.dtr_files[0].parent, artifacts_dir)
            self.assertEqual(summary.consistency_files[0].parent, artifacts_dir)

            meta = analysis.scan_flat_file_meta(output_csv, encoding="utf-8")
            self.assertEqual(meta.header[:10], converter.META_COLUMNS)
            self.assertEqual(meta.numeric_test_cols, ["520123", "530045"])
            self.assertEqual(meta.meta_rows["Test Name"]["520123"], "TXPA_OUTPUT_PWR")
            self.assertEqual(meta.meta_rows["Yield"]["520123"], "50")
            self.assertEqual(meta.meta_rows["Yield"]["530045"], "100")

            df = analysis._read_unit_data(
                output_csv,
                data_start_line_index=meta.data_start_line_index,
                usecols=meta.header,
                encoding="utf-8",
            )
            self.assertEqual(df.shape[0], 2)
            self.assertEqual(df.loc[0, "PF"], "F")
            self.assertEqual(df.loc[0, "FIRST_FAIL_TEST"], "TXPA_OUTPUT_PWR")
            self.assertEqual(df.loc[1, "PF"], "P")
            self.assertEqual(float(df.loc[1, "520123"]), 10.0)

            with summary.dtr_files[0].open("r", encoding="utf-8", newline="") as handle:
                dtr_rows = list(csv.reader(handle, delimiter=converter.DELIMITER))
            self.assertEqual(dtr_rows[0], ["Index", "Category", "Message"])
            self.assertEqual(dtr_rows[1][1], "ALARM")
            self.assertIn("Overrange alarm", dtr_rows[1][2])

            with summary.consistency_files[0].open("r", encoding="utf-8") as handle:
                consistency = json.load(handle)
            self.assertEqual(consistency["generated_rows"], 2)
            self.assertEqual(consistency["numeric_tests"], 2)
            self.assertEqual(consistency["dtr_record_count"], 1)
            self.assertTrue(consistency["checks"]["pir_equals_prr"])
            self.assertTrue(consistency["checks"]["all_rows_have_measurements"])
            self.assertEqual(summary.file_results[0].dtr_record_count, 1)

    def test_convert_stdf_before_analysis_allows_existing_input_folder_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            generated_csv_dir = tmp_path / "generated_csv"
            generated_csv_dir.mkdir()
            config = launcher._validate_and_normalize_config(
                {
                    "input_folder": str(generated_csv_dir),
                    "output_folder": str(TEST_DATA_ANALYSIS_DIR / "tests"),
                    "modules": ["TXPA"],
                    "convert_stdf_before_analysis": True,
                },
                TEST_DATA_ANALYSIS_DIR / "configs" / "config_default.yaml",
            )

            self.assertEqual(config["input_folder"], generated_csv_dir.resolve())

    def test_convert_stdf_before_analysis_allows_missing_generated_csv_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            reports_dir = tmp_path / "reports"
            reports_dir.mkdir()
            config = launcher._validate_and_normalize_config(
                {
                    "input_folder": str(tmp_path / "generated_csv"),
                    "output_folder": str(reports_dir),
                    "modules": ["TXPA"],
                    "convert_stdf_before_analysis": True,
                },
                tmp_path / "config.yaml",
            )

            self.assertEqual(config["input_folder"], (tmp_path / "generated_csv").resolve())

    def test_convert_stdf_before_analysis_allows_existing_csvs_in_input_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            generated_csv_dir = tmp_path / "generated_csv"
            generated_csv_dir.mkdir()
            (generated_csv_dir / "already_generated.csv").write_text("UNIT_ID;SITE_NUM\n", encoding="utf-8")

            config = launcher._validate_and_normalize_config(
                {
                    "input_folder": str(generated_csv_dir),
                    "output_folder": str(tmp_path / "reports"),
                    "modules": ["TXPA"],
                    "convert_stdf_before_analysis": True,
                },
                tmp_path / "config.yaml",
            )

            self.assertEqual(config["input_folder"], generated_csv_dir.resolve())

    def test_prepare_stdf_inputs_skips_conversion_when_csv_already_exists(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_dir = tmp_path / "generated_csv"
            input_dir.mkdir()
            existing_csv = input_dir / "already_generated.csv"
            existing_csv.write_text("UNIT_ID;SITE_NUM\n", encoding="utf-8")

            previous_single_file = analysis.SINGLE_FILE
            try:
                analysis.SINGLE_FILE = None
                stdout = io.StringIO()
                with redirect_stdout(stdout), mock.patch.object(
                    launcher.stdf_to_flat_csv,
                    "convert_stdf_folder",
                ) as convert_mock:
                    launcher._prepare_stdf_inputs(
                        {
                            "convert_stdf_before_analysis": True,
                            "input_folder": input_dir,
                        }
                    )

                convert_mock.assert_not_called()
                self.assertIn("STDF pre-conversion skipped", stdout.getvalue())
                self.assertIn("existing CSV input file(s)", stdout.getvalue())
            finally:
                analysis.SINGLE_FILE = previous_single_file

    def test_prepare_stdf_inputs_skips_conversion_when_target_csv_exists(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_dir = tmp_path / "generated_csv"
            input_dir.mkdir()
            existing_csv = input_dir / "device_lot.csv"
            existing_csv.write_text("UNIT_ID;SITE_NUM\n", encoding="utf-8")

            previous_single_file = analysis.SINGLE_FILE
            try:
                analysis.SINGLE_FILE = "device_lot.std"
                stdout = io.StringIO()
                with redirect_stdout(stdout), mock.patch.object(
                    launcher.stdf_to_flat_csv,
                    "convert_stdf_folder",
                ) as convert_mock:
                    launcher._prepare_stdf_inputs(
                        {
                            "convert_stdf_before_analysis": True,
                            "input_folder": input_dir,
                        }
                    )

                convert_mock.assert_not_called()
                self.assertEqual(analysis.SINGLE_FILE, "device_lot.csv")
                self.assertIn("STDF pre-conversion skipped", stdout.getvalue())
                self.assertIn("device_lot.csv", stdout.getvalue())
            finally:
                analysis.SINGLE_FILE = previous_single_file


if __name__ == "__main__":
    unittest.main()
