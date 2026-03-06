from __future__ import annotations

import sys
import tempfile
import textwrap
import unittest
from pathlib import Path


TEST_DATA_ANALYSIS_DIR = Path(__file__).resolve().parents[1]
if str(TEST_DATA_ANALYSIS_DIR) not in sys.path:
    sys.path.insert(0, str(TEST_DATA_ANALYSIS_DIR))

import Tests_Data_Analysis as analysis


SAMPLE_INPUT = TEST_DATA_ANALYSIS_DIR / "tests" / "smoke_input" / "smoke_Q2_sample.csv"


class MetaParsingUnitTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.meta = analysis.scan_flat_file_meta(SAMPLE_INPUT, encoding="latin1")

    def test_scan_flat_file_meta_extracts_expected_structure(self) -> None:
        self.assertEqual(self.meta.header[:6], ["UNIT_ID", "SITE_NUM", "WAFER", "X", "Y", "LOT"])
        self.assertEqual(self.meta.numeric_test_cols, ["520123", "530045"])
        self.assertEqual(self.meta.data_start_line_index, 9)
        self.assertIn("Test Name", self.meta.meta_rows)
        self.assertIn("Yield", self.meta.meta_rows)
        self.assertEqual(self.meta.meta_rows["Test Name"]["520123"], "TXPA_OUTPUT_PWR")
        self.assertEqual(self.meta.meta_rows["Cpk"]["530045"], "2.10")

    def test_meta_accessors_return_expected_values(self) -> None:
        self.assertEqual(analysis._test_name_from_meta(self.meta, "520123"), "TXPA_OUTPUT_PWR")
        self.assertEqual(analysis._test_name_from_meta(self.meta, "530045"), "DPLL_LOCK_TIME")
        self.assertEqual(analysis._module_from_test_name("txpa_output_pwr"), "TXPA")
        self.assertEqual(analysis._module_from_test_name("ab"), "AB")

        low, high, unit = analysis._limits_from_meta(self.meta, "520123")
        self.assertEqual(low, 9.5)
        self.assertEqual(high, 10.5)
        self.assertEqual(unit, "dBm")

        yield_pct, cpk = analysis._yield_cpk_from_meta(self.meta, "520123")
        self.assertEqual(yield_pct, 95.0)
        self.assertEqual(cpk, 1.2)

        yield_pct_ok, cpk_ok = analysis._yield_cpk_from_meta(self.meta, "530045")
        self.assertEqual(yield_pct_ok, 100.0)
        self.assertEqual(cpk_ok, 2.1)

    def test_scan_flat_file_meta_rejects_invalid_inputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            empty_file = Path(tmp_dir) / "empty.csv"
            empty_file.write_text("", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Empty file"):
                analysis.scan_flat_file_meta(empty_file)

            no_numeric = Path(tmp_dir) / "no_numeric.csv"
            no_numeric.write_text(
                textwrap.dedent(
                    """\
                    LOT;SITE_NUM;WAFER
                    Test Name;;; 
                    1;0;W1
                    """
                ),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ValueError, "Could not find numeric test columns"):
                analysis.scan_flat_file_meta(no_numeric)


class StatusLogicUnitTests(unittest.TestCase):
    def test_status_for_test_prioritizes_yield_fail(self) -> None:
        status = analysis._status_for_test(
            yield_pct=99.0,
            cpk=0.5,
            yield_threshold=100.0,
            cpk_low=1.67,
            cpk_high=20.0,
        )
        self.assertEqual(status, "FAILS")

    def test_status_for_test_handles_cpk_limits_and_boundaries(self) -> None:
        self.assertEqual(
            analysis._status_for_test(
                yield_pct=100.0,
                cpk=1.2,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
            ),
            "Cpk<1.67",
        )
        self.assertEqual(
            analysis._status_for_test(
                yield_pct=100.0,
                cpk=25.0,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
            ),
            "Cpk>20",
        )
        self.assertIsNone(
            analysis._status_for_test(
                yield_pct=100.0,
                cpk=1.67,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
            )
        )
        self.assertIsNone(
            analysis._status_for_test(
                yield_pct=None,
                cpk=20.0,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
            )
        )
        self.assertIsNone(
            analysis._status_for_test(
                yield_pct=None,
                cpk=None,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
            )
        )


if __name__ == "__main__":
    unittest.main()
