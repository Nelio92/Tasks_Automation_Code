from __future__ import annotations

import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

import pandas as pd


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
        self.assertEqual(
            analysis._shorten_test_name("DPLL_ElapsTime____S980 <> DPLL_ElapsTime____S980  -1"),
            "DPLL_ElapsTime____S980",
        )

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
        self.assertEqual(status, "Fails + Cpk<1.67")

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
        self.assertEqual(
            analysis._status_for_test(
                yield_pct=100.0,
                cpk=20.0,
                yield_threshold=100.0,
                cpk_low=1.67,
                cpk_high=20.0,
                site_to_site_delta=True,
            ),
            "Site-to-Site Delta",
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

    def test_assess_test_metrics_detects_site_delta_and_unique_values(self) -> None:
        sample_count = 90
        metric_frame = pd.DataFrame(
            {
                "SITE_NUM": ([0] * 30) + ([1] * 30) + ([2] * 30),
                "WAFER": (["W1"] * 45) + (["W2"] * 45),
                "X": list(range(sample_count)),
                "Y": [idx % 6 for idx in range(sample_count)],
            }
        )
        series = pd.Series(([0.0, 1.0] * 15) + ([9.0, 10.0] * 15) + ([9.0, 10.0] * 15))

        assessment = analysis._assess_test_metrics(
            series=series,
            meta_cols=metric_frame,
            unit="V",
            yield_pct=100.0,
            cpk=25.0,
            yield_threshold=100.0,
            cpk_low=1.67,
            cpk_high=20.0,
            wafer_sig="S21P",
        )

        self.assertIn(analysis.METRIC_CPK_HIGH, assessment.metric_keys)
        self.assertIn(analysis.METRIC_SITE_DELTA, assessment.metric_keys)
        self.assertIn(analysis.METRIC_UNIQUE_VALUES, assessment.metric_keys)
        self.assertEqual(assessment.priority, "MEDIUM")

    def test_assess_test_metrics_detects_multimodality_reason(self) -> None:
        sample_count = 100
        metric_frame = pd.DataFrame(
            {
                "SITE_NUM": ([0] * 50) + ([1] * 50),
                "WAFER": (["W1"] * 50) + (["W2"] * 50),
                "X": list(range(sample_count)),
                "Y": [idx % 10 for idx in range(sample_count)],
            }
        )
        series = pd.Series(([0.0] * 50) + ([8.0] * 50))

        assessment = analysis._assess_test_metrics(
            series=series,
            meta_cols=metric_frame,
            unit="V",
            yield_pct=100.0,
            cpk=2.0,
            yield_threshold=100.0,
            cpk_low=1.67,
            cpk_high=20.0,
            wafer_sig="S21P",
        )

        self.assertIn(analysis.METRIC_UNIQUE_VALUES, assessment.metric_keys)
        self.assertIn(analysis.METRIC_MULTIMODALITY, assessment.metric_keys)
        self.assertEqual(assessment.priority, "MEDIUM")
        self.assertGreaterEqual(assessment.peak_count, 2)
        self.assertIsNotNone(assessment.multimodality_reason)

    def test_unique_values_ignores_missing_and_hash_units(self) -> None:
        digital_like = pd.Series(([0.0, 1.0] * 20), dtype=float)

        unique_none, is_analog_none = analysis._unique_value_count(digital_like, unit=None)
        self.assertIsNone(unique_none)
        self.assertFalse(is_analog_none)

        unique_hash, is_analog_hash = analysis._unique_value_count(digital_like, unit="#")
        self.assertIsNone(unique_hash)
        self.assertFalse(is_analog_hash)

    def test_cdf_plot_by_site_creates_png_for_site_delta_data(self) -> None:
        series = pd.Series([0.1, 0.2, 0.3, 1.1, 1.2, 1.3], dtype=float)
        meta_cols = pd.DataFrame({"SITE_NUM": [0, 0, 0, 1, 1, 1]})

        with tempfile.TemporaryDirectory() as tmp_dir:
            out_path = Path(tmp_dir) / "site_cdf.png"
            analysis._cdf_plot_by_site_png(
                series,
                meta_cols=meta_cols,
                title="Example Test",
                out_path=out_path,
                low_limit=0.0,
                high_limit=2.0,
            )
            self.assertTrue(out_path.exists())


class CorrelationHelperUnitTests(unittest.TestCase):
    def test_safe_spearman_correlation_does_not_require_scipy(self) -> None:
        a = pd.Series([10, 20, 30, 40, 50], dtype=float)
        b = pd.Series([1, 2, 3, 4, 5], dtype=float)
        c = pd.Series([5, 4, 3, 2, 1], dtype=float)

        self.assertAlmostEqual(float(analysis._safe_spearman_correlation(a, b)), 1.0, places=12)
        self.assertAlmostEqual(float(analysis._safe_spearman_correlation(a, c)), -1.0, places=12)
        self.assertIsNone(analysis._safe_spearman_correlation(pd.Series([1.0]), pd.Series([2.0])))

class WaferNormalizationUnitTests(unittest.TestCase):
    def test_normalize_wafer_ids_extracts_numeric_values(self) -> None:
        normalized = analysis._normalize_wafer_ids(
            pd.Series(["WafNr=24", "24", "Wafer 007", "nan", "", "ABC", None])
        )

        self.assertEqual(normalized.iloc[0], "24")
        self.assertEqual(normalized.iloc[1], "24")
        self.assertEqual(normalized.iloc[2], "7")
        self.assertTrue(pd.isna(normalized.iloc[3]))
        self.assertTrue(pd.isna(normalized.iloc[4]))
        self.assertTrue(pd.isna(normalized.iloc[5]))
        self.assertTrue(pd.isna(normalized.iloc[6]))

        def test_supports_wafer_maps_excludes_packaged_and_q_files(self) -> None:
            self.assertFalse(analysis._supports_wafer_maps("device_B11_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_HT_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_B21_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_RT_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_Q11_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_Q21_sample.csv"))
            self.assertFalse(analysis._supports_wafer_maps("device_Q31_sample.csv"))
            self.assertTrue(analysis._supports_wafer_maps("device_S21P_sample.csv"))

class SheetNameUnitTests(unittest.TestCase):
    def test_unique_sheet_name_handles_truncation_collisions(self) -> None:
        base = "3FT6Y120R04_024_S11P_20260210192844_M5358ACSH1D3311_RBGEUFRF22"
        other = base + "_2"

        first = analysis._unique_sheet_name(base, [])
        second = analysis._unique_sheet_name(other, [first])

        self.assertLessEqual(len(first), 31)
        self.assertLessEqual(len(second), 31)
        self.assertNotEqual(first.lower(), second.lower())

if __name__ == "__main__":
    unittest.main()
