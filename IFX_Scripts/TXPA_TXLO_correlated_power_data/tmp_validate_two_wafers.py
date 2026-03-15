from __future__ import annotations

import importlib.util
import shutil
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(r"C:/UserData/Learning/Software_Programming/GitHub_Nelio92")
SOURCE_DIR = Path(r"C:/UserData/Infineon/TE_CTRX/CTRX_8188_8144/Data_Reviews/FE_Test_D/Test_D33")
FILE_NAMES = [
    "3FTCU151R01_014_S11P_20260217083728_M5358ACSH1D3311_RBGEUFRF32.std.tar.gz",
    "3FTCU151R01_014_S21P_20260222052107_M5358ACSC2D3311_RBGEUFRF32.std.tar.gz",
    "3FTCU151R01_014_S31P_20260226105450_M5358ACSA3D3411_RBGEUFRF32.std.tar.gz",
    "3FTCU151R01_015_S11P_20260218083933_M5358ACSH1D3311_RBGEUFRF32.std.tar.gz",
    "3FTCU151R01_015_S21P_20260222132446_M5358ACSC2D3311_RBGEUFRF32.std.tar.gz",
    "3FTCU151R01_015_S31P_20260226142944_M5358ACSA3D3411_RBGEUFRF32.std.tar.gz",
]


def load_module():
    script_path = REPO_ROOT / "Tasks_Automation_Code/IFX_Scripts/TXPA_TXLO_correlated_power_data/generate_txpa_txlo_correlated_power_report.py"
    spec = importlib.util.spec_from_file_location("txpa_report", script_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def main() -> int:
    module = load_module()
    with tempfile.TemporaryDirectory(prefix="txpa_txlo_two_wafers_") as temp_dir:
        test_dir = Path(temp_dir)
        for name in FILE_NAMES:
            shutil.copy2(SOURCE_DIR / name, test_dir / name)

        print(f"TEST_INPUT={test_dir}", flush=True)
        module.INPUT_FOLDER = test_dir
        module.ENABLE_PRE_CORRELATION = False
        module.MAX_FILES = None
        rc = int(module.main())

        outputs = sorted((test_dir / "Outputs").glob("txpa_txlo_power_cdf_*"))
        latest_output = outputs[-1] if outputs else None
        print(f"LATEST_OUTPUT={latest_output}", flush=True)
        if latest_output is not None:
            summary = latest_output / "aggregated_series_summary.csv"
            processed = latest_output / "processed_input_files.csv"
            print(f"SUMMARY_EXISTS={summary.exists()}", flush=True)
            print(f"PROCESSED_EXISTS={processed.exists()}", flush=True)
        return rc


if __name__ == "__main__":
    raise SystemExit(main())
