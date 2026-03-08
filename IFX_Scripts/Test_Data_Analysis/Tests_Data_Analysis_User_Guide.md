# Tests_Data_Analysis.py — Team Deployment & User Guide

## 1) Purpose

`Tests_Data_Analysis.py` automates fast review of large flat production CSV files and generates:

- A **Yield/Cpk Excel report** for selected modules.
- Embedded **CDF plots** and **wafer maps** for affected tests.
- Optional **Correlation report** (Pearson/Spearman) for module tests.

This enables Test Engineering to quickly identify failing/unstable tests without manual spreadsheet analysis.

---

## 2) What the script expects

### Input data format

The script expects semicolon-separated flat CSV files with this structure:

1. **Header row** containing metadata columns and many numeric test columns (for example `520123`, `530045`, ...).
2. A metadata block containing rows such as:
   - `Test Name`
   - `Low`
   - `High`
   - `Unit`
   - `Cpk`
   - `Yield`
3. Unit/device rows after metadata.

### Important assumptions

- Delimiter is `;`.
- Numeric test columns are detected by digit-only column names.
- Module is inferred from first 4 characters of `Test Name` (e.g. `DPLL`, `TXPA`, `TXLO`).

---

## 3) Outputs generated

By default under `PROD_Data/Outputs`:

1. `Test_Data_Analysis_Report.xlsx`
   - One data sheet per input CSV containing only affected tests.
   - A paired `*_PLOTS` sheet with embedded CDF + wafer map images.
   - Hyperlinks from data sheet (`CDF Plot` column) to plot locations.

2. `cdf_plots/<csv_stem>/...png`
   - CDF PNG per affected test.
   - Wafer-map PNG per affected test.

3. Optional: `Correlation_Report.xlsx`
   - Generated when `GENERATE_CORRELATION_REPORT = True`.

If a workbook is already open, the script writes a timestamped fallback file.

---

## 4) Current configuration model (in-script)

The script is configured through the **USER PARAMETERS** section at the top of the file.

Main parameters:

- `INPUT_FOLDER`
- `OUTPUT_FOLDER`
- `MODULES`
- `YIELD_THRESHOLD`
- `CPK_LOW`, `CPK_HIGH`
- `OUTLIER_MAD_MULTIPLIER`
- `MAX_FILES`
- `SINGLE_FILE`
- `ENCODING`
- `GENERATE_CORRELATION_REPORT`
- `CORRELATION_METHODS`
- `PEARSON_ABS_MIN_FOR_REPORT`
- `WAFERMAP_CIRCLE_AREA_MULT`
- `CONVERT_STDF_BEFORE_ANALYSIS`
- `STDF_SINGLE_FILE`
- `STDF_FILE_PATTERNS`

When STDF pre-conversion is enabled, the launcher uses `INPUT_FOLDER` as the single working folder for the whole flow: it reads source `.stdf` / `.std` files from that folder with the `pystdf` backend, generates the flat CSV files into that same folder, and then continues with the normal analysis flow from those generated CSV files. For each converted STDF file, the converter also writes a DTR sidecar CSV containing alarm/error text records and a JSON consistency report summarizing record counts, malformed-record skips, and basic `PIR`/`PRR`/row-count checks. In the launcher-driven flow, those sidecar artifacts are written into [Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis](Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis) output under `OUTPUT_FOLDER/Artifacts`.

---

## 5) Local setup and run

## Prerequisites

- Python 3.10+ recommended.
- Windows environment validated.

Install dependencies:

```powershell
pip install -r Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/requirements-tests-data-analysis.txt
```

Run only through the YAML launcher. Direct execution of `Tests_Data_Analysis.py` is intentionally disabled.

Note: all runtime settings are now owned by the YAML config files. `run_tests_data_analysis.py` loads the selected YAML profile and applies those values before `Tests_Data_Analysis.py` runs.

Optional STDF automation is available through the same launcher. Set `convert_stdf_before_analysis: true` to run the conversion step automatically before the report generation starts. The launcher now uses only `input_folder` for both the source STDF files and the generated CSV files. If needed, `stdf_single_file` limits conversion to one source STDF file and `stdf_file_patterns` narrows the discovery glob patterns. The standalone converter module remains reusable on its own, while the launcher simply calls it before starting the normal CSV analysis phase.

### Team runner (recommended)

Use the deployment launcher with YAML configuration:

```powershell
./Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/run_tests_data_analysis.ps1
```

Use a specific profile:

```powershell
./Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/run_tests_data_analysis.ps1 -ConfigPath Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/configs/config_txpa_focus.yaml
```

Dry run of resolved parameters:

```powershell
.venv/Scripts/python.exe Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/run_tests_data_analysis.py --config Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/configs/config_default.yaml --dry-run
```

---

## 6) How affected tests are selected

A test is included in the report if:

- `Yield < YIELD_THRESHOLD`, or
- `Cpk < CPK_LOW`, or
- `Cpk > CPK_HIGH`

Only tests whose module prefix is in `MODULES` are evaluated.

---

## 7) Interpretation notes

- **Status**:
  - `FAILS` => yield below threshold.
  - `Cpk<...` => too low capability.
  - `Cpk>...` => unusually high Cpk (possible data/spec issue).
- **Fail Chips** is computed against `Low`/`High` limits when available.
- **Outliers** use robust MAD criterion.
- **LTL/UTL 6s/12s** are proposed robust sigma-based limits.
- **Comment** includes heuristics such as spread, multimodality, site/wafer/coordinate signatures.

---

## 8) Professional deployment approach (recommended for Test Engineering)

For team-wide, maintainable use, deploy as an **internal versioned tool**.

### Phase A — Standardize execution

1. Keep one canonical source location in Git.
2. Add a pinned dependency file (for example `requirements-tests-data-analysis.txt`).
3. Provide one team launcher script (PowerShell), e.g.:
   - activates venv
   - installs/updates dependencies
   - runs the analysis script

### Phase B — Externalize configuration

Current script requires editing constants in source. Professional practice is to move configuration to:

- `config.yaml` (preferred), or
- command-line arguments.

Typical per-team config profiles:

- `config_default.yaml`
- `config_txpa.yaml`
- `config_dpll.yaml`

This avoids code edits and prevents accidental parameter drift between engineers.

### Phase C — Release and traceability

1. Version tags per release (`v1.0.0`, `v1.1.0`, ...).
2. Changelog for threshold/logic changes.
3. Release note template including:
   - data format compatibility
   - dependency changes
   - output schema changes

### Phase D — Quality gate (CI)

Add a small CI pipeline that runs on merge:

- Lint/format checks.
- Smoke test on a small sample CSV.
- Validation that output workbook is generated.

This prevents broken scripts from being distributed.

### Phase E — Distribution options

Choose one:

1. **Git + launcher script (fastest start)**
   - Best for engineering teams already using Git.
2. **Internal wheel package (`pip install`)**
   - Best for controlled version rollout.
3. **Single executable (PyInstaller)**
   - Best for users without Python setup.

For your current environment, option 1 is the shortest path to reliable team adoption.

---

## 9) Suggested team folder layout

```text
Tasks_Automation_Code/
  IFX_Scripts/
    Test_Data_Analysis/
      Tests_Data_Analysis.py
      Tests_Data_Analysis_User_Guide.md
      requirements-tests-data-analysis.txt
      run_tests_data_analysis.py
      run_tests_data_analysis.ps1
      configs/
        config_default.yaml
        config_txpa_focus.yaml
      tests/
        ...
```

---

## 9.1) Deployed artifacts included

- `requirements-tests-data-analysis.txt`
  - Pinned runtime dependencies.
- `configs/config_default.yaml`
  - Standard profile aligned with current script defaults.
- `configs/config_txpa_focus.yaml`
  - Focused profile for TXPA/TXPB/TXPC/TXPD analysis with correlation enabled.
- `run_tests_data_analysis.py`
  - YAML-driven wrapper that applies config and runs `Tests_Data_Analysis.py`.
- `run_tests_data_analysis.ps1`
  - Team launcher: creates `.venv` if needed, installs dependencies, and runs wrapper.
- `tests/smoke_input/smoke_Q2_sample.csv`
  - Tiny checked-in sample input used for end-to-end smoke validation.
- `tests/test_smoke_run_tests_data_analysis.py`
  - Automated smoke test that runs the YAML wrapper and verifies reports/plots are created.
- `tests/test_meta_parsing_and_status.py`
  - Focused unit tests for flat-file meta parsing and threshold/status classification.
- `.github/workflows/ifx-tests-data-analysis-smoke.yml`
  - CI workflow that runs the smoke test and unit tests on relevant changes.

## 9.2) Smoke test

Run the smoke test locally from repository root:

```text
python -m unittest discover -s Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/tests -p "test_*.py" -v
```

What it validates:

- YAML-driven launcher starts correctly.
- Sample input parses successfully.
- `Test_Data_Analysis_Report.xlsx` is created.
- `Correlation_Report.xlsx` is created.
- Plot PNGs are generated under `cdf_plots`.
- Parser/status unit tests stay green for key helper logic.

---

## 10) Troubleshooting

- **No CSV found**
  - Check `INPUT_FOLDER` and file extension `.csv`.
- **No module tests found**
  - Confirm `Test Name` row exists and module prefixes match `MODULES`.
- **Excel file cannot be overwritten**
  - Close workbook and rerun (script also creates timestamped fallback).
- **Plots missing**
  - Ensure `matplotlib` is installed.
- **Encoding issues**
  - Adjust `ENCODING` (default `latin1`).

---

## 11) Recommended next engineering improvements

1. Extend YAML config validation with stricter file/content checks and clearer remediation hints.
2. Expand unit-test coverage beyond meta parsing/status logic into plotting and correlation helpers.
3. Expand CI smoke coverage with additional sample inputs and richer output-content assertions.
4. Add release tagging + changelog automation.
5. Add optional executable packaging (PyInstaller) for non-Python users.

---

## 12) Quick start for engineers

1. Select a profile in `Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/configs`.
2. (Optional) set `single_file` in the selected YAML for first validation run.
3. Run `run_tests_data_analysis.ps1`.
4. Open `Test_Data_Analysis_Report.xlsx` in output folder.
5. Use `CDF Plot` links and `*_PLOTS` sheets for deep-dive review.
