# Test Data Analysis — User Guide

## 1) What this tool does

`TestDataAnalysis.exe` analyzes production test data and generates review files for Test Engineering.

Main outputs:
- `Test_Data_Analysis_Report.xlsx`
- `cdf_plots/...` PNG files
- optional `Correlation_Report.xlsx`
- optional STDF conversion artifacts under `Artifacts`

The tool is delivered as an executable. No Python installation is required for normal use.

---

## 2) Files included in this release folder

This release folder contains:

- `TestDataAnalysis.exe`
  - main executable
- `configs/`
  - YAML configuration profiles
- `Tests_Data_Analysis_User_Guide.md`
  - this guide

Available config profiles in this release:
- `configs/config.yaml`

---

## 3) What input data is supported

The tool can work with:
- flat production CSV files
- STDF/EFF source files when pre-conversion is enabled in the YAML config, including:
  - `.std`
  - `.stdf`
  - `.eff`
  - `.std.gz`, `.stdf.gz`
  - `.std.bz2`, `.stdf.bz2`
  - `.std.xz`, `.stdf.xz`
  - `.std.tar.gz`, `.stdf.tar.gz`

### Expected CSV structure

The CSV input is expected to be semicolon-separated and contain:
1. a header row with metadata columns and numeric test-number columns
2. metadata rows such as:
   - `Test Name`
   - `Low`
   - `High`
   - `Unit`
   - `Cpk`
   - `Yield`
3. device/unit rows after the metadata block

Important assumptions:
- delimiter is `;`
- numeric test columns use digit-only names
- module is inferred from the first 4 characters of `Test Name`

---

## 4) Outputs you will get

Typical outputs are written to the `output_folder` defined in the selected YAML config.

### Main report
- `Test_Data_Analysis_Report.xlsx`
  - one analysis sheet per input CSV
  - plot sheet(s) with embedded CDF and wafer map images
  - hyperlinks to plot locations when available

### Plot images
- `cdf_plots/<input_name>/...png`
  - CDF plots
  - wafer map plots when supported by the input

### Optional correlation report
- `Correlation_Report.xlsx`
  - generated only if enabled in the YAML config

### Optional STDF conversion artifacts
When STDF conversion is enabled, the tool can also create:
- generated CSV files in the configured input folder
- `Artifacts/...` files such as DTR sidecar CSV and consistency JSON reports

---

## 5) Before you run it

1. Put the release folder in a location you can access.
2. Open the YAML config you want to use.
3. Check these important values:
   - `input_folder`
   - `output_folder`
   - `modules`
   - `convert_stdf_before_analysis`
4. Make sure the input and output locations are valid for your machine.

Note:
- relative paths in the YAML file are resolved from the executable/release context
- absolute paths are recommended for shared team use

---

## 6) How to run the tool

Open PowerShell in the release folder and use one of the commands below.

### Run with the default profile

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml
```

### Run with an explicit config path

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml
```

### Dry-run only

This checks the config and prints the resolved settings without running the analysis.

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml --dry-run
```

---

## 7) Typical user cases

### User case 1 — Analyze a folder of flat CSV files

Use this when your input folder already contains flat production CSV files.

Recommended setup:
- `convert_stdf_before_analysis: false`
- `input_folder`: folder containing CSV files
- `output_folder`: folder where reports should be written

Run:

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml
```

Result:
- the tool reads CSV files directly
- generates the Excel analysis workbook
- generates CDF plots and wafer maps where applicable

---

### User case 2 — Analyze STDF or EFF files directly

Use this when your input folder contains STDF/EFF source files and you want the executable to convert them automatically first.

Recommended setup:
- `convert_stdf_before_analysis: true`
- `input_folder`: folder containing the source files
- optional `stdf_single_file`: only one source file to convert
- optional `stdf_file_patterns`: restrict which source files are picked up

Common examples for `stdf_single_file`:
- `lot123.std`
- `lot123.stdf.gz`
- `lot123.eff`
- `lot123.std.tar.gz`

Run:

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml
```

Result:
- STDF/EFF files are converted to flat CSV
- analysis runs on the generated CSV files
- conversion artifacts are written under `Artifacts`

---

### User case 3 — Run only one STDF/EFF source file

Use this when STDF/EFF pre-conversion is enabled and you want to convert and analyze only one source file.

If using STDF conversion, also optionally set:
- `stdf_single_file: your_file.std`

Supported examples:
- `stdf_single_file: your_file.eff`
- `stdf_single_file: your_file.std.tar.gz`

Then run the selected profile.

---

### User case 4 — Validate settings before a long run

Use dry-run before running a large dataset.

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml --dry-run
```

This helps confirm:
- correct config file is used
- input path is correct
- output path is correct
- selected modules are correct
- STDF conversion mode is correct

---

### User case 5 — Create a team-specific config copy

If you do not want to modify the shared config profiles, make your own copy.

Example:
- copy `configs/config.yaml`
- rename it to `configs/config_teamA.yaml`
- edit only the paths and thresholds you need

Run:

```powershell
./TestDataAnalysis.exe --config ./configs/config_teamA.yaml
```

---

## 8) Important config fields

You normally only need to edit a few YAML fields.

### Paths
- `input_folder`
  - where the tool reads CSV or STDF files
- `output_folder`
  - where the reports and plots are written

### Selection
- `modules`
  - list of module prefixes to include

### Thresholds
- `yield_threshold`
- `cpk_low`
- `cpk_high`

### STDF conversion
- `convert_stdf_before_analysis`
- `stdf_single_file`
- `stdf_file_patterns`

Default source discovery supports `.std`, `.stdf`, `.eff`, compressed variants, and `.tar.gz` packages containing a supported source file.

### Correlation report
- `generate_correlation_report`

---

## 9) How affected tests are selected

A test is included in the main report if at least one of these conditions is true:
- `Yield < yield_threshold`
- `Cpk < cpk_low`
- `Cpk > cpk_high`

Only tests whose module prefix is listed in `modules` are evaluated.

---

## 10) Interpretation notes

### Status
- `FAILS`
  - yield below threshold
- `Cpk<...`
  - capability too low
- `Cpk>...`
  - unusually high Cpk; review for data/spec issues

### Other useful fields
- **Fail Chips**
  - based on `Low` / `High` limits when available
- **Outliers**
  - based on a robust MAD method
- **Findings**
  - includes heuristic notes such as spread, multimodality, site behavior, wafer behavior, or coordinate signatures

---

## 11) Troubleshooting

### Nothing happens or the run exits early
- first run a dry-run:

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml --dry-run
```

### Config file not found
- make sure the file exists in `configs/`
- use the full config path if needed

### Input folder does not exist
- correct `input_folder` in the YAML file
- verify the folder path on your machine

### No CSV files found
- confirm the input folder contains flat CSV files
- if using STDF/EFF input, make sure `convert_stdf_before_analysis: true`

### No module tests found
- check the `modules` list in the YAML file
- confirm the module prefixes exist in the input data

### Excel file cannot be overwritten
- close the workbook and run again
- if needed, use a different `output_folder`

### Plots or correlation report are missing
- check whether the corresponding feature is enabled by the config
- correlation report is only created if `generate_correlation_report: true`

### STDF conversion did not run
- check `convert_stdf_before_analysis`
- verify the folder contains matching source files (`.std`, `.stdf`, `.eff`, supported compressed forms, or `.tar.gz` packages)
- verify `stdf_single_file` or `stdf_file_patterns` are not too restrictive

---

## 12) Recommended way to use this release

For normal users:
- do not edit the executable
- only edit or copy YAML config files
- run the executable with the selected config
- use dry-run first for new datasets

Recommended pattern:
1. copy a config
2. adjust input/output paths
3. run `--dry-run`
4. run the full analysis
5. review the generated workbook and plots

---

## 13) Example commands

### Default run

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml
```

### Dry-run

```powershell
./TestDataAnalysis.exe --config ./configs/config.yaml --dry-run
```

### Run with a custom config copy

```powershell
./TestDataAnalysis.exe --config ./configs/config_myproject.yaml
```

### Run from another folder using an absolute config path

```powershell
C:\path\to\release_pyinstaller\TestDataAnalysis.exe --config C:\path\to\release_pyinstaller\configs\config.yaml
```

---

## 14) Support note

If the tool fails even after a dry-run check, provide these items when asking for support:
- the exact command used
- the YAML config file used
- the console output/error text
- whether the input was CSV, STDF, or EFF
- one small representative sample file if allowed
