TEST DATA REVIEWER - QUICK START
================================

This folder contains a ready-to-use executable:

  TestDataReviewer.exe

You do NOT need to install Python to use it.

--------------------------------------------------
1) BEFORE YOU START
--------------------------------------------------

Open one of the YAML config files in the configs folder and check:

- input_folder
- output_folder
- modules
- convert_stdf_before_analysis

Available example configs:
- configs\config.yaml

Tip:
If you are using a new dataset, first make a copy of a config file and edit the copy.

--------------------------------------------------
2) RUN THE TOOL
--------------------------------------------------

Open PowerShell in this folder and run:

  .\TestDataReviewer.exe --config .\configs\config.yaml

Or use another config:

  .\TestDataReviewer.exe --config .\configs\config_teamA.yaml

--------------------------------------------------
3) DRY-RUN FIRST (RECOMMENDED)
--------------------------------------------------

To validate the config without starting the analysis:

  .\TestDataReviewer.exe --config .\configs\config.yaml --dry-run

Use dry-run to confirm:
- the correct config is used
- the input path is correct
- the output path is correct
- STDF conversion is enabled/disabled as expected

--------------------------------------------------
4) TYPICAL USE CASES
--------------------------------------------------

A) Analyze CSV files directly
- put CSV files in the configured input folder
- set convert_stdf_before_analysis: false
- run the executable

B) Analyze STDF or EFF files directly
- put source files in the configured input folder
- set convert_stdf_before_analysis: true
- run the executable

C) Analyze only one STDF/EFF source file
- set stdf_single_file in the YAML config
- run the executable

--------------------------------------------------
5) OUTPUTS
--------------------------------------------------

The tool writes results into the configured output folder.

Typical outputs:
- Test_Data_Reviewer_Report.xlsx
- cdf_plots\...
- Correlation_Report.xlsx (if enabled)
- Artifacts\... (if STDF conversion is enabled)

--------------------------------------------------
6) IF SOMETHING GOES WRONG
--------------------------------------------------

First try:

  .\TestDataReviewer.exe --config .\configs\config.yaml --dry-run

Check for these common problems:
- config file path is wrong
- input_folder does not exist
- output_folder is invalid
- no matching CSV/STDF/EFF files are found
- Excel output file is already open

For more details, read:

  Test_Data_Reviewer_User_Guide.md