# TXPA Correlation Factors: -40°C Estimation Note (BE)

Date: 2026-02-06

This note documents a practical method to derive missing cold (-40°C) TXPA correlation factors from the available 25°C and 135°C factors in `TXPA_Corr_Factors_BE.xlsx`.

## Summary

- Available data: two temperature points per row (25°C and 135°C).
- Goal: estimate correlation factors at -40°C.
- Important: -40°C is outside the measured range [25°C, 135°C]. Any method is **extrapolation**, not interpolation.

Because only two temperature points are available, the functional relationship vs temperature cannot be validated (linearity/curvature is unknown). The approach below balances:
- stability for temperature-insensitive rows (don’t overfit), and
- robustness for temperature-sensitive rows (don’t rely on a single assumed model).

## Outputs

The analysis exports three formatted Excel workbooks (auto-fit column widths):

- `Tasks_Automation_Code/Reports/txpa_corr_factors_be_cold_estimates.xlsx`
  - Contains extracted factors at 25°C/135°C, multiple -40°C model estimates, an uncertainty envelope, and a single recommended -40°C factor.

- `Tasks_Automation_Code/Reports/txpa_corr_factors_be_cold_high_risk.xlsx`
  - Subset of rows flagged as high-risk for extrapolation (large temperature sensitivity and/or large model uncertainty).

- `Tasks_Automation_Code/Reports/txpa_corr_factors_be_group_summary.xlsx`
  - Grouped summary by (Voltage corner, Frequency_GHz, PA Channel).

Plots:

- `Tasks_Automation_Code/Reports/txpa_corr_factors_be_models_per_testcase.pdf`
  - One page per test case: measured points at 25°C/135°C plus the three model curves, and temperature-specific subplots (-40°C, 25°C, 135°C).

## Definitions

Let:
- $f_{25}$ = factor at 25°C
- $f_{135}$ = factor at 135°C
- $\Delta = f_{135} - f_{25}$
- $|\Delta|$ = absolute temperature sensitivity proxy (measured over 110°C)

### Temperature sensitivity classes

Based on the observed two-point change:
- **low**: $|\Delta| \le 0.05$
- **medium**: $0.05 < |\Delta| \le 0.2$
- **high**: $|\Delta| > 0.2$

These thresholds are heuristic and can be tightened/loosened depending on acceptable error.

## Recommended -40°C policy

### 1) Low sensitivity: use $f_{25}$

If $|\Delta| \le 0.05$, set:

- $f_{-40,rec} = f_{25}$

Rationale: the factor is already stable across 25°C→135°C, so extrapolating further adds model risk for minimal benefit.

In the export, this is recorded as:
- `f_-40_recommended_method = use_f25_low_sensitivity`

### 2) Medium/high sensitivity: use an ensemble across plausible two-point models

If $|\Delta| > 0.05$, compute three two-point extrapolations using (25°C, 135°C):

- **Model A (linear in T):** assume $f(T)$ is linear in °C.
- **Model B (linear in 1/K):** assume $f$ is linear vs $1/(T+273.15)$.
- **Model C (linear in log(K)):** assume $f$ is linear vs $\log(T+273.15)$.

Then set the recommended cold factor to the **median** across the three models:

- $f_{-40,rec} = \text{median}(f_{-40,A}, f_{-40,B}, f_{-40,C})$

Rationale: with only two points, model choice dominates uncertainty for temperature-sensitive rows; the median is a robust central estimate.

In the export, this is recorded as:
- `f_-40_recommended_method = ensemble_median_3models`

### Uncertainty envelope

To quantify model-form uncertainty, the export includes:

- `f_-40_models_min` = min across the three models
- `f_-40_models_max` = max across the three models
- `f_-40_models_envelope_width` = max - min

Large envelope width indicates that the -40°C estimate is highly dependent on the assumed model and should be treated as high-risk.

## High-risk screening (review list)

Rows are flagged for review if either:
- `abs_delta_135_25 > 0.2` (high temperature sensitivity), or
- `f_-40_models_envelope_width > 0.2` (high model uncertainty)

Those rows are exported to `txpa_corr_factors_be_cold_high_risk.xlsx`.

## When extrapolation may be unsuitable

Extrapolation is most questionable when one or more are true:
- Very large $|\Delta|$ (strong temperature dependence)
- Large model envelope width at -40°C (model-form uncertainty)
- Sign changes between 25°C and 135°C (non-monotonic or crossing behavior)

If such cases are critical to product performance, the best remedy is to measure correlation factors at -40°C for those specific conditions (targeted cold characterization) and re-fit a temperature model using ≥3 temperature points.

## How to use the results

- For automation: consume `f_-40_recommended` directly.
- For reporting/review: sort by `f_-40_models_envelope_width` and/or `abs_delta_135_25` and inspect the high-risk list.

## Regenerating the XLSX outputs

Run:

`C:/UserData/Learning/Software_Programming/GitHub_Nelio92/.venv/Scripts/python.exe Tasks_Automation_Code/Reports/generate_txpa_corr_factors_be_cold_estimates_xlsx.py`

To regenerate the plots PDF:

`C:/UserData/Learning/Software_Programming/GitHub_Nelio92/.venv/Scripts/python.exe Tasks_Automation_Code/Reports/generate_txpa_corr_factors_be_model_plots.py`

