# TX supply compensation simulation

This folder contains a small Monte‑Carlo simulation to compare:

- **Per-pair average compensation** (TX1&TX5, TX2&TX6, TX3&TX7, TX4&TX8)
- **One global average compensation** applied to all TX channels at once

## Run

From the repo root:

```powershell
C:/UserData/Learning/Software_Programming/GitHub_Nelio92/.venv/Scripts/python.exe `
  Tasks_Automation_Code/Reports/tx_supply_compensation_sim/simulate_tx_supply_compensation.py `
  --n-trials 10000 --seed 20260124 --acceptable-abs-residual-mv 20
```

Outputs are written into a dated folder (e.g. `output_YYYYMMDD/`).

## Model

Each TX channel has a required compensation offset (mV). If a method applies a different offset, the **residual error** is:

`residual = applied_offset - required_offset`

The script reports residual statistics (mean/RMS/max, percent exceeding thresholds) and generates plots.

It also reports acceptance metrics based on `--acceptable-abs-residual-mv` (default 20mV), such as:

- % of channels within the acceptance window
- % of trials where all 8 channels are within the acceptance window
