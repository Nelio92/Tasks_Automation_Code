# Test_Data_Reviewer test assets

This folder contains a tiny checked-in sample input and an end-to-end smoke test for `run_test_data_reviewer.py`.

## Local run

From the repository root:

```text
python -m unittest discover -s Tasks_Automation_Code/IFX_Scripts/Test_Data_Analysis/tests -p "test_*.py" -v
```

The test runs the launcher on `tests/smoke_input/smoke_Q2_sample.csv`, writes outputs to a temporary folder, and asserts that the Excel reports and plot PNGs are created.
