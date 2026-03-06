param(
    [string]$ConfigPath = "Tasks_Automation_Code/IFX_Scripts/configs/config_default.yaml",
    [switch]$SkipInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path (Join-Path $scriptRoot "../..")
Set-Location $repoRoot

$venvPython = Join-Path $repoRoot ".venv/Scripts/python.exe"
if (-not (Test-Path $venvPython)) {
    Write-Host "Creating virtual environment in .venv ..."
    python -m venv .venv
}

if (-not $SkipInstall) {
    Write-Host "Installing/updating dependencies ..."
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install -r "Tasks_Automation_Code/IFX_Scripts/requirements-tests-data-analysis.txt"
}

Write-Host "Running Tests_Data_Analysis with config: $ConfigPath"
& $venvPython "Tasks_Automation_Code/IFX_Scripts/run_tests_data_analysis.py" --config $ConfigPath
