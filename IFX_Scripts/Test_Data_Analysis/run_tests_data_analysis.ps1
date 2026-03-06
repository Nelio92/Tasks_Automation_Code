param(
    [string]$ConfigPath = "",
    [switch]$SkipInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-WorkspaceAnchor {
    param([string]$StartPath)

    $current = Resolve-Path $StartPath
    while ($true) {
        if ((Test-Path (Join-Path $current ".git")) -or (Test-Path (Join-Path $current ".venv"))) {
            return $current
        }

        $parent = Split-Path -Parent $current
        if (-not $parent -or $parent -eq $current) {
            return $StartPath
        }
        $current = $parent
    }
}

$workspaceRoot = Get-WorkspaceAnchor -StartPath $scriptRoot
Set-Location $workspaceRoot

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $scriptRoot "configs/config_default.yaml"
}
$ConfigPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $ConfigPath))

$venvPython = Join-Path $workspaceRoot ".venv/Scripts/python.exe"
if (-not (Test-Path $venvPython)) {
    Write-Host "Creating virtual environment in .venv ..."
    python -m venv .venv
}

if (-not $SkipInstall) {
    Write-Host "Installing/updating dependencies ..."
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install -r (Join-Path $scriptRoot "requirements-tests-data-analysis.txt")
}

Write-Host "Running Tests_Data_Analysis with config: $ConfigPath"
& $venvPython (Join-Path $scriptRoot "run_tests_data_analysis.py") --config $ConfigPath
