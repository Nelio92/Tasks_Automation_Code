param(
    [switch]$Clean,
    [switch]$SkipInstall,
    [string]$ReleaseDir = "release_pyinstaller"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = (Resolve-Path (Join-Path $scriptRoot "../../..")).Path
$venvPython = Join-Path $workspaceRoot ".venv/Scripts/python.exe"
$specPath = Join-Path $scriptRoot "run_tests_data_analysis_pyinstaller.spec"
$distRoot = Join-Path $scriptRoot "dist"
$buildRoot = Join-Path $scriptRoot "build"
$releaseRoot = Join-Path $scriptRoot $ReleaseDir
$exeSource = Join-Path $distRoot "TestDataAnalysis.exe"
$releaseExe = Join-Path $releaseRoot "TestDataAnalysis.exe"
$configReleaseDir = Join-Path $releaseRoot "configs"
$releaseReadme = Join-Path $scriptRoot "README.txt"

if (-not (Test-Path $venvPython)) {
    throw "Virtual environment python not found at $venvPython"
}

if ($Clean) {
    foreach ($path in @($buildRoot, $distRoot, $releaseRoot)) {
        if (Test-Path $path) {
            Remove-Item -Recurse -Force $path
        }
    }
}

if (-not $SkipInstall) {
    & $venvPython -m pip install --upgrade pip
    & $venvPython -m pip install -r (Join-Path $scriptRoot "requirements-tests-data-analysis.txt")
    & $venvPython -m pip install -r (Join-Path $scriptRoot "requirements-pyinstaller-build.txt")
}

Push-Location $scriptRoot
try {
    & $venvPython -m PyInstaller --noconfirm --clean $specPath
}
finally {
    Pop-Location
}

if (-not (Test-Path $exeSource)) {
    throw "PyInstaller build did not create expected executable: $exeSource"
}

New-Item -ItemType Directory -Force -Path $releaseRoot | Out-Null
New-Item -ItemType Directory -Force -Path $configReleaseDir | Out-Null
Copy-Item $exeSource $releaseExe -Force
Copy-Item (Join-Path $scriptRoot "configs/*.yaml") $configReleaseDir -Force
Copy-Item (Join-Path $scriptRoot "Tests_Data_Analysis_User_Guide.md") $releaseRoot -Force
if (Test-Path $releaseReadme) {
    Copy-Item $releaseReadme $releaseRoot -Force
}
$releaseLauncher = Join-Path $releaseRoot "run_tests_data_analysis.ps1"
if (Test-Path $releaseLauncher) {
    Remove-Item $releaseLauncher -Force
}

Write-Host "PyInstaller release created at: $releaseRoot"
Write-Host "Executable: $releaseExe"
Write-Host "Configs:    $configReleaseDir"
