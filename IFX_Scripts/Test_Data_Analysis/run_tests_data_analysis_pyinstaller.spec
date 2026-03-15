# -*- mode: python ; coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata

PROJECT_DIR = Path.cwd().resolve()
ENTRY_SCRIPT = PROJECT_DIR / "run_tests_data_analysis.py"
CONFIG_DIR = PROJECT_DIR / "configs"
USER_GUIDE = PROJECT_DIR / "Tests_Data_Analysis_User_Guide.md"

packages_with_dynamic_imports = [
    "yaml",
    "matplotlib",
    "openpyxl",
    "PIL",
    "pystdf",
]

metadata_packages = [
    "pandas",
    "numpy",
    "PyYAML",
    "matplotlib",
    "openpyxl",
    "pystdf",
]

_datas: list[tuple[str, str]] = []
_binaries: list[tuple[str, str]] = []
_hiddenimports: list[str] = []

for package_name in packages_with_dynamic_imports:
    datas, binaries, hiddenimports = collect_all(package_name)
    _datas += datas
    _binaries += binaries
    _hiddenimports += hiddenimports

for package_name in metadata_packages:
    _datas += copy_metadata(package_name)

_hiddenimports += [
    "numpy._core._exceptions",
    "numpy._core._multiarray_umath",
    "numpy._core._methods",
]

for config_path in sorted(CONFIG_DIR.glob("*.yaml")):
    _datas.append((str(config_path), "configs"))

_datas.append((str(USER_GUIDE), "."))

block_cipher = None

a = Analysis(
    [str(ENTRY_SCRIPT)],
    pathex=[str(PROJECT_DIR)],
    binaries=_binaries,
    datas=_datas,
    hiddenimports=sorted(set(_hiddenimports + ["yaml"])),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["scipy", "sklearn", "pytest", "hypothesis"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TestDataAnalysis",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
