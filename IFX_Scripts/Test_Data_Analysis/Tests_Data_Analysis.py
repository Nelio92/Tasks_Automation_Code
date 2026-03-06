"""Flat CSV test data analysis + reporting.

This script is designed for large, semicolon-separated "flat" CSV exports where:
- The header row contains metadata columns (e.g. LOT/WAFER/X/Y/SITE_NUM) and many
  numeric test columns (e.g. 520123, 530045, ...).
- A small block of meta rows follows the header ("Test Name", "Low", "High",
  "Unit", "Cpk", "Yield", ...).
- Unit/device rows come after that meta block, with measurements per test column.

The report is module-agnostic: a test's "module" is extracted from the first 4
characters of its test name (e.g. DPLL, TXPA, TXLO). Modules to analyze are
provided via a user configuration section at the top of this file.

Outputs:
- One Excel workbook containing one sheet per input file with only the tests of
  interest that have yield < 100% or Cpk outside thresholds.
- A per-input-file plots sheet embedding CDF plots; hyperlinks in the data sheet
    jump to the embedded plot, and clicking an embedded plot opens the full PNG.
- Optionally, a separate correlation workbook (one sheet per input file)
  computing Pearson and Spearman correlations for each module test vs all tests.

Note on "hover previews": Excel doesn't support showing an image on hyperlink
hover via openpyxl. This implementation embeds images in a dedicated plots sheet
and links to those locations.
"""

from __future__ import annotations
import csv
import math
import os
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Literal
from xml.etree import ElementTree as ET


DEFAULT_ENCODING = "latin1"
DELIMITER = ";"


# ================================================
# INJECTED BY YAML LAUNCHER - DO NOT EDIT MANUALLY
# ================================================
# These values are intentionally populated by `run_tests_data_analysis.py`
# from a YAML config file before `run()` is called.

# Input/Output
INPUT_FOLDER: Path | None = None
OUTPUT_FOLDER: Path | None = None

# Which modules to analyze (first 4 chars of test name)
MODULES: list[str] = []

# Thresholds
YIELD_THRESHOLD: float | None = None
CPK_LOW: float | None = None
CPK_HIGH: float | None = None

# Outlier detection (|x-median| > OUTLIER_MAD_MULTIPLIER * MAD)
OUTLIER_MAD_MULTIPLIER: float | None = None

# Optional controls
MAX_FILES: int | None = None
SINGLE_FILE: str | None = None
ENCODING: str | None = None

# Optional correlation workbook
GENERATE_CORRELATION_REPORT: bool | None = None
CORRELATION_METHODS: list[Literal["pearson", "spearman"]] = []
PEARSON_ABS_MIN_FOR_REPORT: float | None = None

# Wafer map display controls
# Scales the *area* of the wafer outline circle; 2.0 => 2× area, 3.0 => 3× area.
WAFERMAP_CIRCLE_AREA_MULT: float | None = None


def _require_runtime_configuration() -> dict[str, Any]:
    missing: list[str] = []

    if INPUT_FOLDER is None:
        missing.append("INPUT_FOLDER")
    if OUTPUT_FOLDER is None:
        missing.append("OUTPUT_FOLDER")
    if not MODULES:
        missing.append("MODULES")
    if YIELD_THRESHOLD is None:
        missing.append("YIELD_THRESHOLD")
    if CPK_LOW is None:
        missing.append("CPK_LOW")
    if CPK_HIGH is None:
        missing.append("CPK_HIGH")
    if OUTLIER_MAD_MULTIPLIER is None:
        missing.append("OUTLIER_MAD_MULTIPLIER")
    if ENCODING is None:
        missing.append("ENCODING")
    if GENERATE_CORRELATION_REPORT is None:
        missing.append("GENERATE_CORRELATION_REPORT")
    if PEARSON_ABS_MIN_FOR_REPORT is None:
        missing.append("PEARSON_ABS_MIN_FOR_REPORT")
    if WAFERMAP_CIRCLE_AREA_MULT is None:
        missing.append("WAFERMAP_CIRCLE_AREA_MULT")
    if GENERATE_CORRELATION_REPORT and not CORRELATION_METHODS:
        missing.append("CORRELATION_METHODS")

    if missing:
        joined = ", ".join(missing)
        raise RuntimeError(
            "Tests_Data_Analysis runtime configuration is missing. "
            "Run via run_tests_data_analysis.py with a YAML config file. "
            f"Missing: {joined}"
        )

    return {
        "input_folder": _as_path(INPUT_FOLDER),
        "output_folder": _as_path(OUTPUT_FOLDER),
        "modules": list(MODULES),
        "yield_threshold": float(YIELD_THRESHOLD),
        "cpk_low": float(CPK_LOW),
        "cpk_high": float(CPK_HIGH),
        "outlier_mad_multiplier": float(OUTLIER_MAD_MULTIPLIER),
        "max_files": MAX_FILES,
        "single_file": SINGLE_FILE,
        "encoding": str(ENCODING),
        "generate_correlation_report": bool(GENERATE_CORRELATION_REPORT),
        "correlation_methods": list(CORRELATION_METHODS),
        "pearson_abs_min_for_report": float(PEARSON_ABS_MIN_FOR_REPORT),
        "wafermap_circle_area_mult": float(WAFERMAP_CIRCLE_AREA_MULT),
    }


def _as_path(value: str | os.PathLike[str] | Path) -> Path:
    return value if isinstance(value, Path) else Path(value)


def _excel_col_letter(col_idx_1_based: int) -> str:
    """Convert a 1-based column index to an Excel column letter."""
    if col_idx_1_based < 1:
        raise ValueError("Column index must be >= 1")
    n = col_idx_1_based
    letters: list[str] = []
    while n:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def _safe_sheet_name(name: str) -> str:
    # Excel sheet limit is 31 chars and cannot contain some characters.
    safe = re.sub(r"[\\/*?:\[\]]", "_", name)
    safe = safe.strip() or "Sheet"
    return safe[:31]


def _unique_sheet_name(name: str, existing_names) -> str:
    """Return an Excel-safe unique sheet name capped at 31 characters."""
    base = _safe_sheet_name(name)
    existing_lower = {str(item).lower() for item in existing_names}
    if base.lower() not in existing_lower:
        return base

    counter = 2
    while True:
        suffix = f"_{counter}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}".strip() or f"Sheet{suffix}"
        if candidate.lower() not in existing_lower:
            return candidate
        counter += 1


def _excel_internal_sheet_ref(sheet_name: str) -> str:
    """Return a sheet reference safe for internal Excel hyperlinks.

    Always quote sheet names to support characters like '-' and spaces.
    """
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'"


def _parse_excel_internal_target(target: str) -> tuple[str | None, str | None]:
    """Parse internal Excel hyperlink target like #'Sheet Name'!B4."""
    if not target or not str(target).startswith("#"):
        return None, None

    body = str(target)[1:]
    if "!" not in body:
        return None, None

    raw_sheet, raw_cell = body.split("!", 1)
    sheet = raw_sheet
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, raw_cell


def _is_valid_excel_a1_ref(cell_ref: str) -> bool:
    """Return True for basic A1 references/ranges used in internal hyperlinks."""
    if not cell_ref:
        return False
    pattern = r"^\$?[A-Z]{1,3}\$?[1-9][0-9]*(:\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$"
    return re.match(pattern, str(cell_ref).upper()) is not None


def _self_check_workbook_internal_hyperlinks(workbook) -> list[str]:
    """Validate all internal (#...) hyperlinks in workbook and return issues."""
    issues: list[str] = []
    sheet_names = set(workbook.sheetnames)

    for ws in workbook.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                hl = cell.hyperlink
                if hl is None:
                    continue
                target = getattr(hl, "target", None)
                if target is None or not str(target).startswith("#"):
                    continue

                sheet_name, cell_ref = _parse_excel_internal_target(str(target))
                if not sheet_name or sheet_name not in sheet_names:
                    issues.append(
                        f"{ws.title}!{cell.coordinate}: invalid target sheet in hyperlink '{target}'"
                    )
                    continue
                if not cell_ref or not _is_valid_excel_a1_ref(cell_ref):
                    issues.append(
                        f"{ws.title}!{cell.coordinate}: invalid target cell ref in hyperlink '{target}'"
                    )

    return issues


def _normalize_ooxml_path(path: str) -> str:
    parts: list[str] = []
    for part in PurePosixPath(path).parts:
        if part in ("", ".", "/", "\\"):
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def _resolve_ooxml_target(source_part: str, target: str) -> str:
    return _normalize_ooxml_path(str(PurePosixPath(source_part).parent / target))


def _rels_part_for(part_path: str) -> str:
    part = PurePosixPath(part_path)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def _next_relationship_id(rels_root: ET.Element) -> str:
    used: set[int] = set()
    for rel in rels_root:
        rid = rel.attrib.get("Id", "")
        m = re.fullmatch(r"rId(\d+)", rid)
        if m:
            used.add(int(m.group(1)))

    next_id = 1
    while next_id in used:
        next_id += 1
    return f"rId{next_id}"


def _enable_clickable_plot_images(workbook_path: Path, sheet_image_targets: dict[str, list[str]]) -> None:
    """Patch workbook drawings so clicking an embedded plot opens the PNG file."""
    if not sheet_image_targets:
        return

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns_xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    drawing_rel_type = f"{ns_rel}/drawing"
    hyperlink_rel_type = f"{ns_rel}/hyperlink"

    ET.register_namespace("", ns_main)
    ET.register_namespace("r", ns_rel)
    ET.register_namespace("xdr", ns_xdr)
    ET.register_namespace("a", ns_a)

    with zipfile.ZipFile(workbook_path, "r") as zin:
        archive = {name: zin.read(name) for name in zin.namelist()}

    workbook_part = "xl/workbook.xml"
    workbook_rels_part = "xl/_rels/workbook.xml.rels"
    if workbook_part not in archive or workbook_rels_part not in archive:
        return

    workbook_root = ET.fromstring(archive[workbook_part])
    workbook_rels_root = ET.fromstring(archive[workbook_rels_part])
    workbook_rels_by_id = {
        rel.attrib.get("Id"): rel for rel in workbook_rels_root.findall(f"{{{ns_pkg}}}Relationship")
    }

    sheet_part_by_name: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{ns_main}}}sheet"):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{ns_rel}}}id")
        rel = workbook_rels_by_id.get(rel_id)
        if not name or rel is None:
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        sheet_part_by_name[name] = _resolve_ooxml_target(workbook_part, target)

    modified: dict[str, bytes] = {}
    for sheet_name, image_targets in sheet_image_targets.items():
        if not image_targets:
            continue

        sheet_part = sheet_part_by_name.get(sheet_name)
        if not sheet_part:
            continue
        sheet_rels_part = _rels_part_for(sheet_part)
        if sheet_rels_part not in archive:
            continue

        sheet_rels_root = ET.fromstring(archive[sheet_rels_part])
        drawing_target = None
        for rel in sheet_rels_root.findall(f"{{{ns_pkg}}}Relationship"):
            if rel.attrib.get("Type") == drawing_rel_type:
                drawing_target = rel.attrib.get("Target")
                break
        if not drawing_target:
            continue

        drawing_part = _resolve_ooxml_target(sheet_part, drawing_target)
        drawing_rels_part = _rels_part_for(drawing_part)
        if drawing_part not in archive or drawing_rels_part not in archive:
            continue

        drawing_root = ET.fromstring(archive[drawing_part])
        drawing_rels_root = ET.fromstring(archive[drawing_rels_part])

        pic_anchors: list[ET.Element] = []
        for anchor_tag in ("oneCellAnchor", "twoCellAnchor", "absoluteAnchor"):
            for anchor in drawing_root.findall(f"{{{ns_xdr}}}{anchor_tag}"):
                if anchor.find(f"{{{ns_xdr}}}pic") is not None:
                    pic_anchors.append(anchor)

        changed = False
        for target_uri, anchor in zip(image_targets, pic_anchors, strict=False):
            pic = anchor.find(f"{{{ns_xdr}}}pic")
            if pic is None:
                continue
            nv_pic = pic.find(f"{{{ns_xdr}}}nvPicPr")
            if nv_pic is None:
                continue
            c_nv_pr = nv_pic.find(f"{{{ns_xdr}}}cNvPr")
            if c_nv_pr is None:
                continue

            for child in list(c_nv_pr):
                if child.tag == f"{{{ns_a}}}hlinkClick":
                    c_nv_pr.remove(child)

            rel_id = _next_relationship_id(drawing_rels_root)
            ET.SubElement(
                drawing_rels_root,
                f"{{{ns_pkg}}}Relationship",
                {
                    "Id": rel_id,
                    "Type": hyperlink_rel_type,
                    "Target": target_uri,
                    "TargetMode": "External",
                },
            )
            ET.SubElement(
                c_nv_pr,
                f"{{{ns_a}}}hlinkClick",
                {
                    f"{{{ns_rel}}}id": rel_id,
                    "tooltip": "Open full-size PNG",
                },
            )
            changed = True

        if changed:
            modified[drawing_part] = ET.tostring(drawing_root, encoding="utf-8", xml_declaration=True)
            modified[drawing_rels_part] = ET.tostring(drawing_rels_root, encoding="utf-8", xml_declaration=True)

    if not modified:
        return

    temp_path = workbook_path.with_suffix(workbook_path.suffix + ".tmp")
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in archive.items():
            zout.writestr(name, modified.get(name, data))

    temp_path.replace(workbook_path)


def _autofit_openpyxl_columns(ws, *, min_width: int = 8, max_width: int = 70, padding: int = 2) -> None:
    if ws.max_column < 1 or ws.max_row < 1:
        return

    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column),
        start=1,
    ):
        max_len = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            text = str(cell.value)
            if "\n" in text:
                text = max(text.splitlines(), key=len)
            max_len = max(max_len, len(text))

        width = max(min_width, min(max_width, max_len + padding))
        ws.column_dimensions[_excel_col_letter(col_idx)].width = width


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().strip('"')
    if text == "" or text.lower() in {"nan", "na", "none"}:
        return None
    try:
        return float(text)
    except ValueError:
        # Handle comma decimals.
        try:
            return float(text.replace(",", "."))
        except ValueError:
            return None


def _module_from_test_name(test_name: str) -> str:
    if not test_name:
        return ""
    s = str(test_name).strip()
    if len(s) < 4:
        return s.upper()
    return s[:4].upper()


@dataclass(frozen=True)
class FlatFileMeta:
    header: list[str]
    numeric_test_cols: list[str]
    data_start_line_index: int  # 0-based line index where unit data begins
    meta_rows: dict[str, dict[str, str]]  # row_name -> {col_name -> raw string}


def scan_flat_file_meta(
    file_path: Path,
    *,
    encoding: str = DEFAULT_ENCODING,
    delimiter: str = DELIMITER,
    needed_meta_rows: Iterable[str] = ("Test Name", "Low", "High", "Unit", "Cpk", "Yield", "Mean", "Stddev"),
    max_scan_lines: int = 200,
) -> FlatFileMeta:
    needed = {r.strip() for r in needed_meta_rows}
    meta_rows: dict[str, dict[str, str]] = {}

    with file_path.open("r", encoding=encoding, errors="replace", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        try:
            header = next(reader)
        except StopIteration:
            raise ValueError(f"Empty file: {file_path}")

        header = [h.strip() for h in header]
        numeric_test_cols = [h for h in header if h.strip().isdigit()]
        if not numeric_test_cols:
            raise ValueError("Could not find numeric test columns in header")

        data_start_line_index = 1
        for line_idx, row in enumerate(reader, start=1):
            if line_idx >= max_scan_lines:
                break
            if not row:
                continue
            key = (row[0] or "").strip().strip('"')
            if key.isdigit():
                data_start_line_index = line_idx
                break

            if key in needed:
                row_map: dict[str, str] = {}
                # Zip may truncate; that's ok.
                for col_name, cell in zip(header, row, strict=False):
                    if col_name in numeric_test_cols:
                        row_map[col_name] = (cell or "").strip()
                meta_rows[key] = row_map

        else:
            # If we didn't break on a unit row, we still set start line to after scanned block.
            data_start_line_index = max_scan_lines

    return FlatFileMeta(
        header=header,
        numeric_test_cols=numeric_test_cols,
        data_start_line_index=data_start_line_index,
        meta_rows=meta_rows,
    )


def _mad(values):
    import numpy as np

    med = np.nanmedian(values)
    return float(np.nanmedian(np.abs(values - med)))


def _robust_sigma(values) -> float:
    import numpy as np

    m = _mad(values)
    if m > 0:
        return 1.4826 * m
    # Fallback to std if MAD collapses.
    return float(np.nanstd(values, ddof=1))


def _count_hist_peaks(values) -> int:
    """Very lightweight multimodality heuristic based on histogram peaks."""
    import numpy as np

    v = values[np.isfinite(values)]
    if v.size < 80:
        return 1
    q25, q75 = np.percentile(v, [25, 75])
    iqr = q75 - q25
    if iqr <= 0:
        return 1
    bin_width = 2 * iqr / (v.size ** (1 / 3))
    if not np.isfinite(bin_width) or bin_width <= 0:
        return 1
    bins = max(10, min(80, int((v.max() - v.min()) / bin_width) + 1))
    hist, _ = np.histogram(v, bins=bins)
    if hist.size < 5:
        return 1

    # Smooth with a small moving average.
    kernel = np.array([1, 2, 3, 2, 1], dtype=float)
    kernel /= kernel.sum()
    smooth = np.convolve(hist.astype(float), kernel, mode="same")
    # Local maxima with a basic prominence threshold.
    threshold = max(5.0, 0.10 * float(smooth.max()))
    peaks = 0
    for i in range(1, len(smooth) - 1):
        if smooth[i] > smooth[i - 1] and smooth[i] > smooth[i + 1] and smooth[i] >= threshold:
            peaks += 1
    return max(1, peaks)


def _cdf_plot_png(
    values,
    *,
    title: str,
    out_path: Path,
    low_limit: float | None = None,
    high_limit: float | None = None,
    proposed_l6: float | None = None,
    proposed_u6: float | None = None,
    proposed_l12: float | None = None,
    proposed_u12: float | None = None,
) -> None:
    import pandas as pd
    import numpy as np

    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        # Plotting is optional; if matplotlib isn't available just skip.
        return

    v_raw = pd.to_numeric(values, errors="coerce").to_numpy(dtype=float)
    finite_mask = np.isfinite(v_raw)
    v = v_raw[finite_mask]
    if v.size == 0:
        return

    low = -np.inf if low_limit is None else float(low_limit)
    high = np.inf if high_limit is None else float(high_limit)
    has_spec = low_limit is not None or high_limit is not None
    fail_mask = (v < low) | (v > high) if has_spec else np.zeros(v.size, dtype=bool)
    n_fail = int(np.count_nonzero(fail_mask))

    order = np.argsort(v)
    v = v[order]
    fail_mask = fail_mask[order]
    y = np.arange(1, v.size + 1) / v.size

    mean_v = float(np.mean(v))
    median_v = float(np.median(v))

    fig, ax = plt.subplots(figsize=(7.0, 4.0), dpi=140)

    # CDF points (not a continuous line), with failing chips highlighted.
    pass_mask = ~fail_mask
    if np.any(pass_mask):
        ax.plot(
            v[pass_mask],
            y[pass_mask],
            linestyle="None",
            marker=".",
            markersize=3.0,
            alpha=0.80,
            color="#1F77B4",
            label=f"Pass chips={int(np.count_nonzero(pass_mask))}",
        )
    if np.any(fail_mask):
        ax.plot(
            v[fail_mask],
            y[fail_mask],
            linestyle="None",
            marker="o",
            markersize=4.4,
            alpha=0.95,
            markerfacecolor="#D62728",
            markeredgecolor="#D62728",
            label=f"Fail chips={n_fail}",
        )
    elif not np.any(pass_mask):
        ax.plot(v, y, linestyle="None", marker=".", markersize=3.0, alpha=0.85, label="Data")

    # Spec limits (red)
    if low_limit is not None and np.isfinite(low_limit):
        ax.axvline(float(low_limit), color="#D62728", linestyle="-", linewidth=1.6, label=f"LTL={_fmt_num(float(low_limit))}")
    if high_limit is not None and np.isfinite(high_limit):
        ax.axvline(float(high_limit), color="#D62728", linestyle="-", linewidth=1.6, label=f"UTL={_fmt_num(float(high_limit))}")

    # Proposed 6σ/12σ limits for Cpk-issue tests (orange/purple)
    if proposed_l6 is not None and np.isfinite(proposed_l6):
        ax.axvline(float(proposed_l6), color="#FF7F0E", linestyle="--", linewidth=1.2, label=f"LTL 6s={_fmt_1dp(float(proposed_l6))}")
    if proposed_u6 is not None and np.isfinite(proposed_u6):
        ax.axvline(float(proposed_u6), color="#FF7F0E", linestyle="--", linewidth=1.2, label=f"UTL 6s={_fmt_1dp(float(proposed_u6))}")
    if proposed_l12 is not None and np.isfinite(proposed_l12):
        ax.axvline(float(proposed_l12), color="#9467BD", linestyle=":", linewidth=1.2, label=f"LTL 12s={_fmt_1dp(float(proposed_l12))}")
    if proposed_u12 is not None and np.isfinite(proposed_u12):
        ax.axvline(float(proposed_u12), color="#9467BD", linestyle=":", linewidth=1.2, label=f"UTL 12s={_fmt_1dp(float(proposed_u12))}")

    ax.grid(True, alpha=0.3)
    if has_spec:
        ax.set_title(title + f"\nFail chips: {n_fail}/{v.size}")
    else:
        ax.set_title(title + f"\nFail chips: N/A (no spec limits)")
    ax.set_xlabel("Value")
    ax.set_ylabel("CDF")
    ax.legend(loc="best", fontsize=8, framealpha=0.9)
    fig.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, format="png")
    plt.close(fig)


def _prepare_wafer_map_frame(values, *, meta_cols):
    import pandas as pd
    import numpy as np

    if meta_cols is None:
        return None, None, None, None
    if not all(c in getattr(meta_cols, "columns", []) for c in ("X", "Y")):
        return None, None, None, None

    v = pd.to_numeric(values, errors="coerce")
    x = pd.to_numeric(meta_cols["X"], errors="coerce")
    y = pd.to_numeric(meta_cols["Y"], errors="coerce")
    wafer = _normalize_wafer_ids(meta_cols["WAFER"]) if "WAFER" in meta_cols.columns else None

    df = pd.DataFrame({"v": v, "X": x, "Y": y})
    if wafer is not None:
        df["WAFER"] = wafer
    for optional_col in ("CHIP_ID", "SITE_NUM", "PF", "FIRST_FAIL_TEST"):
        if optional_col in getattr(meta_cols, "columns", []):
            df[optional_col] = meta_cols[optional_col]

    df = df.dropna(subset=["v", "X", "Y"]).copy()
    if df.empty:
        return None, None, None, None

    if "WAFER" in df.columns and df["WAFER"].notna().any():
        counts = df.dropna(subset=["WAFER"]).groupby("WAFER")["v"].size().sort_values(ascending=False)
        wafers = [str(w) for w in counts.index.tolist()]
    else:
        wafers = ["ALL"]
        df["WAFER"] = "ALL"

    wafers = wafers[:6]
    all_v = df["v"].to_numpy(dtype=float)
    vmin = float(np.nanpercentile(all_v, 1))
    vmax = float(np.nanpercentile(all_v, 99))
    if not np.isfinite(vmin) or not np.isfinite(vmax) or vmin >= vmax:
        vmin = float(np.nanmin(df["v"]))
        vmax = float(np.nanmax(df["v"]))

    return df, wafers, vmin, vmax


def _wafer_map_html(
    values,
    *,
    meta_cols,
    title: str,
    out_path: Path,
    low_limit: float | None = None,
    high_limit: float | None = None,
) -> None:
    """Create an interactive wafer map HTML with hover details for each chip."""
    import numpy as np

    df, wafers, vmin, vmax = _prepare_wafer_map_frame(values, meta_cols=meta_cols)
    if df is None or wafers is None:
        return

    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
    except Exception:
        return

    low = -np.inf if low_limit is None else float(low_limit)
    high = np.inf if high_limit is None else float(high_limit)

    n = len(wafers)
    ncols = 3 if n >= 3 else n
    nrows = int(math.ceil(n / max(1, ncols)))
    subplot_titles: list[str] = []
    wafer_frames: list[tuple[str, Any]] = []
    for w in wafers:
        d = df[df["WAFER"].astype(str) == str(w)].copy()
        wafer_frames.append((str(w), d))
        vv = d["v"].to_numpy(dtype=float)
        fails = (vv < low) | (vv > high)
        subplot_titles.append(f"WAFER={w}  N={len(vv)}  fails={int(np.count_nonzero(fails))}")

    fig = make_subplots(rows=nrows, cols=ncols, subplot_titles=subplot_titles)
    showscale_remaining = True

    for idx, (w, d) in enumerate(wafer_frames, start=1):
        row = (idx - 1) // ncols + 1
        col = (idx - 1) % ncols + 1
        xv = d["X"].to_numpy(dtype=float)
        yv = d["Y"].to_numpy(dtype=float)
        vv = d["v"].to_numpy(dtype=float)
        fails = (vv < low) | (vv > high)

        x_min = float(np.nanmin(xv))
        x_max = float(np.nanmax(xv))
        y_min = float(np.nanmin(yv))
        y_max = float(np.nanmax(yv))
        if x_min == x_max:
            x_min -= 0.5
            x_max += 0.5
        if y_min == y_max:
            y_min -= 0.5
            y_max += 0.5

        point_count = max(1, len(vv))
        marker_size = float(min(42.0, max(12.0, 300.0 / math.sqrt(point_count))))
        chip_id_series = d["CHIP_ID"].astype(str) if "CHIP_ID" in d.columns else None
        site_series = d["SITE_NUM"].astype(str) if "SITE_NUM" in d.columns else None

        hover_text = []
        for i in range(len(d)):
            chip_line = ""
            if chip_id_series is not None and chip_id_series.iloc[i] not in {"", "<NA>", "nan", "None"}:
                chip_line = f"<br>Chip ID={chip_id_series.iloc[i]}"
            site_line = ""
            if site_series is not None and site_series.iloc[i] not in {"", "<NA>", "nan", "None"}:
                site_line = f"<br>Site={site_series.iloc[i]}"
            hover_text.append(
                "<br>".join(
                    [
                        f"Wafer={w}",
                        f"X={_fmt_num(float(xv[i]))}",
                        f"Y={_fmt_num(float(yv[i]))}",
                        f"Value={_fmt_num(float(vv[i]))}",
                        f"Status={'FAIL' if bool(fails[i]) else 'PASS'}",
                    ]
                )
                + chip_line
                + site_line
            )

        fig.add_trace(
            go.Scatter(
                x=xv,
                y=yv,
                mode="markers",
                text=hover_text,
                hovertemplate="%{text}<extra></extra>",
                showlegend=False,
                marker={
                    "symbol": "square",
                    "size": marker_size,
                    "color": vv,
                    "colorscale": "Turbo",
                    "cmin": vmin,
                    "cmax": vmax,
                    "line": {"color": "#2F2F2F", "width": 0.8},
                    "colorbar": {"title": "Value"},
                    "showscale": showscale_remaining,
                },
            ),
            row=row,
            col=col,
        )
        showscale_remaining = False

        if np.any(fails):
            fig.add_trace(
                go.Scatter(
                    x=xv[fails],
                    y=yv[fails],
                    mode="markers",
                    hoverinfo="skip",
                    showlegend=False,
                    marker={
                        "symbol": "square-open",
                        "size": marker_size + 6,
                        "color": "#D62728",
                        "line": {"color": "#D62728", "width": 2.0},
                    },
                ),
                row=row,
                col=col,
            )

        xref = f"x{idx}" if idx > 1 else "x"
        yref = f"y{idx}" if idx > 1 else "y"
        fig.add_shape(
            type="circle",
            xref=xref,
            yref=yref,
            x0=x_min,
            x1=x_max,
            y0=y_min,
            y1=y_max,
            line={"color": "#666666", "width": 1.5},
        )
        fig.update_xaxes(range=[x_min, x_max], showgrid=True, gridcolor="rgba(0,0,0,0.12)", row=row, col=col)
        fig.update_yaxes(
            range=[y_min, y_max],
            showgrid=True,
            gridcolor="rgba(0,0,0,0.12)",
            scaleratio=1,
            scaleanchor=xref,
            row=row,
            col=col,
        )

    fig.update_layout(
        title={"text": title, "x": 0.5},
        template="plotly_white",
        dragmode="zoom",
        hovermode="closest",
        width=max(1200, 480 * ncols),
        height=max(700, 420 * nrows),
        margin={"l": 40, "r": 40, "t": 90, "b": 40},
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.write_html(out_path, include_plotlyjs=True, full_html=True, auto_open=False)


def _wafer_map_png(
    values,
    *,
    meta_cols,
    title: str,
    out_path: Path,
    low_limit: float | None = None,
    high_limit: float | None = None,
) -> None:
    """Create a wafer map scatter plot using X/Y and (optionally) WAFER.

    Colors show the test value distribution; spec-fail points are highlighted.
    """
    import pandas as pd
    import numpy as np

    if meta_cols is None:
        return
    if not all(c in getattr(meta_cols, "columns", []) for c in ("X", "Y")):
        return

    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import Ellipse
        import matplotlib.colors as mcolors
    except Exception:
        return

    import warnings

    v = pd.to_numeric(values, errors="coerce")
    x = pd.to_numeric(meta_cols["X"], errors="coerce")
    y = pd.to_numeric(meta_cols["Y"], errors="coerce")
    wafer = _normalize_wafer_ids(meta_cols["WAFER"]) if "WAFER" in meta_cols.columns else None

    df = pd.DataFrame({"v": v, "X": x, "Y": y})
    if wafer is not None:
        df["WAFER"] = wafer

    df = df.dropna(subset=["v", "X", "Y"]).copy()
    if df.empty:
        return

    # Choose wafers to plot.
    if "WAFER" in df.columns and df["WAFER"].notna().any():
        counts = df.dropna(subset=["WAFER"]).groupby("WAFER")["v"].size().sort_values(ascending=False)
        wafers = [str(w) for w in counts.index.tolist()]
    else:
        wafers = ["ALL"]
        df["WAFER"] = "ALL"

    max_wafers = 6
    wafers = wafers[:max_wafers]

    # Shared color scaling across subplots.
    all_v = df["v"].to_numpy(dtype=float)
    vmin = float(np.nanpercentile(all_v, 1))
    vmax = float(np.nanpercentile(all_v, 99))
    if not np.isfinite(vmin) or not np.isfinite(vmax) or vmin >= vmax:
        vmin = float(np.nanmin(df["v"]))
        vmax = float(np.nanmax(df["v"]))

    # Vivid multi-color gradient that makes extremes pop.
    try:
        cmap = plt.get_cmap("turbo")
    except Exception:
        cmap = plt.get_cmap("viridis")

    norm = mcolors.Normalize(vmin=vmin, vmax=vmax, clip=True)
    wafermap_scale = 3.0

    n = len(wafers)
    ncols = 3 if n >= 3 else n
    nrows = int(math.ceil(n / max(1, ncols)))
    fig_w = (10.0 if n <= 2 else 12.5) * wafermap_scale
    fig_h = (6.8 if n <= 2 else max(7.0, 4.8 * nrows)) * wafermap_scale
    fig, axes = plt.subplots(nrows=nrows, ncols=ncols, figsize=(fig_w, fig_h), dpi=140)
    if not isinstance(axes, np.ndarray):
        axes = np.array([axes])
    axes = axes.ravel()

    # Reserve a right margin for the colorbar so it never covers the wafer maps.
    fig.subplots_adjust(left=0.06, right=0.86, bottom=0.08, top=0.90, wspace=0.18, hspace=0.28)

    low = -np.inf if low_limit is None else float(low_limit)
    high = np.inf if high_limit is None else float(high_limit)

    mappable = None
    for ax, w in zip(axes, wafers, strict=False):
        d = df[df["WAFER"].astype(str) == str(w)]
        if d.empty:
            ax.axis("off")
            continue

        xv = d["X"].to_numpy(dtype=float)
        yv = d["Y"].to_numpy(dtype=float)
        vv = d["v"].to_numpy(dtype=float)

        fails = (vv < low) | (vv > high)
        n_fail = int(np.count_nonzero(fails))

        point_count = max(1, len(vv))
        base_chip_marker_size = float(min(165.0, max(55.0, 2200.0 / math.sqrt(point_count))))
        chip_marker_size = base_chip_marker_size * wafermap_scale
        fail_marker_size = chip_marker_size * 1.8

        sc = ax.scatter(
            xv,
            yv,
            c=vv,
            cmap=cmap,
            norm=norm,
            s=chip_marker_size,
            marker="s",
            linewidths=0.20 * wafermap_scale,
            edgecolors="#2F2F2F",
            alpha=0.98,
        )
        mappable = sc

        if np.any(fails):
            ax.scatter(
                xv[fails],
                yv[fails],
                facecolors="none",
                edgecolors="#D62728",
                linewidths=1.6 * wafermap_scale,
                s=fail_marker_size,
                marker="s",
                label=f"fails={n_fail}",
            )

        # Wafer outline based on the visible chip coordinate envelope.
        x_min = float(np.nanmin(xv))
        x_max = float(np.nanmax(xv))
        y_min = float(np.nanmin(yv))
        y_max = float(np.nanmax(yv))

        if x_min == x_max:
            x_min -= 0.5
            x_max += 0.5
        if y_min == y_max:
            y_min -= 0.5
            y_max += 0.5

        cx = 0.5 * (x_min + x_max)
        cy = 0.5 * (y_min + y_max)
        ax.add_patch(
            Ellipse(
                (cx, cy),
                width=(x_max - x_min),
                height=(y_max - y_min),
                fill=False,
                color="#666666",
                linewidth=0.8 * wafermap_scale,
                alpha=0.8,
            )
        )

        ax.set_xlim(x_min, x_max)
        ax.set_ylim(y_min, y_max)

        ax.set_title(f"WAFER={w}  N={len(vv)}  fails={n_fail}", fontsize=10 * wafermap_scale)
        ax.set_aspect("equal", adjustable="box")
        ax.grid(True, alpha=0.12)
        ax.tick_params(labelsize=9 * wafermap_scale)

    # Hide any unused axes.
    for ax in axes[len(wafers) :]:
        ax.axis("off")

    if mappable is not None:
        # Dedicated axis for colorbar (no overlap with subplot area).
        cax = fig.add_axes([0.87, 0.18, 0.03, 0.62])
        cbar = fig.colorbar(mappable, cax=cax)
        cbar.ax.tick_params(labelsize=8 * wafermap_scale)
        cbar.set_label("Value", fontsize=9 * wafermap_scale)

    fig.suptitle(title, fontsize=10 * wafermap_scale)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, format="png")
    plt.close(fig)


def _parse_filename_wafer_signature(file_name: str) -> str | None:
    m = re.search(r"\b(S11P|S21P|S31P|B11P|B21P)\b", file_name, flags=re.IGNORECASE)
    if not m:
        return None
    return m.group(1).upper()


def _parse_insertion_temperature_label(file_name: str) -> str:
    """Best-effort extraction of insertion temperature from file name."""
    upper = file_name.upper()
    if "S11P" in upper or "B11P" in upper or "HT" in upper or "Q3" in upper:
        return "Hot (135°C)"
    if "S21P" in upper or "Q2" in upper:
        return "Cold (-40°C)"
    if "S31P" in upper or "B21P" in upper or "RT" in upper or "Q1" in upper:
        return "Ambient (25°C)"
    return "Unknown"


def _fmt_num(x: float) -> str:
    return f"{float(x):.6g}"

def _normalize_wafer_ids(values):
    """Extract numeric wafer identifiers from mixed WAFER labels."""
    import pandas as pd

    series = pd.Series(values, copy=False).astype("string").str.strip()
    series = series.mask(series.eq("") | series.str.lower().eq("nan"))
    extracted = series.str.extract(r"(\d+)", expand=False)
    numeric = pd.to_numeric(extracted, errors="coerce")
    normalized = numeric.astype("Int64").astype("string")
    return normalized.mask(normalized.isna())

def _fmt_1dp(x: float) -> str:
    return f"{float(x):.1f}"


def _build_plot_title(
    test_name: str,
    test_col: str,
    temp_label: str,
    cpk: float | None,
    mean_v: float,
    median_v: float,
) -> str:
    first = f"{test_name} ({test_col}) | {temp_label}"
    cpk_txt = _fmt_num(float(cpk)) if cpk is not None and math.isfinite(float(cpk)) else "N/A"
    return first + "\n" + f"Cpk={cpk_txt}; mean={_fmt_num(mean_v)}; median={_fmt_num(median_v)}"


def _apply_module_group_row_colors(ws, *, module_col_header: str = "Module", first_data_row: int = 2) -> None:
    """Apply alternating light fills per contiguous module block (for filtering/readability)."""
    try:
        from openpyxl.styles import PatternFill
    except Exception:
        return

    if ws.max_row < first_data_row or ws.max_column < 1:
        return

    module_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=col_idx).value).strip() == module_col_header:
            module_col_idx = col_idx
            break
    if module_col_idx is None:
        return

    palette = [
        "D9E1F2",  # light blue
        "E2EFDA",  # light green
        "FFF2CC",  # light yellow
        "FCE4D6",  # light orange
        "EDEDED",  # light gray
    ]

    current_module = None
    color_idx = -1
    fill = None
    for row_idx in range(first_data_row, ws.max_row + 1):
        module_val = ws.cell(row=row_idx, column=module_col_idx).value
        module_str = "" if module_val is None else str(module_val)
        if module_str != current_module:
            current_module = module_str
            color_idx = (color_idx + 1) % len(palette)
            fill = PatternFill(patternType="solid", fgColor=palette[color_idx])

        if fill is None:
            continue
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill


def _build_comment(
    *,
    series,
    meta_cols,
    outlier_mad_multiplier: float,
    low_limit: float | None,
    high_limit: float | None,
    wafer_sig: str | None,
) -> str:
    import numpy as np
    import pandas as pd

    vals = pd.to_numeric(series, errors="coerce")
    finite = vals.dropna().to_numpy(dtype=float)
    finite = finite[np.isfinite(finite)]
    if finite.size == 0:
        return "No numeric data"

    med = float(np.median(finite))
    mad = _mad(finite)
    sig = _robust_sigma(finite)

    parts: list[str] = []

    if mad > 0:
        outliers = np.abs(finite - med) > (outlier_mad_multiplier * mad)
        n_out = int(outliers.sum())
        if n_out > 0:
            parts.append(f"Outliers: {n_out}/{finite.size} (>{outlier_mad_multiplier:g}×MAD)")

    # Instability / large spread heuristic.
    if sig > 0 and abs(med) > 0:
        robust_cv = sig / abs(med)
        if robust_cv >= 0.05:
            parts.append(f"Large spread (robust CV={robust_cv:.2%})")

    # Multi-modality heuristic.
    peaks = _count_hist_peaks(finite)
    if peaks >= 2:
        parts.append(f"Possible multi-modal distribution (peaks≈{peaks})")

    # Site effect heuristic.
    if "SITE_NUM" in meta_cols.columns:
        df_site = pd.DataFrame({"v": vals, "SITE_NUM": pd.to_numeric(meta_cols["SITE_NUM"], errors="coerce")})
        df_site = df_site.dropna(subset=["v", "SITE_NUM"])
        if not df_site.empty and df_site["SITE_NUM"].nunique() >= 2:
            med_by_site = df_site.groupby("SITE_NUM")["v"].median().sort_index()
            rng = float(med_by_site.max() - med_by_site.min())
            if sig > 0 and rng / sig >= 3.0:
                worst = med_by_site.idxmax() if abs(med_by_site.max() - med) >= abs(med_by_site.min() - med) else med_by_site.idxmin()
                parts.append(f"Site effect (site medians range={rng:.4g}, worst site={int(worst)})")

    # Wafer signature heuristics.
    if wafer_sig:
        parts.append(f"File wafer signature: {wafer_sig}")

    if "WAFER" in meta_cols.columns:
        df_w = pd.DataFrame({"v": vals, "WAFER": _normalize_wafer_ids(meta_cols["WAFER"])})
        df_w = df_w.dropna(subset=["v"])  # keep WAFER even if blank
        wafer_series = df_w["WAFER"]
        wafers = wafer_series.dropna().unique()
        if wafers.size >= 2:
            med_by_wafer = df_w.dropna(subset=["WAFER"]).groupby("WAFER")["v"].median()
            rng = float(med_by_wafer.max() - med_by_wafer.min())
            if sig > 0 and rng / sig >= 3.0:
                parts.append("Wafer signature suspected (median shifts across wafers)")

    # Coordinate signature (very lightweight).
    for axis in ("X", "Y"):
        if axis in meta_cols.columns:
            coord = pd.to_numeric(meta_cols[axis], errors="coerce")
            df_xy = pd.DataFrame({"v": vals, axis: coord}).dropna()
            if df_xy.shape[0] >= 50:
                rho = _safe_spearman_correlation(df_xy["v"], df_xy[axis])
                if rho is not None and np.isfinite(rho) and abs(float(rho)) >= 0.30:
                    parts.append(f"Coordinate signature: spearman(v,{axis})={float(rho):+.2f}")

    # Fail clustering note.
    if low_limit is not None or high_limit is not None:
        low = -np.inf if low_limit is None else low_limit
        high = np.inf if high_limit is None else high_limit
        fails = (finite < low) | (finite > high)
        n_fail = int(fails.sum())
        if n_fail > 0:
            parts.append(f"Spec fails: {n_fail}/{finite.size}")

    return "; ".join(parts) if parts else "OK"


def _read_unit_data(
    file_path: Path,
    *,
    data_start_line_index: int,
    usecols: list[str],
    encoding: str = DEFAULT_ENCODING,
) -> Any:
    import pandas as pd

    # Keep header (line 0), skip meta rows 1..data_start_line_index-1.
    def _skip(i: int) -> bool:
        return 0 < i < data_start_line_index

    df = pd.read_csv(
        file_path,
        sep=DELIMITER,
        encoding=encoding,
        low_memory=False,
        usecols=usecols,
        skiprows=_skip,
    )
    df.columns = [c.strip() for c in df.columns]
    return df


def _limits_from_meta(meta: FlatFileMeta, test_col: str) -> tuple[float | None, float | None, str | None]:
    low_raw = meta.meta_rows.get("Low", {}).get(test_col)
    high_raw = meta.meta_rows.get("High", {}).get(test_col)
    unit_raw = meta.meta_rows.get("Unit", {}).get(test_col)
    return _to_float(low_raw), _to_float(high_raw), (unit_raw.strip() if unit_raw else None)


def _yield_cpk_from_meta(meta: FlatFileMeta, test_col: str) -> tuple[float | None, float | None]:
    y_raw = meta.meta_rows.get("Yield", {}).get(test_col)
    c_raw = meta.meta_rows.get("Cpk", {}).get(test_col)
    return _to_float(y_raw), _to_float(c_raw)


def _test_name_from_meta(meta: FlatFileMeta, test_col: str) -> str:
    return (meta.meta_rows.get("Test Name", {}).get(test_col) or "").strip().strip('"')


def _status_for_test(
    *,
    yield_pct: float | None,
    cpk: float | None,
    yield_threshold: float,
    cpk_low: float,
    cpk_high: float,
) -> str | None:
    if yield_pct is not None and yield_pct < yield_threshold:
        return "FAILS"
    if cpk is not None and cpk < cpk_low:
        return f"Cpk<{cpk_low:g}"
    if cpk is not None and cpk > cpk_high:
        return f"Cpk>{cpk_high:g}"
    return None


def _proposed_sigma_limits(values) -> tuple[float | None, float | None, float | None, float | None]:
    import numpy as np
    import pandas as pd

    v = pd.to_numeric(values, errors="coerce").dropna().to_numpy(dtype=float)
    v = v[np.isfinite(v)]
    if v.size == 0:
        return None, None, None, None
    med = float(np.median(v))
    sig = _robust_sigma(v)
    if not np.isfinite(sig) or sig <= 0:
        return med, med, med, med

    def _floor_1dp(x: float) -> float:
        return math.floor(x * 10.0) / 10.0

    def _ceil_1dp(x: float) -> float:
        return math.ceil(x * 10.0) / 10.0

    l6, u6 = med - 6.0 * sig, med + 6.0 * sig
    l12, u12 = med - 12.0 * sig, med + 12.0 * sig

    # Quantize to 1 decimal: conservative rounding (floor lows, ceil uppers).
    l6 = _floor_1dp(float(l6))
    u6 = _ceil_1dp(float(u6))
    l12 = _floor_1dp(float(l12))
    u12 = _ceil_1dp(float(u12))

    return l6, u6, l12, u12


def generate_yield_cpk_report(
    *,
    input_folder: Path,
    output_folder: Path,
    modules: list[str],
    outlier_mad_multiplier: float,
    yield_threshold: float,
    cpk_low: float,
    cpk_high: float,
    max_files: int | None,
    single_file: str | None,
    encoding: str = DEFAULT_ENCODING,
) -> Path:
    import numpy as np
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Font, PatternFill

    output_folder.mkdir(parents=True, exist_ok=True)
    plots_root = output_folder / "cdf_plots"
    plots_root.mkdir(parents=True, exist_ok=True)

    if single_file:
        csv_paths = [input_folder / single_file]
    else:
        csv_paths = sorted([p for p in input_folder.glob("*.csv") if p.is_file()])
    if max_files is not None:
        csv_paths = csv_paths[:max_files]
    if not csv_paths:
        raise SystemExit(f"No .csv files found in: {input_folder}")

    modules_upper = {m.strip().upper() for m in modules if m.strip()}
    if not modules_upper:
        raise SystemExit("No modules provided. Example: --modules DPLL,TXPA,TXLO")

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    plot_image_targets_by_sheet: dict[str, list[str]] = {}

    for file_path in csv_paths:
        print(f"Processing: {file_path.name}")
        meta = scan_flat_file_meta(file_path, encoding=encoding)
        wafer_sig = _parse_filename_wafer_signature(file_path.name)
        temp_label = _parse_insertion_temperature_label(file_path.name)

        # Determine tests of interest by module prefix from Test Name row.
        interest_cols: list[str] = []
        interest_names: dict[str, str] = {}
        interest_modules: dict[str, str] = {}
        for test_col in meta.numeric_test_cols:
            tn = _test_name_from_meta(meta, test_col)
            mod = _module_from_test_name(tn)
            if mod in modules_upper:
                interest_cols.append(test_col)
                interest_names[test_col] = tn
                interest_modules[test_col] = mod

        if not interest_cols:
            print(f"  - No tests matched modules {sorted(modules_upper)}; skipping")
            continue

        # Identify affected tests based on Yield/Cpk thresholds.
        affected: list[str] = []
        status_by_col: dict[str, str] = {}
        yield_by_col: dict[str, float | None] = {}
        cpk_by_col: dict[str, float | None] = {}
        for test_col in interest_cols:
            y, c = _yield_cpk_from_meta(meta, test_col)
            yield_by_col[test_col] = y
            cpk_by_col[test_col] = c
            status = _status_for_test(
                yield_pct=y,
                cpk=c,
                yield_threshold=yield_threshold,
                cpk_low=cpk_low,
                cpk_high=cpk_high,
            )
            if status:
                affected.append(test_col)
                status_by_col[test_col] = status

        # Group affected tests by module for readability/formatting.
        affected.sort(key=lambda c: (interest_modules.get(c, ""), int(c)))

        if not affected:
            print("  - No affected tests (per thresholds); skipping sheet")
            continue

        # Load only needed columns from unit data.
        wanted_meta_cols = [
            c
            for c in ("SITE_NUM", "WAFER", "X", "Y", "LOT", "SUBLOT", "CHIP_ID", "PF", "FIRST_FAIL_TEST")
            if c in meta.header
        ]
        usecols = wanted_meta_cols + affected
        df_units = _read_unit_data(
            file_path,
            data_start_line_index=meta.data_start_line_index,
            usecols=usecols,
            encoding=encoding,
        )

        # Ensure meta col dataframe aligns to series indices for comment generation.
        meta_cols_df = df_units[wanted_meta_cols].copy() if wanted_meta_cols else pd.DataFrame(index=df_units.index)

        # Create sheets.
        sheet_name = _unique_sheet_name(file_path.stem, wb.sheetnames)
        ws = wb.create_sheet(sheet_name)
        plot_sheet_name = _unique_sheet_name((file_path.stem[:25] + "_PLOTS"), wb.sheetnames)
        ws_plots = wb.create_sheet(plot_sheet_name)

        # Tab colors: keep data vs plots tabs distinct.
        ws.sheet_properties.tabColor = "4F81BD"  # blue
        ws_plots.sheet_properties.tabColor = "C0504D"  # red

        headers = [
            "Module",
            "Test Nr",
            "Test Name",
            "Unit",
            "Yield (%)",
            "Cpk",
            "Status",
            "N",
            "Fail Chips",
            "Outliers",
            "Comment",
            "CDF Plot",
            "Original LTL",
            "Original UTL",
            "LTL 6s",
            "UTL 6s",
            "LTL 12s",
            "UTL 12s",
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        status_header_fill = PatternFill(patternType="solid", fgColor="FFF2CC")
        ws.cell(row=1, column=headers.index("Status") + 1).fill = status_header_fill

        ws_plots.append(["Test", "CDF (fails highlighted)", "Wafer map (fails highlighted)"])
        ws_plots["A1"].font = Font(bold=True)
        ws_plots["B1"].font = Font(bold=True)
        ws_plots["C1"].font = Font(bold=True)

        plot_anchor_row = 3
        out_rows = 0
        for test_col in affected:
            test_name = interest_names.get(test_col, "")
            module = interest_modules.get(test_col, "")
            low, high, unit = _limits_from_meta(meta, test_col)
            y = yield_by_col.get(test_col)
            c = cpk_by_col.get(test_col)
            status = status_by_col[test_col]

            series = df_units[test_col]
            numeric = pd.to_numeric(series, errors="coerce")
            finite = numeric.dropna().to_numpy(dtype=float)
            finite = finite[np.isfinite(finite)]
            n = int(finite.size)
            if n == 0:
                continue

            low_eval = -np.inf if low is None else float(low)
            high_eval = np.inf if high is None else float(high)
            if low is None and high is None:
                n_fail = 0
            else:
                n_fail = int(((finite < low_eval) | (finite > high_eval)).sum())

            # Outliers
            med = float(np.median(finite))
            mad = _mad(finite)
            n_out = 0
            if mad > 0:
                n_out = int((np.abs(finite - med) > (outlier_mad_multiplier * mad)).sum())

            l6, u6, l12, u12 = _proposed_sigma_limits(numeric)
            comment = _build_comment(
                series=numeric,
                meta_cols=meta_cols_df,
                outlier_mad_multiplier=outlier_mad_multiplier,
                low_limit=low,
                high_limit=high,
                wafer_sig=wafer_sig,
            )

            # Plot file and embed in plot sheet.
            safe_test = re.sub(r"[^A-Za-z0-9._-]+", "_", test_name)[:80] or test_col
            plot_path = plots_root / file_path.stem / f"{test_col}_{safe_test}.png"
            include_sigma_limits = status.startswith("Cpk<") or status.startswith("Cpk>")
            title = _build_plot_title(
                test_name=test_name,
                test_col=test_col,
                temp_label=temp_label,
                cpk=c,
                mean_v=float(np.mean(finite)),
                median_v=float(np.median(finite)),
            )
            _cdf_plot_png(
                numeric,
                title=title,
                out_path=plot_path,
                low_limit=low,
                high_limit=high,
                proposed_l6=(l6 if include_sigma_limits else None),
                proposed_u6=(u6 if include_sigma_limits else None),
                proposed_l12=(l12 if include_sigma_limits else None),
                proposed_u12=(u12 if include_sigma_limits else None),
            )

            wafer_map_path = plots_root / file_path.stem / f"{test_col}_{safe_test}_wafermap.png"
            _wafer_map_png(
                numeric,
                meta_cols=meta_cols_df,
                title=f"{test_name} ({test_col}) | {temp_label}",
                out_path=wafer_map_path,
                low_limit=low,
                high_limit=high,
            )
            wafer_map_html_path = plots_root / file_path.stem / f"{test_col}_{safe_test}_wafermap.html"
            _wafer_map_html(
                numeric,
                meta_cols=meta_cols_df,
                title=f"{test_name} ({test_col}) | {temp_label}",
                out_path=wafer_map_html_path,
                low_limit=low,
                high_limit=high,
            )

            # Write to plots sheet.
            ws_plots[f"A{plot_anchor_row}"] = f"{test_col} {test_name}"
            ws_plots[f"A{plot_anchor_row}"].font = Font(bold=True)

            # Zoomable links (open the PNG in an external viewer).
            cdf_link = ws_plots[f"B{plot_anchor_row}"]
            cdf_link.value = "Open CDF PNG"
            if plot_path.exists():
                plot_uri = plot_path.resolve().as_uri()
                cdf_link.hyperlink = plot_uri
                cdf_link.font = Font(color="0000EE", underline="single")
                plot_image_targets_by_sheet.setdefault(ws_plots.title, []).append(plot_uri)

            wafer_link = ws_plots[f"C{plot_anchor_row}"]
            wafer_link.value = "Open interactive wafer map"
            wafer_click_path = wafer_map_html_path if wafer_map_html_path.exists() else wafer_map_path
            if wafer_click_path.exists():
                wafer_uri = wafer_click_path.resolve().as_uri()
                wafer_link.hyperlink = wafer_uri
                wafer_link.font = Font(color="0000EE", underline="single")
            else:
                wafer_uri = None
            cdf_anchor_cell = f"B{plot_anchor_row + 1}"
            wafer_anchor_cell = f"K{plot_anchor_row + 1}"
            if plot_path.exists():
                img = XLImage(str(plot_path))
                img.width = 520
                img.height = 360
                ws_plots.add_image(img, cdf_anchor_cell)

            if wafer_map_path.exists():
                wimg = XLImage(str(wafer_map_path))
                wimg.width = 520
                wimg.height = 360
                ws_plots.add_image(wimg, wafer_anchor_cell)
                if wafer_uri is not None:
                    plot_image_targets_by_sheet.setdefault(ws_plots.title, []).append(wafer_uri)
            # Reserve some rows for the image.
            plot_link_target = f"#{_excel_internal_sheet_ref(ws_plots.title)}!{cdf_anchor_cell}"
            plot_anchor_row += 30

            # Write row to data sheet.
            row = [
                module,
                int(test_col),
                test_name,
                unit,
                y,
                c,
                status,
                n,
                n_fail,
                n_out,
                comment,
                "View",
                low,
                high,
                l6,
                u6,
                l12,
                u12,
            ]
            ws.append(row)
            out_rows += 1

            row_idx = 1 + out_rows
            for col_name in ("LTL 6s", "UTL 6s", "LTL 12s", "UTL 12s"):
                col_idx = headers.index(col_name) + 1
                ws.cell(row=row_idx, column=col_idx).number_format = "0.0"

            link_cell = ws.cell(row=1 + out_rows, column=headers.index("CDF Plot") + 1)
            link_cell.hyperlink = plot_link_target
            link_cell.font = Font(color="0000EE", underline="single")

        if out_rows == 0:
            # Avoid leaving empty sheets.
            wb.remove(ws)
            wb.remove(ws_plots)
            continue

        fail_chips_col_idx = headers.index("Fail Chips") + 1
        fail_chips_col_letter = _excel_col_letter(fail_chips_col_idx)
        fail_chips_range = f"{fail_chips_col_letter}2:{fail_chips_col_letter}{1 + out_rows}"
        ws.conditional_formatting.add(
            fail_chips_range,
            ColorScaleRule(
                start_type="min",
                start_color="63BE7B",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFEB84",
                end_type="max",
                end_color="F8696B",
            ),
        )

        # Apply module block coloring on the data sheet.
        _apply_module_group_row_colors(ws)

        # Auto-fit columns.
        _autofit_openpyxl_columns(ws)
        _autofit_openpyxl_columns(ws_plots)

        # Freeze header row.
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{_excel_col_letter(ws.max_column)}{ws.max_row}"

    from datetime import datetime

    out_xlsx = output_folder / "Yield_Cpk_Report.xlsx"
    link_issues = _self_check_workbook_internal_hyperlinks(wb)
    if link_issues:
        preview = "\n".join(f"- {msg}" for msg in link_issues[:15])
        more = "" if len(link_issues) <= 15 else f"\n- ... and {len(link_issues) - 15} more"
        raise ValueError("Internal hyperlink self-check failed:\n" + preview + more)

    try:
        wb.save(out_xlsx)
        _enable_clickable_plot_images(out_xlsx, plot_image_targets_by_sheet)
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = output_folder / f"Yield_Cpk_Report_{ts}.xlsx"
        wb.save(alt)
        _enable_clickable_plot_images(alt, plot_image_targets_by_sheet)
        print(f"Could not overwrite (file open?): {out_xlsx}")
        print(f"Saved instead: {alt}")
        return alt

    print(f"Saved: {out_xlsx}")
    return out_xlsx


def generate_correlation_workbook(
    *,
    input_folder: Path,
    output_folder: Path,
    modules: list[str],
    max_files: int | None,
    single_file: str | None,
    methods: list[Literal["pearson", "spearman"]],
    pearson_abs_min_for_report: float,
    encoding: str = DEFAULT_ENCODING,
) -> Path:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font

    output_folder.mkdir(parents=True, exist_ok=True)
    if single_file:
        csv_paths = [input_folder / single_file]
    else:
        csv_paths = sorted([p for p in input_folder.glob("*.csv") if p.is_file()])
    if max_files is not None:
        csv_paths = csv_paths[:max_files]
    if not csv_paths:
        raise SystemExit(f"No .csv files found in: {input_folder}")

    modules_upper = {m.strip().upper() for m in modules if m.strip()}
    if not modules_upper:
        raise SystemExit("No modules provided.")

    wb = Workbook()
    wb.remove(wb.active)
    excel_max_rows = 1_048_576
    excel_max_data_rows = excel_max_rows - 1  # account for header row

    for file_path in csv_paths:
        print(f"Correlation: {file_path.name}")
        meta = scan_flat_file_meta(file_path, encoding=encoding)

        test_name_by_col = {c: _test_name_from_meta(meta, c) for c in meta.numeric_test_cols}
        module_by_col = {c: _module_from_test_name(test_name_by_col[c]) for c in meta.numeric_test_cols}
        module_cols = [c for c in meta.numeric_test_cols if module_by_col[c] in modules_upper]
        if not module_cols:
            print("  - No module tests found; skipping")
            continue

        # Load ALL numeric tests for correlation (optional heavy step).
        df = _read_unit_data(
            file_path,
            data_start_line_index=meta.data_start_line_index,
            usecols=meta.numeric_test_cols,
            encoding=encoding,
        )
        df = df.apply(pd.to_numeric, errors="coerce")

        sheet = wb.create_sheet(_unique_sheet_name(file_path.stem, wb.sheetnames))
        headers = [
            "Module",
            "Test Nr",
            "Test Name",
            "Other Test Nr",
            "Other Test Name",
        ] + [m.title() for m in methods]
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        # Precompute ranks for spearman if needed.
        df_rank = None
        if "spearman" in methods:
            df_rank = df.rank(axis=0, method="average", na_option="keep")

        out_rows = 0
        truncated = False
        for test_col in module_cols:
            if out_rows >= excel_max_data_rows:
                truncated = True
                break
            s = df[test_col]
            if s.dropna().nunique() < 2:
                continue
            pearson = _safe_corr_against_all(df, test_col) if "pearson" in methods else None
            spearman = _safe_corr_against_all(df_rank, test_col) if ("spearman" in methods and df_rank is not None) else None

            for other_col in meta.numeric_test_cols:
                if out_rows >= excel_max_data_rows:
                    truncated = True
                    break
                if other_col == test_col:
                    continue

                pearson_value = None
                if pearson is not None:
                    pearson_value = _to_float(pearson.get(other_col))
                if pearson_value is None or abs(float(pearson_value)) <= float(pearson_abs_min_for_report):
                    continue

                row = [
                    module_by_col[test_col],
                    int(test_col),
                    test_name_by_col[test_col],
                    int(other_col),
                    test_name_by_col[other_col],
                ]
                if pearson is not None:
                    row.append(pearson_value)
                if spearman is not None:
                    row.append(_to_float(spearman.get(other_col)))
                sheet.append(row)
                out_rows += 1

        if truncated:
            print(f"  - Correlation rows truncated at Excel limit ({excel_max_data_rows} data rows)")

        if out_rows == 0:
            wb.remove(sheet)
            continue

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{_excel_col_letter(sheet.max_column)}{sheet.max_row}"
        _autofit_openpyxl_columns(sheet)

    from datetime import datetime

    out_xlsx = output_folder / "Correlation_Report.xlsx"
    try:
        wb.save(out_xlsx)
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = output_folder / f"Correlation_Report_{ts}.xlsx"
        wb.save(alt)
        print(f"Could not overwrite (file open?): {out_xlsx}")
        print(f"Saved instead: {alt}")
        return alt

    print(f"Saved: {out_xlsx}")
    return out_xlsx


def _safe_pair_correlation(a, b) -> float | None:
    """Compute Pearson correlation for two aligned series safely.

    Returns None when there are not enough paired values or one side is constant.
    """
    import numpy as np
    import pandas as pd

    aa = pd.to_numeric(a, errors="coerce")
    bb = pd.to_numeric(b, errors="coerce")
    valid = aa.notna() & bb.notna()
    if int(valid.sum()) < 2:
        return None

    av = aa[valid].to_numpy(dtype=float)
    bv = bb[valid].to_numpy(dtype=float)
    if av.size < 2 or bv.size < 2:
        return None

    std_a = float(np.std(av, ddof=1))
    std_b = float(np.std(bv, ddof=1))
    if not np.isfinite(std_a) or not np.isfinite(std_b) or std_a == 0.0 or std_b == 0.0:
        return None

    corr = float(np.corrcoef(av, bv)[0, 1])
    if not np.isfinite(corr):
        return None
    return corr


def _safe_spearman_correlation(a, b) -> float | None:
    """Compute Spearman correlation without requiring SciPy."""
    import pandas as pd

    aa = pd.to_numeric(a, errors="coerce")
    bb = pd.to_numeric(b, errors="coerce")
    valid = aa.notna() & bb.notna()
    if int(valid.sum()) < 2:
        return None

    aa_rank = aa[valid].rank(method="average")
    bb_rank = bb[valid].rank(method="average")
    return _safe_pair_correlation(aa_rank, bb_rank)


def _safe_corr_against_all(df, target_col: str) -> dict[str, float | None]:
    """Safe correlation of target column against all columns in df."""
    out: dict[str, float | None] = {}
    target = df[target_col]
    for col in df.columns:
        if col == target_col:
            continue
        out[col] = _safe_pair_correlation(target, df[col])
    return out


def run() -> int:
    """Run analysis using runtime configuration supplied by the YAML wrapper."""
    config = _require_runtime_configuration()
    input_folder = config["input_folder"]
    output_folder = config["output_folder"]

    output_folder.mkdir(parents=True, exist_ok=True)

    generate_yield_cpk_report(
        input_folder=input_folder,
        output_folder=output_folder,
        modules=config["modules"],
        outlier_mad_multiplier=config["outlier_mad_multiplier"],
        yield_threshold=config["yield_threshold"],
        cpk_low=config["cpk_low"],
        cpk_high=config["cpk_high"],
        max_files=config["max_files"],
        single_file=config["single_file"],
        encoding=config["encoding"],
    )

    if config["generate_correlation_report"]:
        generate_correlation_workbook(
            input_folder=input_folder,
            output_folder=output_folder,
            modules=config["modules"],
            max_files=config["max_files"],
            single_file=config["single_file"],
            methods=config["correlation_methods"],
            pearson_abs_min_for_report=config["pearson_abs_min_for_report"],
            encoding=config["encoding"],
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        "Direct execution of Tests_Data_Analysis.py is disabled. "
        "Use run_tests_data_analysis.py --config <yaml> or run_tests_data_analysis.ps1."
    )
