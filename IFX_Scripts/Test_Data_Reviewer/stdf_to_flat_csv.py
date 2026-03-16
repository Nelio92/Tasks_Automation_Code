from __future__ import annotations

import argparse
import bz2
import concurrent.futures
import contextlib
import csv
import gzip
import io
import json
import lzma
import math
import multiprocessing
import os
import re
import sys
import tarfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence


DELIMITER = ";"
DEFAULT_PATTERNS = (
    "*.stdf",
    "*.std",
    "*.eff",
    "*.stdf.gz",
    "*.stdf.xz",
    "*.stdf.bz2",
    "*.std.gz",
    "*.std.xz",
    "*.std.bz2",
    "*.stdf.tar.gz",
    "*.std.tar.gz",
)
META_COLUMNS = ["UNIT_ID", "SITE_NUM", "WAFER", "X", "Y", "LOT", "SUBLOT", "CHIP_ID", "PF", "FIRST_FAIL_TEST"]
SUMMARY_ROWS = ("Test Name", "Low", "High", "Unit", "Cpk", "Yield", "Mean", "Stddev")
COMPRESSED_SUFFIXES = {".gz", ".xz", ".bz2"}
ARCHIVE_SUFFIXES = {".tar"}
SUPPORTED_INPUT_SUFFIXES = {".stdf", ".std", ".eff"}


@dataclass(slots=True)
class TestColumn:
    column_id: str
    test_num: int
    test_name: str
    unit: str = ""
    low: float | None = None
    high: float | None = None


@dataclass(slots=True)
class PartRow:
    unit_id: int
    site_num: int | None = None
    wafer: str | None = None
    x: int | None = None
    y: int | None = None
    lot: str | None = None
    sublot: str | None = None
    chip_id: str | None = None
    pf: str | None = None
    first_fail_test: str | None = None
    measurements: dict[str, float | int | str | None] = field(default_factory=dict)


@dataclass(slots=True)
class ColumnAccumulator:
    count: int = 0
    passed: int = 0
    value_sum: float = 0.0
    value_sum_sq: float = 0.0

    def add(self, value: float, *, low: float | None, high: float | None) -> None:
        self.count += 1
        self.value_sum += float(value)
        self.value_sum_sq += float(value) * float(value)
        if not _value_fails_limits(float(value), low, high):
            self.passed += 1


@dataclass(slots=True)
class ConversionSummary:
    converted_files: int
    converted_parts: int
    converted_tests: int
    output_files: list[Path]
    dtr_files: list[Path] = field(default_factory=list)
    consistency_files: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    file_results: list["ConversionFileResult"] = field(default_factory=list)


@dataclass(slots=True)
class ConversionFileResult:
    source_file: Path
    csv_file: Path
    dtr_file: Path | None
    consistency_file: Path
    generated_rows: int
    numeric_tests: int
    record_counts: dict[str, int]
    malformed_record_count: int
    malformed_record_types: dict[str, int]
    dtr_record_count: int
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class _ConversionState:
    ordered_columns: list[TestColumn] = field(default_factory=list)
    columns_by_key: dict[tuple[int, str], TestColumn] = field(default_factory=dict)
    used_column_ids: set[str] = field(default_factory=set)
    parts: list[PartRow] = field(default_factory=list)
    active_parts_by_site: dict[int | None, PartRow] = field(default_factory=dict)
    current_wafer: str | None = None
    current_lot: str | None = None
    current_sublot: str | None = None
    next_unit_id: int = 1
    record_counts: Counter[str] = field(default_factory=Counter)
    malformed_record_types: Counter[str] = field(default_factory=Counter)
    dtr_messages: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    accumulators_by_column_id: dict[str, ColumnAccumulator] = field(default_factory=dict)


def _clean_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        result = float(value)
        if math.isnan(result) or math.isinf(result):
            return None
        return result
    text = str(value).strip()
    if not text or text.lower() in {"nan", "na", "none", "inf", "-inf", "+inf"}:
        return None
    try:
        result = float(text)
    except ValueError:
        try:
            result = float(text.replace(",", "."))
        except ValueError:
            return None
    if math.isnan(result) or math.isinf(result):
        return None
    return result


def _to_int(value: Any) -> int | None:
    number = _to_float(value)
    if number is None:
        return None
    try:
        return int(number)
    except (TypeError, ValueError, OverflowError):
        return None


def _record_to_dict(record: Any) -> dict[str, Any]:
    if isinstance(record, dict):
        return dict(record)
    if hasattr(record, "to_dict"):
        raw = record.to_dict()
        if isinstance(raw, dict):
            return dict(raw)
    fields: dict[str, Any] = {}
    if hasattr(record, "get_fields"):
        for field_name in record.get_fields():
            try:
                fields[str(field_name)] = record.get_value(field_name)
            except Exception:
                continue
    return fields


def _record_id(record: Any, record_dict: dict[str, Any]) -> str:
    record_name = getattr(record, "id", None)
    if record_name:
        return str(record_name)
    for key in ("REC_ID", "record_id", "id"):
        if key in record_dict and record_dict[key]:
            return str(record_dict[key])
    return ""


def _field(record_dict: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in record_dict:
            return record_dict[name]
    return None


def _stdf_flag_indicates_fail(*flags: Any) -> bool:
    for flag in flags:
        if flag in (None, "", (), [], {}):
            continue
        if isinstance(flag, (list, tuple)):
            if any(bool(item) for item in flag):
                return True
            continue
        if isinstance(flag, (bytes, bytearray)):
            if any(byte != 0 for byte in flag):
                return True
            continue
        if isinstance(flag, int):
            if flag != 0:
                return True
            continue
        if isinstance(flag, str):
            text = flag.strip().lower()
            if not text or text in {"0", "pass", "p", "false", "[]"}:
                continue
            return True
        if bool(flag):
            return True
    return False


def _value_fails_limits(value: float | None, low: float | None, high: float | None) -> bool:
    if value is None:
        return False
    if low is not None and value < low:
        return True
    if high is not None and value > high:
        return True
    return False


def csv_name_for_source(source_name: str) -> str:
    name = Path(source_name).name
    suffixes = Path(name).suffixes
    while suffixes and suffixes[-1].lower() in (COMPRESSED_SUFFIXES | ARCHIVE_SUFFIXES | SUPPORTED_INPUT_SUFFIXES):
        name = Path(name).with_suffix("").name
        suffixes = Path(name).suffixes
    stem = Path(name).stem
    return f"{stem}.csv"


def dtr_name_for_source(source_name: str) -> str:
    return f"{Path(csv_name_for_source(source_name)).stem}_dtr_records.csv"


def consistency_name_for_source(source_name: str) -> str:
    return f"{Path(csv_name_for_source(source_name)).stem}_conversion_consistency.json"


def _source_kind_for_name(source_name: str) -> str:
    suffixes = [suffix.lower() for suffix in Path(source_name).suffixes]
    while suffixes and suffixes[-1] in (COMPRESSED_SUFFIXES | ARCHIVE_SUFFIXES):
        suffixes.pop()
    if suffixes and suffixes[-1] == ".eff":
        return "eff"
    return "stdf"


def _is_tar_archive_path(path: Path) -> bool:
    suffixes = {suffix.lower() for suffix in path.suffixes}
    if ".tar" in suffixes:
        return True
    try:
        return tarfile.is_tarfile(path)
    except (OSError, tarfile.TarError):
        return False


def _select_archive_member_name(path: Path) -> str:
    with tarfile.open(path, "r:*") as archive:
        regular_members = [member for member in archive.getmembers() if member.isfile()]
        if not regular_members:
            raise ValueError(f"Archive does not contain any regular files: {path}")

        preferred_members = [
            member
            for member in regular_members
            if any(Path(member.name).name.lower().endswith(ext) for ext in SUPPORTED_INPUT_SUFFIXES)
        ]
        if len(preferred_members) == 1:
            return preferred_members[0].name
        if len(preferred_members) > 1:
            preferred_members.sort(key=lambda member: (Path(member.name).suffix.lower() != ".std", len(member.name), member.name.lower()))
            return preferred_members[0].name

        if len(regular_members) == 1:
            return regular_members[0].name

        sample_names = ", ".join(member.name for member in regular_members[:5])
        raise ValueError(f"Could not determine which archive member to convert from {path}. Candidates: {sample_names}")


def _cell_value(row: Sequence[str], index_by_name: dict[str, int], column_name: str) -> str | None:
    idx = index_by_name.get(column_name)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _normalize_eff_row(row: Sequence[str], expected_len: int) -> list[str]:
    values = [str(cell).strip() for cell in row]
    if len(values) < expected_len:
        values.extend([""] * (expected_len - len(values)))
    return values[:expected_len]


def _resolve_eff_first_fail_test(raw_value: str | None, test_names_by_column_id: dict[str, str]) -> str | None:
    text = _clean_text(raw_value)
    if not text:
        return None
    return test_names_by_column_id.get(text, text)


def _build_eff_state(reader: Iterable[Sequence[str]]) -> _ConversionState:
    state = _ConversionState()
    state.record_counts["EFF"] += 1

    header_row: list[str] | None = None
    meta_rows: dict[str, list[str]] = {}
    first_data_row: list[str] | None = None

    for raw_row in reader:
        if not raw_row:
            continue
        key = str(raw_row[0]).strip()
        if header_row is None:
            if key.startswith("<+EFF:"):
                header_row = [str(cell).strip() for cell in raw_row]
            continue

        normalized_row = _normalize_eff_row(raw_row, len(header_row))
        key = normalized_row[0]
        if key.startswith("<"):
            meta_rows[key] = normalized_row
            continue
        first_data_row = normalized_row
        break

    if header_row is None:
        raise ValueError("EFF input is missing the <+EFF:...> header row")

    index_by_name = {name: idx for idx, name in enumerate(header_row) if name}
    numeric_indices = [idx for idx, name in enumerate(header_row) if idx > 0 and str(name).isdigit()]
    if not numeric_indices:
        raise ValueError("EFF input does not contain numeric test columns")

    pname_row = meta_rows.get("<+PName>", [""] * len(header_row))
    unit_row = meta_rows.get("<Unit>", [""] * len(header_row))
    usl_row = meta_rows.get("<USL>", [""] * len(header_row))
    lsl_row = meta_rows.get("<LSL>", [""] * len(header_row))

    columns_by_index: dict[int, TestColumn] = {}
    test_names_by_column_id: dict[str, str] = {}
    for idx in numeric_indices:
        header_name = header_row[idx]
        test_num = _to_int(header_name)
        if test_num is None:
            continue
        column = _get_or_create_test_column(
            columns_by_key=state.columns_by_key,
            ordered_columns=state.ordered_columns,
            used_column_ids=state.used_column_ids,
            test_num=test_num,
            test_name=_clean_text(pname_row[idx]) or header_name,
            unit=_clean_text(unit_row[idx]) or "",
            low=_to_float(lsl_row[idx]),
            high=_to_float(usl_row[idx]),
            variant="EFF",
        )
        columns_by_index[idx] = column
        test_names_by_column_id[header_name] = column.test_name

    def process_data_row(row: Sequence[str]) -> None:
        normalized = _normalize_eff_row(row, len(header_row))
        unit_id = _to_int(_cell_value(normalized, index_by_name, "VNr"))
        if unit_id is None:
            unit_id = state.next_unit_id
        state.next_unit_id = max(state.next_unit_id, unit_id + 1)

        pf_value = _clean_text(_cell_value(normalized, index_by_name, "PF"))
        if pf_value:
            pf_value = pf_value[:1].upper()

        part = PartRow(
            unit_id=unit_id,
            site_num=_to_int(_cell_value(normalized, index_by_name, "SITE_NUM")),
            wafer=_clean_text(_cell_value(normalized, index_by_name, "WAFER")),
            x=_to_int(_cell_value(normalized, index_by_name, "X")),
            y=_to_int(_cell_value(normalized, index_by_name, "Y")),
            lot=_clean_text(_cell_value(normalized, index_by_name, "LOT")),
            sublot=_clean_text(_cell_value(normalized, index_by_name, "SUBLOT")),
            chip_id=_clean_text(_cell_value(normalized, index_by_name, "CHIP_ID")),
            pf=pf_value,
            first_fail_test=_resolve_eff_first_fail_test(
                _cell_value(normalized, index_by_name, "FIRST_FAIL_TEST"),
                test_names_by_column_id,
            ),
        )

        for idx, column in columns_by_index.items():
            value = _to_float(normalized[idx])
            part.measurements[column.column_id] = value
            _update_column_accumulator(state, column=column, value=value)
            if part.first_fail_test is None:
                part.first_fail_test = _detect_failed_test_name(
                    column.test_name,
                    value,
                    column.low,
                    column.high,
                )

        if part.pf is None:
            part.pf = "F" if part.first_fail_test else "P"
        state.parts.append(part)
        state.record_counts["EFF_ROW"] += 1

    if first_data_row is not None:
        process_data_row(first_data_row)
    for raw_row in reader:
        if not raw_row:
            continue
        process_data_row(raw_row)

    return state


def _unique_numeric_column_id(test_num: int, used: set[str]) -> str:
    base = str(int(test_num))
    if base not in used:
        used.add(base)
        return base
    duplicate_index = 1
    while True:
        candidate = f"{base}{duplicate_index:03d}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        duplicate_index += 1


def _get_or_create_test_column(
    *,
    columns_by_key: dict[tuple[int, str], TestColumn],
    ordered_columns: list[TestColumn],
    used_column_ids: set[str],
    test_num: int,
    test_name: str,
    unit: str | None,
    low: float | None,
    high: float | None,
    variant: str,
) -> TestColumn:
    key = (int(test_num), variant)
    existing = columns_by_key.get(key)
    if existing is not None:
        if not existing.unit and unit:
            existing.unit = unit
        if existing.low is None and low is not None:
            existing.low = low
        if existing.high is None and high is not None:
            existing.high = high
        if (not existing.test_name or existing.test_name == str(existing.test_num)) and test_name:
            existing.test_name = test_name
        return existing

    column = TestColumn(
        column_id=_unique_numeric_column_id(test_num, used_column_ids),
        test_num=int(test_num),
        test_name=test_name or str(test_num),
        unit=unit or "",
        low=low,
        high=high,
    )
    columns_by_key[key] = column
    ordered_columns.append(column)
    return column


def _mean(values: Sequence[float]) -> float | None:
    if not values:
        return None
    return float(sum(values) / len(values))


def _sample_stddev(values: Sequence[float]) -> float | None:
    if len(values) < 2:
        return None
    mean = _mean(values)
    if mean is None:
        return None
    variance = sum((value - mean) ** 2 for value in values) / (len(values) - 1)
    return float(math.sqrt(variance))


def _cpk(values: Sequence[float], low: float | None, high: float | None) -> float | None:
    if low is None and high is None:
        return None
    stddev = _sample_stddev(values)
    mean = _mean(values)
    if stddev is None or mean is None or stddev <= 0.0:
        return None
    candidates: list[float] = []
    if high is not None:
        candidates.append((high - mean) / (3.0 * stddev))
    if low is not None:
        candidates.append((mean - low) / (3.0 * stddev))
    finite = [float(value) for value in candidates if math.isfinite(value)]
    return min(finite) if finite else None


def _yield_percent(values: Sequence[float], low: float | None, high: float | None) -> float | None:
    if not values:
        return None
    if low is None and high is None:
        return 100.0
    passed = sum(1 for value in values if not _value_fails_limits(value, low, high))
    return 100.0 * passed / len(values)


def _format_number(value: float | int | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if math.isnan(value) or math.isinf(value):
        return ""
    text = f"{float(value):.12g}"
    return "0" if text == "-0" else text


def _build_summary_row(name: str, ordered_columns: Sequence[TestColumn], parts: Sequence[PartRow]) -> list[str]:
    row = [name] + ["" for _ in META_COLUMNS[1:]]
    for column in ordered_columns:
        numeric_values = [
            float(value)
            for part in parts
            for value in [part.measurements.get(column.column_id)]
            if isinstance(value, (int, float)) and not isinstance(value, bool)
        ]
        if name == "Test Name":
            row.append(column.test_name)
        elif name == "Low":
            row.append(_format_number(column.low))
        elif name == "High":
            row.append(_format_number(column.high))
        elif name == "Unit":
            row.append(column.unit)
        elif name == "Cpk":
            row.append(_format_number(_cpk(numeric_values, column.low, column.high)))
        elif name == "Yield":
            row.append(_format_number(_yield_percent(numeric_values, column.low, column.high)))
        elif name == "Mean":
            row.append(_format_number(_mean(numeric_values)))
        elif name == "Stddev":
            row.append(_format_number(_sample_stddev(numeric_values)))
        else:
            row.append("")
    return row


def _summary_value_from_accumulator(name: str, column: TestColumn, accumulator: ColumnAccumulator | None) -> str:
    if name == "Test Name":
        return column.test_name
    if name == "Low":
        return _format_number(column.low)
    if name == "High":
        return _format_number(column.high)
    if name == "Unit":
        return column.unit
    if accumulator is None or accumulator.count <= 0:
        return ""

    mean = accumulator.value_sum / accumulator.count
    if name == "Mean":
        return _format_number(mean)
    if name == "Yield":
        return _format_number(100.0 * accumulator.passed / accumulator.count)

    stddev: float | None = None
    if accumulator.count >= 2:
        variance = (accumulator.value_sum_sq - ((accumulator.value_sum * accumulator.value_sum) / accumulator.count)) / (accumulator.count - 1)
        variance = max(0.0, float(variance))
        stddev = float(math.sqrt(variance))

    if name == "Stddev":
        return _format_number(stddev)
    if name == "Cpk":
        if stddev is None or stddev <= 0.0 or (column.low is None and column.high is None):
            return ""
        candidates: list[float] = []
        if column.high is not None:
            candidates.append((float(column.high) - mean) / (3.0 * stddev))
        if column.low is not None:
            candidates.append((mean - float(column.low)) / (3.0 * stddev))
        finite = [float(value) for value in candidates if math.isfinite(value)]
        return _format_number(min(finite) if finite else None)
    return ""


def _build_summary_row_from_accumulators(
    name: str,
    ordered_columns: Sequence[TestColumn],
    accumulators_by_column_id: dict[str, ColumnAccumulator],
) -> list[str]:
    row = [name] + ["" for _ in META_COLUMNS[1:]]
    for column in ordered_columns:
        row.append(_summary_value_from_accumulator(name, column, accumulators_by_column_id.get(column.column_id)))
    return row


def _update_column_accumulator(
    state: _ConversionState,
    *,
    column: TestColumn,
    value: float | None,
) -> None:
    if value is None:
        return
    accumulator = state.accumulators_by_column_id.get(column.column_id)
    if accumulator is None:
        accumulator = ColumnAccumulator()
        state.accumulators_by_column_id[column.column_id] = accumulator
    accumulator.add(value, low=column.low, high=column.high)


def _detect_failed_test_name(test_name: str | None, value: float | None, low: float | None, high: float | None, *flags: Any) -> str | None:
    if _value_fails_limits(value, low, high) or _stdf_flag_indicates_fail(*flags):
        return test_name or None
    return None


def _finalize_part(part: PartRow, record_dict: dict[str, Any]) -> PartRow:
    part.site_num = _to_int(_field(record_dict, "SITE_NUM")) if part.site_num is None else part.site_num
    part.x = _to_int(_field(record_dict, "X_COORD", "X")) if part.x is None else part.x
    part.y = _to_int(_field(record_dict, "Y_COORD", "Y")) if part.y is None else part.y
    part.chip_id = _clean_text(_field(record_dict, "PART_ID", "CHIP_ID")) or part.chip_id
    if part.pf is None:
        part.pf = "F" if part.first_fail_test or _stdf_flag_indicates_fail(_field(record_dict, "PART_FLG")) else "P"
    return part


def _create_part_for_site(state: _ConversionState, site_num: int | None) -> PartRow:
    part = PartRow(
        unit_id=state.next_unit_id,
        site_num=site_num,
        wafer=state.current_wafer,
        lot=state.current_lot,
        sublot=state.current_sublot,
    )
    state.next_unit_id += 1
    state.active_parts_by_site[site_num] = part
    return part


def _get_or_create_active_part(state: _ConversionState, site_num: int | None) -> PartRow:
    if site_num is None and len(state.active_parts_by_site) == 1:
        return next(iter(state.active_parts_by_site.values()))
    existing = state.active_parts_by_site.get(site_num)
    if existing is not None:
        return existing
    return _create_part_for_site(state, site_num)


def _process_record(state: _ConversionState, record_name: str, record_dict: dict[str, Any]) -> None:
    record_name = record_name.upper()
    state.record_counts[record_name] += 1

    if record_name == "MIR":
        state.current_lot = _clean_text(_field(record_dict, "LOT_ID")) or state.current_lot
        state.current_sublot = _clean_text(_field(record_dict, "SBLOT_ID", "SUBLOT_ID", "SUBLOT")) or state.current_sublot
        return

    if record_name == "WIR":
        state.current_wafer = _clean_text(_field(record_dict, "WAFER_ID", "WAFER")) or state.current_wafer
        return

    if record_name == "DTR":
        text = _clean_text(_field(record_dict, "TEXT_DAT"))
        if text:
            state.dtr_messages.append(text)
        return

    if record_name == "PIR":
        site_num = _to_int(_field(record_dict, "SITE_NUM"))
        _create_part_for_site(state, site_num)
        return

    if record_name == "PRR":
        site_num = _to_int(_field(record_dict, "SITE_NUM"))
        part = state.active_parts_by_site.pop(site_num, None)
        if part is None:
            part = _create_part_for_site(state, site_num)
            state.active_parts_by_site.pop(site_num, None)
        part = _finalize_part(part, record_dict)
        state.parts.append(part)
        return

    if record_name not in {"PTR", "MPR"}:
        return

    site_num = _to_int(_field(record_dict, "SITE_NUM"))
    current_part = _get_or_create_active_part(state, site_num)

    test_num = _to_int(_field(record_dict, "TEST_NUM"))
    if test_num is None:
        return

    test_name = _clean_text(_field(record_dict, "TEST_TXT", "TEST_NAM", "VECT_NAM")) or str(test_num)
    unit = _clean_text(_field(record_dict, "UNITS", "UNIT")) or ""
    low = _to_float(_field(record_dict, "LO_LIMIT", "LLM", "LOW_LIMIT"))
    high = _to_float(_field(record_dict, "HI_LIMIT", "HLM", "HIGH_LIMIT"))
    test_flag = _field(record_dict, "TEST_FLG")
    param_flag = _field(record_dict, "PARM_FLG")

    if record_name == "PTR":
        result = _to_float(_field(record_dict, "RESULT"))
        column = _get_or_create_test_column(
            columns_by_key=state.columns_by_key,
            ordered_columns=state.ordered_columns,
            used_column_ids=state.used_column_ids,
            test_num=test_num,
            test_name=test_name,
            unit=unit,
            low=low,
            high=high,
            variant="PTR",
        )
        current_part.measurements[column.column_id] = result
        _update_column_accumulator(state, column=column, value=result)
        if current_part.first_fail_test is None:
            current_part.first_fail_test = _detect_failed_test_name(test_name, result, low, high, test_flag, param_flag)
        return

    results_raw = _field(record_dict, "RTN_RSLT", "RESULTS")
    if not isinstance(results_raw, (list, tuple)):
        single_result = _to_float(results_raw)
        results = [single_result] if single_result is not None else []
    else:
        results = [_to_float(value) for value in results_raw]
    if not results:
        return

    pin_labels_raw = _field(record_dict, "RTN_INDX", "RTN_STAT", "PMR_INDX")
    if isinstance(pin_labels_raw, (list, tuple)):
        pin_labels = [str(item) for item in pin_labels_raw]
    else:
        pin_labels = []

    for index, result in enumerate(results, start=1):
        pin_suffix = pin_labels[index - 1] if index - 1 < len(pin_labels) and pin_labels[index - 1] else str(index)
        expanded_name = f"{test_name}[{pin_suffix}]"
        column = _get_or_create_test_column(
            columns_by_key=state.columns_by_key,
            ordered_columns=state.ordered_columns,
            used_column_ids=state.used_column_ids,
            test_num=test_num,
            test_name=expanded_name,
            unit=unit,
            low=low,
            high=high,
            variant=f"MPR:{pin_suffix}",
        )
        current_part.measurements[column.column_id] = result
        _update_column_accumulator(state, column=column, value=result)
        if current_part.first_fail_test is None:
            current_part.first_fail_test = _detect_failed_test_name(expanded_name, result, low, high, test_flag, param_flag)


def _write_conversion_output(
    state: _ConversionState,
    output_csv_path: Path,
    *,
    source_path: Path | None = None,
    artifacts_output_folder: Path | None = None,
) -> ConversionSummary:
    if state.active_parts_by_site:
        for part in state.active_parts_by_site.values():
            if part.measurements or part.chip_id or part.site_num is not None:
                state.parts.append(part)
        state.active_parts_by_site.clear()

    header = META_COLUMNS + [column.column_id for column in state.ordered_columns]
    with output_csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=DELIMITER)
        writer.writerow(header)
        for summary_name in SUMMARY_ROWS:
            writer.writerow(
                _build_summary_row_from_accumulators(
                    summary_name,
                    state.ordered_columns,
                    state.accumulators_by_column_id,
                )
            )
        for part in state.parts:
            row = [
                _format_number(part.unit_id),
                _format_number(part.site_num),
                _format_number(part.wafer),
                _format_number(part.x),
                _format_number(part.y),
                _format_number(part.lot),
                _format_number(part.sublot),
                _format_number(part.chip_id),
                _format_number(part.pf),
                _format_number(part.first_fail_test),
            ]
            row.extend(_format_number(part.measurements.get(column.column_id)) for column in state.ordered_columns)
            writer.writerow(row)

    artifact_folder = output_csv_path.parent if artifacts_output_folder is None else Path(artifacts_output_folder)
    artifact_folder.mkdir(parents=True, exist_ok=True)

    dtr_file: Path | None = None
    if state.dtr_messages:
        dtr_file = artifact_folder / dtr_name_for_source(output_csv_path.name)
        with dtr_file.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, delimiter=DELIMITER)
            writer.writerow(["Index", "Category", "Message"])
            for index, message in enumerate(state.dtr_messages, start=1):
                match = re.match(r"^([A-Z]+)", message)
                category = match.group(1) if match else "INFO"
                writer.writerow([index, category, message])

    warnings = list(state.warnings)
    if state.malformed_record_types:
        warnings.append(
            "Malformed records skipped: "
            + ", ".join(f"{name}={count}" for name, count in sorted(state.malformed_record_types.items()))
        )

    pir_count = state.record_counts.get("PIR", 0)
    prr_count = state.record_counts.get("PRR", 0)
    has_part_records = pir_count > 0 or prr_count > 0

    consistency_file = artifact_folder / consistency_name_for_source(output_csv_path.name)
    consistency_payload = {
        "csv_file": str(output_csv_path),
        "generated_rows": len(state.parts),
        "numeric_tests": len(state.ordered_columns),
        "record_counts": dict(sorted(state.record_counts.items())),
        "malformed_record_count": int(sum(state.malformed_record_types.values())),
        "malformed_record_types": dict(sorted(state.malformed_record_types.items())),
        "dtr_record_count": len(state.dtr_messages),
        "checks": {
            "pir_equals_prr": True if not has_part_records else pir_count == prr_count,
            "pir_equals_generated_rows": True if not has_part_records else pir_count == len(state.parts),
            "prr_equals_generated_rows": True if not has_part_records else prr_count == len(state.parts),
            "all_rows_have_measurements": all(bool(part.measurements) for part in state.parts),
        },
        "warnings": warnings,
        "dtr_file": None if dtr_file is None else str(dtr_file),
    }
    with consistency_file.open("w", encoding="utf-8") as handle:
        json.dump(consistency_payload, handle, indent=2)

    file_result = ConversionFileResult(
        source_file=source_path or output_csv_path,
        csv_file=output_csv_path,
        dtr_file=dtr_file,
        consistency_file=consistency_file,
        generated_rows=len(state.parts),
        numeric_tests=len(state.ordered_columns),
        record_counts=dict(sorted(state.record_counts.items())),
        malformed_record_count=int(sum(state.malformed_record_types.values())),
        malformed_record_types=dict(sorted(state.malformed_record_types.items())),
        dtr_record_count=len(state.dtr_messages),
        warnings=warnings,
    )

    return ConversionSummary(
        converted_files=1,
        converted_parts=len(state.parts),
        converted_tests=len(state.ordered_columns),
        output_files=[output_csv_path],
        dtr_files=[] if dtr_file is None else [dtr_file],
        consistency_files=[consistency_file],
        warnings=warnings,
        file_results=[file_result],
    )


@contextlib.contextmanager
def _open_stdf_binary(path: Path, *, archive_member_name: str | None = None):
    with contextlib.ExitStack() as stack:
        if archive_member_name is not None:
            archive = stack.enter_context(tarfile.open(path, "r:*"))
            member = archive.getmember(archive_member_name)
            extracted = archive.extractfile(member)
            if extracted is None:
                raise FileNotFoundError(f"Could not open archive member {archive_member_name!r} from {path}")
            stack.callback(extracted.close)
            yield extracted
            return

        suffixes = [suffix.lower() for suffix in path.suffixes]
        if suffixes and suffixes[-1] == ".gz":
            yield stack.enter_context(gzip.open(path, "rb"))
            return
        if suffixes and suffixes[-1] == ".bz2":
            yield stack.enter_context(bz2.open(path, "rb"))
            return
        if suffixes and suffixes[-1] == ".xz":
            yield stack.enter_context(lzma.open(path, "rb"))
            return
        yield stack.enter_context(path.open("rb"))


@contextlib.contextmanager
def _open_eff_text(path: Path, *, archive_member_name: str | None = None):
    with contextlib.ExitStack() as stack:
        if archive_member_name is not None:
            archive = stack.enter_context(tarfile.open(path, "r:*"))
            member = archive.getmember(archive_member_name)
            extracted = archive.extractfile(member)
            if extracted is None:
                raise FileNotFoundError(f"Could not open archive member {archive_member_name!r} from {path}")
            stack.callback(extracted.close)
            text_stream = stack.enter_context(io.TextIOWrapper(extracted, encoding="latin1", errors="replace", newline=""))
            yield text_stream
            return

        yield stack.enter_context(path.open("r", encoding="latin1", errors="replace", newline=""))


def _record_dict_from_pystdf(rec_type: Any, fields: list[Any]) -> dict[str, Any]:
    field_names = list(getattr(rec_type, "fieldNames", []))
    return {name: value for name, value in zip(field_names, fields)}


def _convert_stdf_file_with_pystdf(
    stdf_path: Path,
    output_csv_path: Path,
    *,
    artifacts_output_folder: Path | None = None,
    archive_member_name: str | None = None,
) -> ConversionSummary:
    try:
        import pystdf.V4 as V4
        from pystdf.IO import Parser
    except ImportError as exc:
        raise RuntimeError(
            "STDF conversion requires the 'pystdf' package. Install it from "
            "requirements-tests-data-analysis.txt or with 'pip install pystdf'."
        ) from exc

    state = _ConversionState()

    class _PystdfConverter(Parser):
        def __init__(self, inp) -> None:
            super().__init__(
                recTypes=(V4.far, V4.mir, V4.wir, V4.pir, V4.ptr, V4.mpr, V4.prr, V4.dtr),
                inp=inp,
            )

        def parse_records(self, count=0):
            i = 0
            self.eof = 0
            while self.eof == 0:
                try:
                    header = self.readHeader()
                except Exception as exc:
                    if exc.__class__.__name__ == "EofException":
                        break
                    raise

                self.header(header)
                if (header.typ, header.sub) in self.recordMap:
                    rec_type = self.recordMap[(header.typ, header.sub)]
                    rec_parser = self.recordParsers[(header.typ, header.sub)]
                    try:
                        fields = rec_parser(self, header, [])
                    except Exception as exc:
                        record_name = rec_type.__class__.__name__.upper()
                        state.malformed_record_types[record_name] += 1
                        state.warnings.append(f"Skipped malformed {record_name} record: {exc}")
                        if header.len > 0:
                            self.inp.read(header.len)
                            header.len = 0
                        print(
                            f"Warning: skipped malformed {record_name} record: {exc}",
                            file=sys.stderr,
                        )
                        continue

                    if len(fields) < len(rec_type.columnNames):
                        fields += [None] * (len(rec_type.columnNames) - len(fields))
                    self.send((rec_type, fields))
                    if header.len > 0:
                        self.inp.read(header.len)
                        header.len = 0
                else:
                    self.inp.read(header.len)

                if count:
                    i += 1
                    if i >= count:
                        break

        def send(self, data) -> None:
            rec_type, fields = data
            record_name = rec_type.__class__.__name__.upper()
            if record_name == "FAR":
                return
            _process_record(state, record_name, _record_dict_from_pystdf(rec_type, fields))

    with _open_stdf_binary(stdf_path, archive_member_name=archive_member_name) as handle:
        _PystdfConverter(handle).parse()

    return _write_conversion_output(
        state,
        output_csv_path,
        source_path=stdf_path,
        artifacts_output_folder=artifacts_output_folder,
    )


def _convert_eff_file(
    eff_path: Path,
    output_csv_path: Path,
    *,
    artifacts_output_folder: Path | None = None,
    archive_member_name: str | None = None,
) -> ConversionSummary:
    with _open_eff_text(eff_path, archive_member_name=archive_member_name) as handle:
        reader = csv.reader(handle, delimiter=DELIMITER)
        state = _build_eff_state(reader)

    return _write_conversion_output(
        state,
        output_csv_path,
        source_path=eff_path,
        artifacts_output_folder=artifacts_output_folder,
    )


def convert_records_to_csv(
    records: Iterable[Any],
    output_csv_path: Path,
    *,
    source_name: str = "<memory>",
    artifacts_output_folder: Path | None = None,
) -> ConversionSummary:
    output_csv_path = Path(output_csv_path)
    output_csv_path.parent.mkdir(parents=True, exist_ok=True)

    state = _ConversionState()

    for record in records:
        record_dict = _record_to_dict(record)
        record_name = _record_id(record, record_dict).upper()
        _process_record(state, record_name, record_dict)

    return _write_conversion_output(
        state,
        output_csv_path,
        source_path=Path(source_name),
        artifacts_output_folder=artifacts_output_folder,
    )


def convert_stdf_file(
    stdf_path: Path,
    output_csv_path: Path,
    *,
    artifacts_output_folder: Path | None = None,
) -> ConversionSummary:
    archive_member_name: str | None = None
    if _is_tar_archive_path(stdf_path):
        archive_member_name = _select_archive_member_name(stdf_path)
        source_kind = _source_kind_for_name(Path(archive_member_name).name)
    else:
        source_kind = _source_kind_for_name(stdf_path.name)

    if source_kind == "eff":
        return _convert_eff_file(
            stdf_path,
            output_csv_path,
            artifacts_output_folder=artifacts_output_folder,
            archive_member_name=archive_member_name,
        )

    return _convert_stdf_file_with_pystdf(
        stdf_path,
        output_csv_path,
        artifacts_output_folder=artifacts_output_folder,
        archive_member_name=archive_member_name,
    )


def _convert_stdf_file_job(
    stdf_path_str: str,
    output_csv_path_str: str,
    artifacts_output_folder_str: str | None,
) -> ConversionSummary:
    artifacts_output_folder = None if artifacts_output_folder_str is None else Path(artifacts_output_folder_str)
    return convert_stdf_file(
        Path(stdf_path_str),
        Path(output_csv_path_str),
        artifacts_output_folder=artifacts_output_folder,
    )


def _iter_stdf_files(input_folder: Path, patterns: Sequence[str]) -> list[Path]:
    matches: dict[str, Path] = {}
    for pattern in patterns:
        for path in input_folder.glob(pattern):
            if path.is_file():
                matches[str(path.resolve()).lower()] = path
    return sorted(matches.values(), key=lambda item: item.name.lower())


def convert_stdf_folder(
    input_folder: Path,
    output_folder: Path,
    *,
    patterns: Sequence[str] | None = None,
    single_file: str | None = None,
    max_files: int | None = None,
    artifacts_output_folder: Path | None = None,
) -> ConversionSummary:
    input_folder = Path(input_folder)
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    chosen_patterns = tuple(patterns or DEFAULT_PATTERNS)

    if single_file:
        stdf_files = [input_folder / single_file]
    else:
        stdf_files = _iter_stdf_files(input_folder, chosen_patterns)
    if max_files is not None:
        stdf_files = stdf_files[:max_files]

    output_files: list[Path] = []
    dtr_files: list[Path] = []
    consistency_files: list[Path] = []
    warnings: list[str] = []
    file_results: list[ConversionFileResult] = []
    converted_parts = 0
    converted_tests = 0

    total_files = len(stdf_files)
    if total_files == 0:
        print("[STDF files] 100% (0/0) | no STDF files matched the current selection")
    else:
        print(f"[STDF files]   0% (0/{total_files}) | starting STDF to CSV conversion")

    for stdf_file in stdf_files:
        if not stdf_file.exists():
            raise FileNotFoundError(f"STDF file not found: {stdf_file}")

    job_specs = [
        (
            stdf_file,
            output_folder / csv_name_for_source(stdf_file.name),
        )
        for stdf_file in stdf_files
    ]

    max_workers = min(total_files, max(1, os.cpu_count() or 1))
    completed_count = 0

    if total_files <= 1 or max_workers <= 1:
        for stdf_file, output_csv_path in job_specs:
            percent = int(round(100.0 * completed_count / total_files)) if total_files > 0 else 100
            print(f"[STDF files] {percent:3d}% ({completed_count}/{total_files}) | converting {stdf_file.name}")
            summary = convert_stdf_file(
                stdf_file,
                output_csv_path,
                artifacts_output_folder=artifacts_output_folder,
            )
            completed_count += 1
            output_files.extend(summary.output_files)
            dtr_files.extend(summary.dtr_files)
            consistency_files.extend(summary.consistency_files)
            warnings.extend(summary.warnings)
            file_results.extend(summary.file_results)
            converted_parts += summary.converted_parts
            converted_tests += summary.converted_tests
            print(
                f"[STDF files] {int(round(100.0 * completed_count / total_files)):3d}% ({completed_count}/{total_files})"
                f" | finished {stdf_file.name}"
            )
    else:
        print(f"[STDF files]   0% (0/{total_files}) | converting with {max_workers} worker processes")
        with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(
                    _convert_stdf_file_job,
                    str(stdf_file),
                    str(output_csv_path),
                    None if artifacts_output_folder is None else str(artifacts_output_folder),
                ): (stdf_file, output_csv_path)
                for stdf_file, output_csv_path in job_specs
            }
            for future in concurrent.futures.as_completed(future_map):
                stdf_file, _ = future_map[future]
                summary = future.result()
                completed_count += 1
                output_files.extend(summary.output_files)
                dtr_files.extend(summary.dtr_files)
                consistency_files.extend(summary.consistency_files)
                warnings.extend(summary.warnings)
                file_results.extend(summary.file_results)
                converted_parts += summary.converted_parts
                converted_tests += summary.converted_tests
                print(
                    f"[STDF files] {int(round(100.0 * completed_count / total_files)):3d}% ({completed_count}/{total_files})"
                    f" | finished {stdf_file.name}"
                )

        output_files.sort(key=lambda item: item.name.lower())
        dtr_files.sort(key=lambda item: item.name.lower())
        consistency_files.sort(key=lambda item: item.name.lower())
        file_results.sort(key=lambda item: item.csv_file.name.lower())

    return ConversionSummary(
        converted_files=len(output_files),
        converted_parts=converted_parts,
        converted_tests=converted_tests,
        output_files=output_files,
        dtr_files=dtr_files,
        consistency_files=consistency_files,
        warnings=warnings,
        file_results=file_results,
    )


def _build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert STDF/EFF inputs into the flat CSV format expected by Test_Data_Reviewer.py")
    parser.add_argument("--input-folder", type=Path, required=True, help="Folder containing .stdf/.std/.eff inputs")
    parser.add_argument("--output-folder", type=Path, required=True, help="Folder where flat CSV files will be written")
    parser.add_argument("--single-file", type=str, default=None, help="Optional single input file name to convert")
    parser.add_argument("--max-files", type=int, default=None, help="Optional maximum number of input files to convert")
    parser.add_argument(
        "--artifacts-folder",
        type=Path,
        default=None,
        help="Optional folder where DTR and consistency sidecar files will be written",
    )
    parser.add_argument(
        "--pattern",
        action="append",
        dest="patterns",
        help="Optional glob pattern(s) for supported input discovery. Can be repeated.",
    )
    return parser


def main() -> int:
    multiprocessing.freeze_support()
    parser = _build_argument_parser()
    args = parser.parse_args()
    summary = convert_stdf_folder(
        input_folder=args.input_folder,
        output_folder=args.output_folder,
        patterns=args.patterns,
        single_file=args.single_file,
        max_files=args.max_files,
        artifacts_output_folder=args.artifacts_folder,
    )
    print(f"Converted {summary.converted_files} input file(s) into {args.output_folder}")
    print(f"  Parts exported: {summary.converted_parts}")
    print(f"  Numeric tests exported: {summary.converted_tests}")
    for output_file in summary.output_files:
        print(f"  - {output_file}")
    for dtr_file in summary.dtr_files:
        print(f"  DTR: {dtr_file}")
    for consistency_file in summary.consistency_files:
        print(f"  Consistency: {consistency_file}")
    for warning in summary.warnings:
        print(f"  Warning: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
