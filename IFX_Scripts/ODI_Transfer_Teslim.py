"""ODI import CSV generator (flat execution, no CLI).

Reads ODI (offset/drift information) from the reference Excel workbook and generates a
semicolon-separated CSV compatible with the `ImportOdi` command described in
`ImportOdi.md`.

Key rules implemented:
- Read ODI from sheets 50_TXGE ... 58_TXPS.
- Module name = substring after underscore in sheet name (e.g. 52_DPLL -> DPLL).
- Test number from column "Test Number" (formula results via openpyxl data_only=True).
- ODI values from columns whose header starts with "ODI" (also formula results).
- Only non-zero ODI values are exported.
- ODI scope: header contains LTL -> LowerLimit; contains UTL -> UpperLimit.
- Insertion name: last token of the ODI column header (e.g. "ODI LTL S1" -> S1).
- Output mapping: ODI value goes to column "S"; column "R" is kept empty.

Configure parameters in CONFIG below and run this file directly.
"""

from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import csv
from typing import Iterable, Optional
from openpyxl import load_workbook


@dataclass(frozen=True)
class OdiImportConfig:
	reference_xlsx: Path
	output_csv: Path

	# Selection
	modules: Optional[set[str]] = None  # None => all modules (50..58)
	odi_columns: Optional[set[str]] = None  # None => all ODI columns; else match by header

	# Output fields
	comment: str = ""
	jira_tasks: str = ""  # comma-separated (e.g. "RSIPPTE-493,ABC-123")
	odi_source: str = ""  # e.g. "ATE", "INS_OFF"
	test_variants: str = ""  # comma-separated variants (e.g. "8188")


CONFIG = OdiImportConfig(
	# Reference workbook (the repo contains CTRX8188A_TE_TX.xlsx)
	reference_xlsx=Path(__file__).resolve().parents[2] / "CTRX8188A_TE_TX.xlsx",

	# Output file to generate
	output_csv=Path(__file__).resolve().parents[2] / "ODI_FE_Test_CS_Workaround_2.csv",

	# Optional filters (examples):
	# modules={"DPLL", "TXGE"},
	# odi_columns={"ODI UTL B1", "ODI UTL B2"},
	#modules=None, 
	#odi_columns=None,
    #modules={"TXGE","DPLL","TXPA","TXPB","TXPC","TXPD","TXLO","TXPS"},
	#odi_columns={"ODI LTL S1","ODI UTL S1","ODI LTL S2","ODI UTL S2"},
	modules={"TXPA","TXPB","TXPC","TXPD"},
	odi_columns={"ODI LTL S1"},
	# Output metadata
	comment="Workaround to remove some ODIs because of current Teslim issues 2",
	jira_tasks="RSIPPTE-493",
	odi_source="INS_OFF",
	test_variants="8188",
)


CSV_HEADER = [
	"TestNumber",
	"OdiSource",
	"Insertions",
	"TestVariants",
	"OdiScope",
	"R",
	"S",
	"CalcMethod",
	"Link2Evidence",
	"Comment",
	"JiraTasks",
]


@dataclass(frozen=True)
class GenerationResult:
	output_csv: Path
	row_count: int
	processed_sheets: tuple[str, ...]
	selected_modules: Optional[tuple[str, ...]]
	selected_odi_columns: Optional[tuple[str, ...]]


def _normalize_header(value: str) -> str:
	return " ".join(value.strip().split()).casefold()


def _is_nonzero(value: object) -> bool:
	if value is None:
		return False
	if isinstance(value, bool):
		return bool(value)
	if isinstance(value, (int, float)):
		return abs(float(value)) > 1e-12
	if isinstance(value, str):
		text = value.strip()
		if text == "":
			return False
		try:
			return abs(float(text)) > 1e-12
		except ValueError:
			return True
	return True


def _to_float(value: object) -> float:
	if isinstance(value, (int, float)):
		return float(value)
	if isinstance(value, str):
		return float(value.strip())
	raise TypeError(f"Unsupported ODI value type: {type(value)}")


def _normalize_s_value(value: float) -> float:
	# Excel cached results often contain tiny FP noise; normalize for grouping and output.
	if abs(value) <= 1e-12:
		return 0.0
	return round(value, 12)


def _format_s_value(value: float) -> str:
	if abs(value - round(value)) <= 1e-12:
		return str(int(round(value)))
	# Significant digits formatting to avoid artifacts like 0.9000000000000004
	return format(value, ".15g")


def _extract_scope_and_insertion(odi_header: str) -> tuple[str, str]:
	normalized = _normalize_header(odi_header)
	if "ltl" in normalized:
		scope = "LowerLimit"
	elif "utl" in normalized:
		scope = "UpperLimit"
	else:
		raise ValueError(
			f"ODI column header must contain LTL or UTL to derive scope: {odi_header!r}"
		)

	# Last token is the insertion name (S1, S2, B1, ...)
	tokens = odi_header.strip().split()
	if not tokens:
		raise ValueError("Empty ODI header")
	insertion = tokens[-1]
	return scope, insertion


def _iter_target_sheets(sheetnames: Iterable[str]) -> list[str]:
	targets: list[str] = []
	for name in sheetnames:
		# Expect e.g. "50_TXGE" ... "58_TXPS"
		if "_" not in name:
			continue
		prefix, _ = name.split("_", 1)
		if not prefix.isdigit():
			continue
		idx = int(prefix)
		if 50 <= idx <= 58:
			targets.append(name)
	return targets


def generate_odi_csv(config: OdiImportConfig) -> GenerationResult:
	if not config.reference_xlsx.exists():
		raise FileNotFoundError(f"Reference workbook not found: {config.reference_xlsx}")

	config.output_csv.parent.mkdir(parents=True, exist_ok=True)

	# data_only=True is required to read cached formula results.
	wb = load_workbook(config.reference_xlsx, data_only=True, read_only=True)
	sheetnames = _iter_target_sheets(wb.sheetnames)

	selected_modules_norm: Optional[set[str]] = (
		None
		if config.modules is None
		else {m.strip().casefold() for m in config.modules if m.strip()}
	)

	normalized_selected_odi_cols: Optional[set[str]] = (
		None
		if config.odi_columns is None
		else {_normalize_header(c) for c in config.odi_columns}
	)

	# Group rows when multiple insertions have same TestNumber + scope + S value
	# (Example ODI_Test.csv groups B1,B2 for the same value)
	grouped: dict[tuple[int, str, float], set[str]] = {}
	processed_sheets: list[str] = []

	for sheet_name in sheetnames:
		module = sheet_name.split("_", 1)[1]
		if selected_modules_norm is not None and module.strip().casefold() not in selected_modules_norm:
			continue
		processed_sheets.append(sheet_name)

		ws = wb[sheet_name]

		header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
		if not header_row:
			continue

		headers_str = [h if isinstance(h, str) else "" for h in header_row]

		try:
			test_number_idx = headers_str.index("Test Number")
		except ValueError:
			continue

		odi_cols: list[tuple[int, str]] = []  # (0-based index, header)
		for idx, header in enumerate(headers_str):
			if not header:
				continue
			if not header.strip().startswith("ODI"):
				continue
			if normalized_selected_odi_cols is not None:
				if _normalize_header(header) not in normalized_selected_odi_cols:
					continue
			odi_cols.append((idx, header))

		if not odi_cols:
			continue

		blank_streak = 0
		for row in ws.iter_rows(min_row=2, values_only=True):
			if test_number_idx >= len(row):
				continue

			test_number = row[test_number_idx]
			if test_number in (None, ""):
				blank_streak += 1
				if blank_streak > 200:
					break
				continue
			blank_streak = 0

			# Excel might store it as float even if displayed as int
			try:
				test_number_int = int(float(test_number))
			except Exception:
				continue

			for col_idx, header in odi_cols:
				if col_idx >= len(row):
					continue
				odi_value = row[col_idx]
				if not _is_nonzero(odi_value):
					continue

				scope, insertion = _extract_scope_and_insertion(header)
				s_value = _normalize_s_value(_to_float(odi_value))
				if abs(s_value) <= 1e-12:
					continue

				key = (test_number_int, scope, s_value)
				grouped.setdefault(key, set()).add(insertion)

	rows = []
	for (test_number, scope, s_value), insertions in grouped.items():
		rows.append(
			{
				"TestNumber": str(test_number),
				"OdiSource": config.odi_source,
				"Insertions": ",".join(sorted(insertions)),
				"TestVariants": config.test_variants,
				"OdiScope": scope,
				"R": "",
				"S": _format_s_value(s_value),
				"CalcMethod": "",
				"Link2Evidence": "",
				"Comment": config.comment,
				"JiraTasks": config.jira_tasks,
			}
		)

	rows.sort(key=lambda r: (int(r["TestNumber"]), r["OdiScope"], r["Insertions"]))

	with config.output_csv.open("w", newline="", encoding="utf-8") as f:
		writer = csv.DictWriter(f, fieldnames=CSV_HEADER, delimiter=";")
		writer.writeheader()
		writer.writerows(rows)

	return GenerationResult(
		output_csv=config.output_csv,
		row_count=len(rows),
		processed_sheets=tuple(processed_sheets),
		selected_modules=None if config.modules is None else tuple(sorted(config.modules)),
		selected_odi_columns=None if config.odi_columns is None else tuple(sorted(config.odi_columns)),
	)


def main() -> None:
	result = generate_odi_csv(CONFIG)
	print(f"Generated ODI CSV: {result.output_csv}")
	print(f"Rows written: {result.row_count}")
	print(f"Sheets processed: {len(result.processed_sheets)}")

	if result.row_count == 0:
		selected_modules = "ALL" if result.selected_modules is None else ", ".join(result.selected_modules)
		selected_odi_cols = "ALL" if result.selected_odi_columns is None else ", ".join(result.selected_odi_columns)
		print(
			"WARNING: No ODI rows were produced. Common causes:\n"
			"- The selected module(s) or ODI column name(s) do not match the sheet headers exactly.\n"
			"- The workbook contains formulas but does not have cached results saved.\n"
			"  openpyxl cannot calculate formulas; it can only read cached values (data_only=True).\n"
			"  Fix: open the workbook in Excel, force recalculation (Ctrl+Alt+F9), then save it.\n"
			f"Selected modules: {selected_modules}\n"
			f"Selected ODI columns: {selected_odi_cols}\n"
			f"Sheets processed: {result.processed_sheets}"
		)


if __name__ == "__main__":
	main()

