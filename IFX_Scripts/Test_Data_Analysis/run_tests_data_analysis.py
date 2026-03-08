from __future__ import annotations
import argparse
import importlib
from pathlib import Path
from typing import Any
import Tests_Data_Analysis as analysis
import stdf_to_flat_csv


TEST_DATA_ANALYSIS_ROOT = Path(__file__).resolve().parent


ALLOWED_CORRELATION_METHODS = {"pearson", "spearman"}
REQUIRED_CONFIG_KEYS = {"input_folder", "output_folder", "modules"}
OPTIONAL_CONFIG_KEYS = {
    "yield_threshold",
    "cpk_low",
    "cpk_high",
    "outlier_mad_multiplier",
    "max_files",
    "single_file",
    "encoding",
    "generate_correlation_report",
    "correlation_methods",
    "pearson_abs_min_for_report",
    "wafermap_circle_area_mult",
    "convert_stdf_before_analysis",
    "stdf_single_file",
    "stdf_file_patterns",
    "stdf_input_folder",
}
ALLOWED_CONFIG_KEYS = REQUIRED_CONFIG_KEYS | OPTIONAL_CONFIG_KEYS


class ConfigValidationError(ValueError):
    """Raised when a YAML configuration file fails schema validation."""


def _progress_percent(current: int, total: int) -> int:
    if total <= 0:
        return 100
    bounded = min(max(current, 0), total)
    return int(round(100.0 * bounded / total))


def _print_progress(stage: str, current: int, total: int, detail: str | None = None) -> None:
    pct = _progress_percent(current, total)
    suffix = "" if not detail else f" | {detail}"
    print(f"[{stage}] {pct:3d}% ({current}/{total}){suffix}")


def _resolve_from_test_data_analysis_root(path_value: str | Path) -> Path:
    candidate = path_value if isinstance(path_value, Path) else Path(path_value)
    if candidate.is_absolute():
        return candidate

    first_part = candidate.parts[0] if candidate.parts else ""
    search_roots = (TEST_DATA_ANALYSIS_ROOT, *TEST_DATA_ANALYSIS_ROOT.parents)
    for base in search_roots:
        resolved = (base / candidate).resolve()
        anchor = (base / first_part) if first_part else resolved
        if resolved.exists() or anchor.exists():
            return resolved

    return (TEST_DATA_ANALYSIS_ROOT / candidate).resolve()


def _resolve_config_path(config_arg: Path) -> Path:
    candidates = [config_arg]
    if not config_arg.is_absolute():
        candidates.append(TEST_DATA_ANALYSIS_ROOT / config_arg)

    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.exists():
            return resolved

    return candidates[-1].resolve()


def _format_config_errors(errors: list[str]) -> str:
    bullet_list = "\n".join(f"- {item}" for item in errors)
    return f"Invalid YAML config:\n{bullet_list}"


def _is_sequence_but_not_string(value: Any) -> bool:
    return isinstance(value, list | tuple)


def _has_existing_csv_inputs(input_folder: Path | None, single_file: str | None) -> bool:
    if input_folder is None or not input_folder.exists() or not input_folder.is_dir():
        return False

    if single_file:
        csv_name = str(single_file)
        if not csv_name.lower().endswith(".csv"):
            csv_name = stdf_to_flat_csv.csv_name_for_source(csv_name)
        return (input_folder / csv_name).is_file()

    return any(path.is_file() for path in input_folder.glob("*.csv"))


def _validate_and_normalize_config(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    errors: list[str] = []
    normalized: dict[str, Any] = {}
    convert_stdf_before_analysis = bool(config.get("convert_stdf_before_analysis", False))

    unknown_keys = sorted(set(config) - ALLOWED_CONFIG_KEYS)
    if unknown_keys:
        errors.append(
            "unknown key(s): " + ", ".join(unknown_keys)
        )

    missing_required_keys = sorted(key for key in REQUIRED_CONFIG_KEYS if key not in config)
    if missing_required_keys:
        errors.append(
            "missing required key(s): " + ", ".join(missing_required_keys)
        )

    def require_non_empty_string(key: str) -> None:
        value = config.get(key)
        if not isinstance(value, str) or not value.strip():
            errors.append(f"{key} must be a non-empty string")
            return
        normalized[key] = value.strip()

    require_non_empty_string("input_folder")
    require_non_empty_string("output_folder")

    input_folder = normalized.get("input_folder")
    if isinstance(input_folder, str):
        input_path = _resolve_from_test_data_analysis_root(input_folder)
        normalized["input_folder"] = input_path
        if not input_path.exists():
            if not convert_stdf_before_analysis:
                errors.append(f"input_folder does not exist: {input_path}")
            else:
                input_parent = input_path.parent if input_path.parent != Path("") else Path(".")
                if not input_parent.exists():
                    errors.append(f"input_folder parent directory does not exist: {input_parent}")
        elif not input_path.is_dir():
            errors.append(f"input_folder is not a directory: {input_path}")

    output_folder = normalized.get("output_folder")
    if isinstance(output_folder, str):
        output_path = _resolve_from_test_data_analysis_root(output_folder)
        normalized["output_folder"] = output_path
        output_parent = output_path.parent if output_path.parent != Path("") else Path(".")
        if not output_parent.exists():
            errors.append(f"output_folder parent directory does not exist: {output_parent}")

    modules = config.get("modules")
    if not _is_sequence_but_not_string(modules):
        errors.append("modules must be a list of non-empty strings")
    else:
        cleaned_modules = [str(item).strip().upper() for item in modules if str(item).strip()]
        if not cleaned_modules:
            errors.append("modules must contain at least one non-empty value")
        elif any(len(module) < 4 for module in cleaned_modules):
            errors.append("modules entries must be at least 4 characters long")
        else:
            normalized["modules"] = cleaned_modules

    def normalize_float(
        key: str,
        *,
        minimum: float | None = None,
        maximum: float | None = None,
        inclusive_min: bool = True,
        inclusive_max: bool = True,
    ) -> None:
        if key not in config:
            return
        value = config[key]
        try:
            number = float(value)
        except (TypeError, ValueError):
            errors.append(f"{key} must be a number")
            return

        if minimum is not None:
            below_min = number < minimum if inclusive_min else number <= minimum
            if below_min:
                operator = ">=" if inclusive_min else ">"
                errors.append(f"{key} must be {operator} {minimum}")
        if maximum is not None:
            above_max = number > maximum if inclusive_max else number >= maximum
            if above_max:
                operator = "<=" if inclusive_max else "<"
                errors.append(f"{key} must be {operator} {maximum}")
        normalized[key] = number

    normalize_float("yield_threshold", minimum=0.0, maximum=100.0)
    normalize_float("cpk_low", minimum=0.0)
    normalize_float("cpk_high", minimum=0.0)
    normalize_float("outlier_mad_multiplier", minimum=0.0, inclusive_min=False)
    normalize_float("pearson_abs_min_for_report", minimum=0.0, maximum=1.0)
    normalize_float("wafermap_circle_area_mult", minimum=0.0, inclusive_min=False)

    if "cpk_low" in normalized and "cpk_high" in normalized:
        if float(normalized["cpk_high"]) < float(normalized["cpk_low"]):
            errors.append("cpk_high must be greater than or equal to cpk_low")

    if "max_files" in config:
        value = config["max_files"]
        if value is None:
            normalized["max_files"] = None
        elif isinstance(value, bool):
            errors.append("max_files must be an integer greater than 0 or null")
        else:
            try:
                max_files = int(value)
            except (TypeError, ValueError):
                errors.append("max_files must be an integer greater than 0 or null")
            else:
                if max_files <= 0:
                    errors.append("max_files must be greater than 0 when provided")
                else:
                    normalized["max_files"] = max_files

    if "single_file" in config:
        value = config["single_file"]
        if value in (None, "", "null"):
            normalized["single_file"] = None
        elif not isinstance(value, str):
            errors.append("single_file must be a string or null")
        elif not value.strip():
            normalized["single_file"] = None
        else:
            normalized["single_file"] = value.strip()

    if "encoding" in config:
        value = config["encoding"]
        if not isinstance(value, str) or not value.strip():
            errors.append("encoding must be a non-empty string")
        else:
            normalized["encoding"] = value.strip()

    if "convert_stdf_before_analysis" in config:
        value = config["convert_stdf_before_analysis"]
        if not isinstance(value, bool):
            errors.append("convert_stdf_before_analysis must be true or false")
        else:
            normalized["convert_stdf_before_analysis"] = value

    if "stdf_input_folder" in config:
        value = config["stdf_input_folder"]
        if value not in (None, "", "null") and (not isinstance(value, str) or not value.strip()):
            errors.append("stdf_input_folder must be a non-empty string or null")

    if "stdf_single_file" in config:
        value = config["stdf_single_file"]
        if value in (None, "", "null"):
            normalized["stdf_single_file"] = None
        elif not isinstance(value, str):
            errors.append("stdf_single_file must be a string or null")
        else:
            normalized["stdf_single_file"] = value.strip() or None

    if "stdf_file_patterns" in config:
        value = config["stdf_file_patterns"]
        if not _is_sequence_but_not_string(value):
            errors.append("stdf_file_patterns must be a list")
        else:
            patterns = [str(item).strip() for item in value if str(item).strip()]
            if not patterns:
                errors.append("stdf_file_patterns must contain at least one non-empty value")
            else:
                normalized["stdf_file_patterns"] = patterns

    if "generate_correlation_report" in config:
        value = config["generate_correlation_report"]
        if not isinstance(value, bool):
            errors.append("generate_correlation_report must be true or false")
        else:
            normalized["generate_correlation_report"] = value

    if "correlation_methods" in config:
        value = config["correlation_methods"]
        if not _is_sequence_but_not_string(value):
            errors.append("correlation_methods must be a list")
        else:
            methods = [str(item).strip().lower() for item in value if str(item).strip()]
            invalid_methods = sorted({method for method in methods if method not in ALLOWED_CORRELATION_METHODS})
            if invalid_methods:
                errors.append(
                    "correlation_methods contains unsupported value(s): " + ", ".join(invalid_methods)
                )
            elif "generate_correlation_report" in normalized and normalized["generate_correlation_report"] and not methods:
                errors.append("correlation_methods must contain at least one method when generate_correlation_report is true")
            else:
                normalized["correlation_methods"] = methods

    if normalized.get("generate_correlation_report") and "correlation_methods" not in config:
        errors.append("correlation_methods is required when generate_correlation_report is true")

    if errors:
        raise ConfigValidationError(f"{config_path}\n{_format_config_errors(errors)}")

    return normalized


def _load_config(config_path: Path) -> dict[str, Any]:
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    yaml_mod = importlib.import_module("yaml")
    with config_path.open("r", encoding="utf-8") as handle:
        data = yaml_mod.safe_load(handle) or {}

    if not isinstance(data, dict):
        raise ValueError("Config root must be a mapping/object")

    return _validate_and_normalize_config(data, config_path)


def _apply_config(config: dict[str, Any]) -> None:
    if "input_folder" in config:
        analysis.INPUT_FOLDER = Path(str(config["input_folder"]))
    if "output_folder" in config:
        analysis.OUTPUT_FOLDER = Path(str(config["output_folder"]))
    if "modules" in config:
        analysis.MODULES = [str(item).strip().upper() for item in list(config["modules"]) if str(item).strip()]
    if "yield_threshold" in config:
        analysis.YIELD_THRESHOLD = float(config["yield_threshold"])
    if "cpk_low" in config:
        analysis.CPK_LOW = float(config["cpk_low"])
    if "cpk_high" in config:
        analysis.CPK_HIGH = float(config["cpk_high"])
    if "outlier_mad_multiplier" in config:
        analysis.OUTLIER_MAD_MULTIPLIER = float(config["outlier_mad_multiplier"])
    if "max_files" in config:
        analysis.MAX_FILES = None if config["max_files"] is None else int(config["max_files"])
    if "single_file" in config:
        value = config["single_file"]
        analysis.SINGLE_FILE = None if value in (None, "", "null") else str(value)
    if "encoding" in config:
        analysis.ENCODING = str(config["encoding"])
    if "generate_correlation_report" in config:
        analysis.GENERATE_CORRELATION_REPORT = bool(config["generate_correlation_report"])
    if "correlation_methods" in config:
        methods = [str(item).strip().lower() for item in list(config["correlation_methods"]) if str(item).strip()]
        analysis.CORRELATION_METHODS = [m for m in methods if m in {"pearson", "spearman"}]
    if "pearson_abs_min_for_report" in config:
        analysis.PEARSON_ABS_MIN_FOR_REPORT = float(config["pearson_abs_min_for_report"])
    if "wafermap_circle_area_mult" in config:
        analysis.WAFERMAP_CIRCLE_AREA_MULT = float(config["wafermap_circle_area_mult"])


def _prepare_stdf_inputs(config: dict[str, Any]) -> None:
    _print_progress(
        "STDF conversion",
        0,
        1,
        "checking whether STDF pre-conversion is required",
    )
    if not config.get("convert_stdf_before_analysis"):
        _print_progress("STDF conversion", 1, 1, "disabled by configuration")
        return

    input_folder = Path(str(config["input_folder"]))
    input_folder.mkdir(parents=True, exist_ok=True)

    target_csv_name: str | None = None
    if analysis.SINGLE_FILE:
        target_csv_name = str(analysis.SINGLE_FILE)
        if not target_csv_name.lower().endswith(".csv"):
            target_csv_name = stdf_to_flat_csv.csv_name_for_source(target_csv_name)

    existing_csvs = sorted(path for path in input_folder.glob("*.csv") if path.is_file())
    skip_preconversion = False
    if target_csv_name:
        skip_preconversion = (input_folder / target_csv_name).is_file()
    else:
        skip_preconversion = bool(existing_csvs)

    if skip_preconversion:
        if target_csv_name:
            analysis.SINGLE_FILE = target_csv_name
            _print_progress(
                "STDF conversion",
                1,
                1,
                f"skipped; existing CSV already available ({target_csv_name})",
            )
            print(f"STDF pre-conversion skipped: existing CSV input already available in {input_folder} ({target_csv_name})")
        else:
            _print_progress(
                "STDF conversion",
                1,
                1,
                f"skipped; found {len(existing_csvs)} existing CSV input file(s)",
            )
            print(f"STDF pre-conversion skipped: found {len(existing_csvs)} existing CSV input file(s) in {input_folder}")
        return

    _print_progress("STDF conversion", 0, 1, "converting STDF files into CSV")
    artifacts_folder = Path(str(config["output_folder"])) / "Artifacts"
    summary = stdf_to_flat_csv.convert_stdf_folder(
        input_folder=input_folder,
        output_folder=input_folder,
        patterns=list(config.get("stdf_file_patterns") or stdf_to_flat_csv.DEFAULT_PATTERNS),
        single_file=config.get("stdf_single_file"),
        max_files=analysis.MAX_FILES,
        artifacts_output_folder=artifacts_folder,
    )
    if summary.converted_files == 0:
        raise RuntimeError(f"No STDF files were converted from {input_folder}")

    if analysis.SINGLE_FILE and not str(analysis.SINGLE_FILE).lower().endswith(".csv"):
        analysis.SINGLE_FILE = stdf_to_flat_csv.csv_name_for_source(str(analysis.SINGLE_FILE))

    _print_progress(
        "STDF conversion",
        1,
        1,
        f"completed; {summary.converted_files} file(s), {summary.converted_parts} part(s), {summary.converted_tests} numeric test column(s)",
    )
    if summary.dtr_files:
        print(f"  DTR sidecar file(s): {len(summary.dtr_files)}")
    if summary.consistency_files:
        print(f"  Consistency report file(s): {len(summary.consistency_files)}")
        print(f"  Artifacts folder: {artifacts_folder}")
    for warning in summary.warnings:
        print(f"  Warning: {warning}")


def _build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run Tests_Data_Analysis.py with a YAML config file")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("configs/config_default.yaml"),
        help="Path to YAML config file",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Load and print resolved config values without running analysis",
    )
    return parser


def main() -> int:
    parser = _build_argument_parser()
    args = parser.parse_args()

    workflow_total_steps = 4
    _print_progress("Workflow", 0, workflow_total_steps, "starting launcher workflow")

    config_path = _resolve_config_path(args.config)
    config = _load_config(config_path)
    _print_progress("Workflow", 1, workflow_total_steps, "configuration loaded and validated")
    _apply_config(config)
    _print_progress("Workflow", 2, workflow_total_steps, "runtime configuration applied")

    resolved = {
        "INPUT_FOLDER": str(analysis.INPUT_FOLDER),
        "OUTPUT_FOLDER": str(analysis.OUTPUT_FOLDER),
        "MODULES": analysis.MODULES,
        "YIELD_THRESHOLD": analysis.YIELD_THRESHOLD,
        "CPK_LOW": analysis.CPK_LOW,
        "CPK_HIGH": analysis.CPK_HIGH,
        "OUTLIER_MAD_MULTIPLIER": analysis.OUTLIER_MAD_MULTIPLIER,
        "MAX_FILES": analysis.MAX_FILES,
        "SINGLE_FILE": analysis.SINGLE_FILE,
        "ENCODING": analysis.ENCODING,
        "GENERATE_CORRELATION_REPORT": analysis.GENERATE_CORRELATION_REPORT,
        "CORRELATION_METHODS": analysis.CORRELATION_METHODS,
        "PEARSON_ABS_MIN_FOR_REPORT": analysis.PEARSON_ABS_MIN_FOR_REPORT,
        "WAFERMAP_CIRCLE_AREA_MULT": analysis.WAFERMAP_CIRCLE_AREA_MULT,
        "CONVERT_STDF_BEFORE_ANALYSIS": bool(config.get("convert_stdf_before_analysis", False)),
        "STDF_SINGLE_FILE": config.get("stdf_single_file"),
        "STDF_FILE_PATTERNS": list(config.get("stdf_file_patterns") or stdf_to_flat_csv.DEFAULT_PATTERNS),
    }

    print(f"Using config: {config_path}")
    for key, value in resolved.items():
        print(f"  {key}: {value}")

    if args.dry_run:
        _print_progress("Workflow", workflow_total_steps, workflow_total_steps, "dry-run completed")
        print("Dry-run mode: no analysis executed.")
        return 0

    _prepare_stdf_inputs(config)
    _print_progress("Workflow", 3, workflow_total_steps, "starting data analysis and report generation")

    result = int(analysis.run())
    _print_progress("Workflow", workflow_total_steps, workflow_total_steps, "workflow completed")
    return result


if __name__ == "__main__":
    raise SystemExit(main())
