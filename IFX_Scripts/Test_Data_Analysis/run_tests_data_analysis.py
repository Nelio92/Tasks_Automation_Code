from __future__ import annotations
import argparse
import importlib
from pathlib import Path
from typing import Any
import Tests_Data_Analysis as analysis


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
}
ALLOWED_CONFIG_KEYS = REQUIRED_CONFIG_KEYS | OPTIONAL_CONFIG_KEYS


class ConfigValidationError(ValueError):
    """Raised when a YAML configuration file fails schema validation."""


def _format_config_errors(errors: list[str]) -> str:
    bullet_list = "\n".join(f"- {item}" for item in errors)
    return f"Invalid YAML config:\n{bullet_list}"


def _is_sequence_but_not_string(value: Any) -> bool:
    return isinstance(value, list | tuple)


def _validate_and_normalize_config(config: dict[str, Any], config_path: Path) -> dict[str, Any]:
    errors: list[str] = []
    normalized: dict[str, Any] = {}

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
        input_path = Path(input_folder)
        if not input_path.exists():
            errors.append(f"input_folder does not exist: {input_path}")
        elif not input_path.is_dir():
            errors.append(f"input_folder is not a directory: {input_path}")

    output_folder = normalized.get("output_folder")
    if isinstance(output_folder, str):
        output_path = Path(output_folder)
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

    config_path = args.config.resolve()
    config = _load_config(config_path)
    _apply_config(config)

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
    }

    print(f"Using config: {config_path}")
    for key, value in resolved.items():
        print(f"  {key}: {value}")

    if args.dry_run:
        print("Dry-run mode: no analysis executed.")
        return 0

    return int(analysis.run())


if __name__ == "__main__":
    raise SystemExit(main())
