from __future__ import annotations

import codecs
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, NoReturn


APPLESCRIPT_LONG_TIMEOUT_SECONDS = 7200
KEYNOTE_BUILD_BATCH_SIZE = 20

DEFAULT_TEXTBOX_STYLE = {
    "font": "HelveticaNeue",
    "fontSize": 50,
    "color": [0, 0, 0],
}

FIELD_SEP = chr(31)

IMAGE_KEYS = {"file", "position", "size"}
TEXT_BOX_KEYS = {"text", "position", "size", "font", "fontSize", "font_size", "color"}
OVERRIDE_KEYS = {
    "target",
    "text",
    "position",
    "size",
    "font",
    "fontSize",
    "font_size",
    "color",
    "opacity",
    "rotation",
}


class KeynoteCLIError(Exception):
    pass


def fail(message: str) -> NoReturn:
    raise KeynoteCLIError(message)


def ensure_runtime_available() -> None:
    if sys.platform != "darwin":
        fail("keynote-cli only works on macOS")
    if shutil.which("osascript") is None:
        fail("osascript was not found. Install macOS scripting tools or ensure osascript is on PATH.")
    if not Path("/Applications/Keynote.app").exists():
        fail("Keynote.app was not found in /Applications. Install Keynote before using keynote-cli.")


def ensure_allowed_keys(obj: dict[str, Any], allowed: set[str], field_name: str) -> None:
    unknown = sorted(set(obj) - allowed)
    if unknown:
        fail(f"{field_name} contains unknown field(s): {', '.join(repr(key) for key in unknown)}")


def ensure_required_keys(obj: dict[str, Any], required: set[str], field_name: str) -> None:
    missing = sorted(required - set(obj))
    if missing:
        fail(f"{field_name} is missing required field(s): {', '.join(repr(key) for key in missing)}")


def resolve_path(path_value: str, base_dir: Path) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def remove_path(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def ensure_existing_file(path: Path, field_name: str, suffix: str | None = None) -> None:
    if not path.exists():
        fail(f"{field_name} does not exist: {path}")
    if not path.is_file():
        fail(f"{field_name} is not a file: {path}")
    if suffix is not None and path.suffix.lower() != suffix.lower():
        fail(f"{field_name} must end with {suffix}: {path}")


def ensure_output_suffix(path: Path, suffix: str, field_name: str) -> None:
    if path.suffix.lower() != suffix.lower():
        fail(f"{field_name} must end with {suffix}: {path}")


def ensure_non_empty_string(value: Any, field_name: str) -> str:
    if not isinstance(value, str):
        fail(f"{field_name} must be a string")
    if not value.strip():
        fail(f"{field_name} must not be empty")
    return value


def ensure_number(value: Any, field_name: str, *, minimum: float | None = None, maximum: float | None = None) -> float:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        fail(f"{field_name} must be numeric")
    numeric = float(value)
    if minimum is not None and numeric < minimum:
        fail(f"{field_name} must be >= {minimum}")
    if maximum is not None and numeric > maximum:
        fail(f"{field_name} must be <= {maximum}")
    return numeric


def numeric_literal(value: Any) -> str:
    if isinstance(value, bool):
        fail("Boolean values are not valid numeric literals")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    fail(f"Expected a number, got {type(value).__name__}")


def applescript_string(value: Any) -> str:
    if value is None:
        value = ""
    if not isinstance(value, str):
        value = str(value)
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    parts = value.split("\n")
    escaped_parts = []
    for part in parts:
        escaped = part.replace("\\", "\\\\").replace('"', '\\"')
        escaped_parts.append(f'"{escaped}"')
    return " & linefeed & ".join(escaped_parts)


def applescript_posix_file(path: Path) -> str:
    return f"POSIX file {applescript_string(str(path))}"


def normalize_color(color: Any, field_name: str = "color") -> list[int]:
    if not isinstance(color, list) or len(color) != 3:
        fail(f"{field_name} must be an array of 3 numbers")
    components: list[int] = []
    for component in color:
        if not isinstance(component, (int, float)) or isinstance(component, bool):
            fail(f"{field_name} components must be numbers")
        components.append(int(round(component)))
    if max(components) <= 255:
        components = [component * 257 for component in components]
    for component in components:
        if component < 0 or component > 65535:
            fail(f"{field_name} components must be between 0 and 255, or 0 and 65535")
    return components


def applescript_color(color: Any) -> str:
    r, g, b = normalize_color(color)
    return "{" + f"{r}, {g}, {b}" + "}"


def validate_point(value: Any, field_name: str) -> list[float]:
    if not isinstance(value, list) or len(value) != 2:
        fail(f"{field_name} must be an array of 2 numbers")
    out: list[float] = []
    for component in value:
        if not isinstance(component, (int, float)) or isinstance(component, bool):
            fail(f"{field_name} must be an array of 2 numbers")
        out.append(float(component))
    return out


def validate_size(value: Any, field_name: str) -> list[float]:
    size = validate_point(value, field_name)
    if size[0] <= 0 or size[1] <= 0:
        fail(f"{field_name} values must both be > 0")
    return size


def validate_target(target: str, field_name: str) -> str:
    if not isinstance(target, str) or not target:
        fail(f"{field_name} must be a non-empty string")

    if target in ("defaultTitleItem", "defaultBodyItem"):
        return target

    for prefix, label in (("textItem:", "text item"), ("image:", "image"), ("shape:", "shape")):
        if target.startswith(prefix):
            try:
                index = int(target.split(":", 1)[1])
            except ValueError:
                fail(f"{field_name} has invalid {label} index: {target!r}")
            if index < 1:
                fail(f"{field_name} has invalid {label} index: {target!r}")
            return target

    fail(
        f"{field_name} has invalid target {target!r}. Use one of: defaultTitleItem, defaultBodyItem, "
        f"textItem:<n>, image:<n>, shape:<n>"
    )


def validate_text_target(target: str, field_name: str) -> str:
    validate_target(target, field_name)
    if target.startswith("image:"):
        fail(f"{field_name} target {target!r} is not valid for text operations")
    return target


def target_to_applescript(target: str, slide_var: str = "newSlide") -> str:
    if target == "defaultTitleItem":
        return f"default title item of {slide_var}"
    if target == "defaultBodyItem":
        return f"default body item of {slide_var}"
    if target.startswith("textItem:"):
        index = int(target.split(":", 1)[1])
        return f"text item {index} of {slide_var}"
    if target.startswith("image:"):
        index = int(target.split(":", 1)[1])
        return f"image {index} of {slide_var}"
    if target.startswith("shape:"):
        index = int(target.split(":", 1)[1])
        return f"shape {index} of {slide_var}"
    fail(f"Invalid target: {target}")


def run_osascript(script: str) -> str:
    result = subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
        timeout=APPLESCRIPT_LONG_TIMEOUT_SECONDS + 60,
    )
    if result.returncode != 0:
        error = result.stderr.strip()
        raise RuntimeError(f"osascript failed: {error}")
    return result.stdout


def decode_escaped(value: str) -> str:
    try:
        return codecs.decode(value, "unicode_escape")
    except Exception:
        return value
