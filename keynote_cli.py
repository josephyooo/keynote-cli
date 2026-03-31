#!/usr/bin/env python3
from __future__ import annotations

import argparse
import codecs
import json
from pathlib import Path
import shlex
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


# ---------------------------------------------------------------------------
# Script parser
# ---------------------------------------------------------------------------

def parse_pair(raw: str) -> list[float]:
    parts = raw.split(",")
    if len(parts) != 2:
        fail(f"Expected X,Y pair, got {raw!r}")
    try:
        return [float(parts[0]), float(parts[1])]
    except ValueError:
        fail(f"Invalid numeric pair: {raw!r}")


def parse_indents(raw: str) -> list[int]:
    parts = raw.split(",")
    try:
        indents = [int(p) for p in parts]
    except ValueError:
        fail(f"Invalid indents: {raw!r}")
    for i in indents:
        if i < 0:
            fail(f"Indent values must be non-negative, got {i}")
    return indents


def parse_color(raw: str) -> list[int]:
    parts = raw.split(",")
    if len(parts) != 3:
        fail(f"Color must be R,G,B — got {raw!r}")
    try:
        return normalize_color([int(p) for p in parts], "color")
    except ValueError:
        fail(f"Invalid color: {raw!r}")


def _unescape_script_text(text: str) -> str:
    return text.replace("\\n", "\n").replace("\\t", "\t").replace("\\\\", "\\")


def parse_script_line(line: str, line_num: int, base_dir: Path) -> dict[str, Any] | None:
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return None

    try:
        tokens = shlex.split(stripped)
    except ValueError as exc:
        fail(f"Line {line_num}: {exc}")

    if not tokens:
        return None

    cmd = tokens[0]

    if cmd == "open":
        parser = argparse.ArgumentParser(prog="open", exit_on_error=False)
        parser.add_argument("template")
        parser.add_argument("--output", required=True)
        parser.add_argument("--force", action="store_true")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError) as exc:
            fail(f"Line {line_num}: open requires TEMPLATE --output OUTPUT [--force]")
        template = resolve_path(args.template, base_dir)
        output = resolve_path(args.output, base_dir)
        ensure_existing_file(template, f"line {line_num} template", ".key")
        ensure_output_suffix(output, ".key", f"line {line_num} output")
        if template == output:
            fail(f"Line {line_num}: output must be different from template")
        return {"op": "open", "template": template, "output": output, "force": args.force}

    if cmd == "add-slide":
        parser = argparse.ArgumentParser(prog="add-slide", exit_on_error=False)
        parser.add_argument("--master", required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-slide requires --master NAME")
        return {"op": "add-slide", "master": args.master}

    if cmd == "set-text":
        parser = argparse.ArgumentParser(prog="set-text", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--target", required=True)
        parser.add_argument("--indents")
        parser.add_argument("text")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-text requires --slide N --target TARGET TEXT")
        validate_target(args.target, f"line {line_num} target")
        text = _unescape_script_text(args.text)
        indents = None
        if args.indents:
            indents = parse_indents(args.indents)
            line_count = text.count("\n") + 1 if text else 0
            if len(indents) != line_count:
                fail(f"Line {line_num}: indents count ({len(indents)}) must match text line count ({line_count})")
        return {"op": "set-text", "slide": args.slide, "target": args.target, "text": text, "indents": indents}

    if cmd == "set-notes":
        parser = argparse.ArgumentParser(prog="set-notes", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("text")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-notes requires --slide N TEXT")
        return {"op": "set-notes", "slide": args.slide, "text": _unescape_script_text(args.text)}

    if cmd == "add-image":
        parser = argparse.ArgumentParser(prog="add-image", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--file", required=True)
        parser.add_argument("--position", required=True)
        parser.add_argument("--size")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-image requires --slide N --file PATH --position X,Y [--size W,H]")
        file_path = resolve_path(args.file, base_dir)
        ensure_existing_file(file_path, f"line {line_num} image file")
        position = parse_pair(args.position)
        size = None
        if args.size:
            size = parse_pair(args.size)
            if size[0] <= 0 or size[1] <= 0:
                fail(f"Line {line_num}: image size values must both be > 0")
        return {"op": "add-image", "slide": args.slide, "file": file_path, "position": position, "size": size}

    if cmd == "add-text-box":
        parser = argparse.ArgumentParser(prog="add-text-box", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--text", required=True)
        parser.add_argument("--position", required=True)
        parser.add_argument("--size", required=True)
        parser.add_argument("--font")
        parser.add_argument("--font-size", type=float)
        parser.add_argument("--color")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-text-box requires --slide N --text TEXT --position X,Y --size W,H")
        position = parse_pair(args.position)
        size = parse_pair(args.size)
        if size[0] <= 0 or size[1] <= 0:
            fail(f"Line {line_num}: text box size values must both be > 0")
        font = args.font or DEFAULT_TEXTBOX_STYLE["font"]
        font_size = args.font_size or DEFAULT_TEXTBOX_STYLE["fontSize"]
        if font_size <= 0:
            fail(f"Line {line_num}: font-size must be > 0")
        color = parse_color(args.color) if args.color else normalize_color(DEFAULT_TEXTBOX_STYLE["color"])
        return {
            "op": "add-text-box", "slide": args.slide,
            "text": _unescape_script_text(args.text),
            "position": position, "size": size,
            "font": font, "fontSize": font_size, "color": color,
        }

    if cmd == "override":
        parser = argparse.ArgumentParser(prog="override", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--target", required=True)
        parser.add_argument("--text")
        parser.add_argument("--position")
        parser.add_argument("--size")
        parser.add_argument("--font")
        parser.add_argument("--font-size", type=float)
        parser.add_argument("--color")
        parser.add_argument("--opacity", type=float)
        parser.add_argument("--rotation", type=float)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: override requires --slide N --target TARGET [properties]")
        validate_target(args.target, f"line {line_num} target")
        if args.target.startswith("image:"):
            if any(getattr(args, f) is not None for f in ("text", "font", "font_size", "color")):
                fail(f"Line {line_num}: image targets cannot set text/font/color fields")
        override: dict[str, Any] = {"op": "override", "slide": args.slide, "target": args.target}
        has_property = False
        if args.text is not None:
            override["text"] = _unescape_script_text(args.text)
            has_property = True
        if args.position:
            override["position"] = parse_pair(args.position)
            has_property = True
        if args.size:
            s = parse_pair(args.size)
            if s[0] <= 0 or s[1] <= 0:
                fail(f"Line {line_num}: size values must both be > 0")
            override["size"] = s
            has_property = True
        if args.font:
            override["font"] = args.font
            has_property = True
        if args.font_size is not None:
            if args.font_size <= 0:
                fail(f"Line {line_num}: font-size must be > 0")
            override["fontSize"] = args.font_size
            has_property = True
        if args.color:
            override["color"] = parse_color(args.color)
            has_property = True
        if args.opacity is not None:
            if args.opacity < 0 or args.opacity > 100:
                fail(f"Line {line_num}: opacity must be between 0 and 100")
            override["opacity"] = args.opacity
            has_property = True
        if args.rotation is not None:
            override["rotation"] = args.rotation
            has_property = True
        if not has_property:
            fail(f"Line {line_num}: override must include at least one property to change")
        return override

    if cmd == "duplicate-slide":
        parser = argparse.ArgumentParser(prog="duplicate-slide", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--to", type=int, dest="to_pos")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: duplicate-slide requires --slide N [--to M]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.to_pos is not None and args.to_pos < 0:
            fail(f"Line {line_num}: --to must be >= 0")
        return {"op": "duplicate-slide", "slide": args.slide, "to": args.to_pos}

    if cmd == "move-slide":
        parser = argparse.ArgumentParser(prog="move-slide", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--to", type=int, required=True, dest="to_pos")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: move-slide requires --slide N --to M")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.to_pos < 1:
            fail(f"Line {line_num}: --to must be >= 1")
        return {"op": "move-slide", "slide": args.slide, "to": args.to_pos}

    if cmd == "replace-text":
        parser = argparse.ArgumentParser(prog="replace-text", exit_on_error=False)
        parser.add_argument("--find", required=True)
        parser.add_argument("--replace", required=True, dest="replace_with")
        parser.add_argument("--slide", type=int)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: replace-text requires --find TEXT --replace TEXT [--slide N]")
        if args.slide is not None and args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        return {"op": "replace-text", "find": args.find, "replace": args.replace_with, "slide": args.slide}

    if cmd == "add-shape":
        parser = argparse.ArgumentParser(prog="add-shape", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--position", required=True)
        parser.add_argument("--size", required=True)
        parser.add_argument("--text")
        parser.add_argument("--rotation", type=float)
        parser.add_argument("--opacity", type=float)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-shape requires --slide N --position X,Y --size W,H")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        position = parse_pair(args.position)
        size = parse_pair(args.size)
        if size[0] <= 0 or size[1] <= 0:
            fail(f"Line {line_num}: shape size values must both be > 0")
        if args.opacity is not None and (args.opacity < 0 or args.opacity > 100):
            fail(f"Line {line_num}: opacity must be between 0 and 100")
        result: dict[str, Any] = {"op": "add-shape", "slide": args.slide, "position": position, "size": size}
        if args.text is not None:
            result["text"] = _unescape_script_text(args.text)
        if args.rotation is not None:
            result["rotation"] = args.rotation
        if args.opacity is not None:
            result["opacity"] = args.opacity
        return result

    if cmd == "set-master":
        parser = argparse.ArgumentParser(prog="set-master", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--master", required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-master requires --slide N --master NAME")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        return {"op": "set-master", "slide": args.slide, "master": args.master}

    if cmd == "set-theme":
        parser = argparse.ArgumentParser(prog="set-theme", exit_on_error=False)
        parser.add_argument("--theme", required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-theme requires --theme NAME")
        return {"op": "set-theme", "theme": args.theme}

    if cmd == "skip-slide":
        parser = argparse.ArgumentParser(prog="skip-slide", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: skip-slide requires --slide N")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        return {"op": "skip-slide", "slide": args.slide}

    if cmd == "unskip-slide":
        parser = argparse.ArgumentParser(prog="unskip-slide", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: unskip-slide requires --slide N")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        return {"op": "unskip-slide", "slide": args.slide}

    if cmd == "delete-shape":
        parser = argparse.ArgumentParser(prog="delete-shape", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--index", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: delete-shape requires --slide N --index I")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.index < 1:
            fail(f"Line {line_num}: index must be >= 1")
        return {"op": "delete-shape", "slide": args.slide, "index": args.index}

    if cmd == "delete-image":
        parser = argparse.ArgumentParser(prog="delete-image", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--index", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: delete-image requires --slide N --index I")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.index < 1:
            fail(f"Line {line_num}: index must be >= 1")
        return {"op": "delete-image", "slide": args.slide, "index": args.index}

    if cmd == "add-line":
        parser = argparse.ArgumentParser(prog="add-line", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--from", required=True, dest="from_point")
        parser.add_argument("--to", required=True, dest="to_point")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-line requires --slide N --from X,Y --to X,Y")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        from_pt = parse_pair(args.from_point)
        to_pt = parse_pair(args.to_point)
        return {"op": "add-line", "slide": args.slide, "from": from_pt, "to": to_pt}

    if cmd == "duplicate-shape":
        parser = argparse.ArgumentParser(prog="duplicate-shape", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--index", type=int, required=True)
        parser.add_argument("--to-slide", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: duplicate-shape requires --slide N --index I --to-slide M")
        if args.slide < 1 or args.to_slide < 1:
            fail(f"Line {line_num}: slide numbers must be >= 1")
        if args.index < 1:
            fail(f"Line {line_num}: index must be >= 1")
        return {"op": "duplicate-shape", "slide": args.slide, "index": args.index, "to_slide": args.to_slide}

    if cmd == "set-style":
        parser = argparse.ArgumentParser(prog="set-style", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--target", required=True)
        parser.add_argument("--bold", action="store_true", default=None)
        parser.add_argument("--no-bold", action="store_true", default=None)
        parser.add_argument("--italic", action="store_true", default=None)
        parser.add_argument("--no-italic", action="store_true", default=None)
        parser.add_argument("--underline", action="store_true", default=None)
        parser.add_argument("--no-underline", action="store_true", default=None)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-style requires --slide N --target TARGET [--bold|--no-bold] [--italic|--no-italic] [--underline|--no-underline]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        validate_target(args.target, f"line {line_num} target")
        style: dict[str, Any] = {"op": "set-style", "slide": args.slide, "target": args.target}
        has_flag = False
        if args.bold:
            style["bold"] = True; has_flag = True
        elif args.no_bold:
            style["bold"] = False; has_flag = True
        if args.italic:
            style["italic"] = True; has_flag = True
        elif args.no_italic:
            style["italic"] = False; has_flag = True
        if args.underline:
            style["underline"] = True; has_flag = True
        elif args.no_underline:
            style["underline"] = False; has_flag = True
        if not has_flag:
            fail(f"Line {line_num}: set-style must include at least one style flag")
        return style

    if cmd == "add-table":
        parser = argparse.ArgumentParser(prog="add-table", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--rows", type=int, required=True)
        parser.add_argument("--cols", type=int, required=True)
        parser.add_argument("--position")
        parser.add_argument("--size")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-table requires --slide N --rows R --cols C [--position X,Y] [--size W,H]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.rows < 1 or args.cols < 1:
            fail(f"Line {line_num}: rows and cols must be >= 1")
        result: dict[str, Any] = {"op": "add-table", "slide": args.slide, "rows": args.rows, "cols": args.cols}
        if args.position:
            result["position"] = parse_pair(args.position)
        if args.size:
            s = parse_pair(args.size)
            if s[0] <= 0 or s[1] <= 0:
                fail(f"Line {line_num}: table size values must both be > 0")
            result["size"] = s
        return result

    if cmd == "set-cell":
        parser = argparse.ArgumentParser(prog="set-cell", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--table", type=int, default=1)
        parser.add_argument("--row", type=int, required=True)
        parser.add_argument("--col", type=int, required=True)
        parser.add_argument("value")
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-cell requires --slide N --row R --col C VALUE [--table I]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.table < 1:
            fail(f"Line {line_num}: table index must be >= 1")
        if args.row < 1 or args.col < 1:
            fail(f"Line {line_num}: row and col must be >= 1")
        return {"op": "set-cell", "slide": args.slide, "table": args.table, "row": args.row, "col": args.col, "value": _unescape_script_text(args.value)}

    if cmd == "add-row":
        parser = argparse.ArgumentParser(prog="add-row", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--table", type=int, default=1)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-row requires --slide N [--table I]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.table < 1:
            fail(f"Line {line_num}: table index must be >= 1")
        return {"op": "add-row", "slide": args.slide, "table": args.table}

    if cmd == "add-col":
        parser = argparse.ArgumentParser(prog="add-col", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--table", type=int, default=1)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: add-col requires --slide N [--table I]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.table < 1:
            fail(f"Line {line_num}: table index must be >= 1")
        return {"op": "add-col", "slide": args.slide, "table": args.table}

    if cmd == "delete-row":
        parser = argparse.ArgumentParser(prog="delete-row", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--table", type=int, default=1)
        parser.add_argument("--row", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: delete-row requires --slide N --row R [--table I]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.table < 1:
            fail(f"Line {line_num}: table index must be >= 1")
        if args.row < 1:
            fail(f"Line {line_num}: row must be >= 1")
        return {"op": "delete-row", "slide": args.slide, "table": args.table, "row": args.row}

    if cmd == "delete-col":
        parser = argparse.ArgumentParser(prog="delete-col", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--table", type=int, default=1)
        parser.add_argument("--col", type=int, required=True)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: delete-col requires --slide N --col C [--table I]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.table < 1:
            fail(f"Line {line_num}: table index must be >= 1")
        if args.col < 1:
            fail(f"Line {line_num}: col must be >= 1")
        return {"op": "delete-col", "slide": args.slide, "table": args.table, "col": args.col}

    if cmd == "set-transition":
        parser = argparse.ArgumentParser(prog="set-transition", exit_on_error=False)
        parser.add_argument("--slide", type=int, required=True)
        parser.add_argument("--style", required=True)
        parser.add_argument("--duration", type=float)
        try:
            args = parser.parse_args(tokens[1:])
        except (SystemExit, argparse.ArgumentError):
            fail(f"Line {line_num}: set-transition requires --slide N --style TYPE [--duration S]")
        if args.slide < 1:
            fail(f"Line {line_num}: slide number must be >= 1")
        if args.duration is not None and args.duration < 0:
            fail(f"Line {line_num}: duration must be >= 0")
        result2: dict[str, Any] = {"op": "set-transition", "slide": args.slide, "style": args.style}
        if args.duration is not None:
            result2["duration"] = args.duration
        return result2

    if cmd == "delete-slides":
        if len(tokens) != 2:
            fail(f"Line {line_num}: delete-slides requires a range (e.g. 1-7 or 5)")
        range_str = tokens[1]
        if "-" in range_str:
            parts = range_str.split("-", 1)
            try:
                start, end = int(parts[0]), int(parts[1])
            except ValueError:
                fail(f"Line {line_num}: invalid slide range: {range_str!r}")
            if start < 1 or end < start:
                fail(f"Line {line_num}: invalid slide range: {range_str!r}")
        else:
            try:
                start = end = int(range_str)
            except ValueError:
                fail(f"Line {line_num}: invalid slide number: {range_str!r}")
            if start < 1:
                fail(f"Line {line_num}: invalid slide number: {range_str!r}")
        return {"op": "delete-slides", "start": start, "end": end}

    if cmd == "save":
        if len(tokens) != 1:
            fail(f"Line {line_num}: save takes no arguments")
        return {"op": "save"}

    fail(f"Line {line_num}: unknown command {cmd!r}")


def parse_script(script_path: Path) -> list[dict[str, Any]]:
    base_dir = script_path.parent.resolve()
    lines = script_path.read_text(encoding="utf-8").splitlines()
    operations: list[dict[str, Any]] = []
    for line_num, line in enumerate(lines, start=1):
        op = parse_script_line(line, line_num, base_dir)
        if op is not None:
            operations.append(op)
    if not operations:
        fail("Script is empty")
    return operations


# ---------------------------------------------------------------------------
# Script to AppleScript compilation
# ---------------------------------------------------------------------------

DOC_OPS = frozenset({
    "duplicate-slide", "move-slide", "replace-text", "add-shape", "set-master", "set-theme",
    "skip-slide", "unskip-slide", "delete-shape", "delete-image", "add-line",
    "duplicate-shape", "set-style", "add-table", "set-cell", "add-row", "add-col",
    "delete-row", "delete-col", "set-transition", "delete-slides",
})


def _group_operations_into_slides(operations: list[dict[str, Any]]) -> dict[str, Any]:
    """Group parsed operations into a structured build plan.

    Returns a dict with:
      template, output, force, slides (list of slide dicts), doc_ops, has_save
    Each slide dict has: master, content, images, text_boxes, overrides, notes.
    doc_ops are document-level operations executed after slide creation, in script order.
    """
    open_op = None
    slides: list[dict[str, Any]] = []
    doc_ops: list[dict[str, Any]] = []
    has_save = False
    # Track per-new-slide operations that reference slide numbers
    deferred_ops: list[dict[str, Any]] = []

    for op in operations:
        if op["op"] == "open":
            if open_op is not None:
                fail("Script contains multiple 'open' commands")
            open_op = op
        elif op["op"] == "add-slide":
            slides.append({
                "master": op["master"],
                "content": [],
                "images": [],
                "text_boxes": [],
                "overrides": [],
                "notes": None,
            })
        elif op["op"] in ("set-text", "set-notes", "add-image", "add-text-box", "override"):
            deferred_ops.append(op)
        elif op["op"] in DOC_OPS:
            doc_ops.append(op)
        elif op["op"] == "save":
            has_save = True

    if open_op is None:
        fail("Script must contain an 'open' command")

    # Resolve deferred operations to their slides
    for op in deferred_ops:
        slide_idx = op["slide"] - 1  # 1-based to 0-based
        if slide_idx < 0 or slide_idx >= len(slides):
            fail(f"Slide {op['slide']} is out of range (have {len(slides)} new slides)")
        slide = slides[slide_idx]

        if op["op"] == "set-text":
            slide["content"].append({
                "target": op["target"],
                "text": op["text"],
                "indents": op["indents"],
            })
        elif op["op"] == "set-notes":
            slide["notes"] = op["text"]
        elif op["op"] == "add-image":
            slide["images"].append({
                "file": op["file"],
                "position": op["position"],
                "size": op["size"],
            })
        elif op["op"] == "add-text-box":
            slide["text_boxes"].append({
                "text": op["text"],
                "position": op["position"],
                "size": op["size"],
                "font": op["font"],
                "fontSize": op["fontSize"],
                "color": op["color"],
            })
        elif op["op"] == "override":
            override_dict: dict[str, Any] = {"target": op["target"]}
            for key in ("text", "position", "size", "font", "fontSize", "color", "opacity", "rotation"):
                if key in op:
                    override_dict[key] = op[key]
            slide["overrides"].append(override_dict)

    return {
        "template": open_op["template"],
        "output": open_op["output"],
        "force": open_op["force"],
        "slides": slides,
        "doc_ops": doc_ops,
        "has_save": has_save,
    }


# ---------------------------------------------------------------------------
# AppleScript generation
# ---------------------------------------------------------------------------

def build_slide_applescript(slide: dict[str, Any], slide_number: int) -> list[str]:
    master = slide["master"]
    body_lines: list[str] = []
    body_lines.append(
        f"set newSlide to make new slide with properties {{base slide: master slide {applescript_string(master)}}}"
    )

    for content_item in slide["content"]:
        target = content_item["target"]
        expr = target_to_applescript(target)
        text = content_item["text"]
        body_lines.append(f"set object text of {expr} to {applescript_string(text)}")
        indents = content_item.get("indents")
        if indents:
            body_lines.append(f"tell object text of {expr}")
            for para_idx, indent_lvl in enumerate(indents, start=1):
                body_lines.append(f"  set indent level of paragraph {para_idx} to {indent_lvl}")
            body_lines.append("end tell")

    if slide.get("notes") is not None:
        body_lines.append(f"set presenter notes of newSlide to {applescript_string(slide['notes'])}")

    for image in slide["images"]:
        properties = [
            f"file:{applescript_posix_file(image['file'])}",
            f"position:{{{numeric_literal(image['position'][0])}, {numeric_literal(image['position'][1])}}}",
        ]
        if image["size"] is not None:
            properties.append(f"width:{numeric_literal(image['size'][0])}")
            properties.append(f"height:{numeric_literal(image['size'][1])}")
        body_lines.append(f"tell newSlide to make new image with properties {{{', '.join(properties)}}}")

    for box in slide["text_boxes"]:
        properties = [
            f"object text:{applescript_string(box['text'])}",
            f"position:{{{numeric_literal(box['position'][0])}, {numeric_literal(box['position'][1])}}}",
            f"width:{numeric_literal(box['size'][0])}",
            f"height:{numeric_literal(box['size'][1])}",
        ]
        body_lines.append("tell newSlide")
        body_lines.append(f"  set newTextBox to make new text item with properties {{{', '.join(properties)}}}")
        body_lines.append(f"  set font of object text of newTextBox to {applescript_string(box['font'])}")
        body_lines.append(f"  set size of object text of newTextBox to {numeric_literal(box['fontSize'])}")
        body_lines.append(f"  set color of object text of newTextBox to {{{box['color'][0]}, {box['color'][1]}, {box['color'][2]}}}")
        body_lines.append("end tell")

    for override in slide["overrides"]:
        expr = target_to_applescript(override["target"])
        if "text" in override:
            body_lines.append(f"set object text of {expr} to {applescript_string(override['text'])}")
        if "position" in override:
            body_lines.append(
                f"set position of {expr} to {{{numeric_literal(override['position'][0])}, {numeric_literal(override['position'][1])}}}"
            )
        if "size" in override:
            body_lines.append(f"set width of {expr} to {numeric_literal(override['size'][0])}")
            body_lines.append(f"set height of {expr} to {numeric_literal(override['size'][1])}")
        if "font" in override:
            body_lines.append(f"set font of object text of {expr} to {applescript_string(override['font'])}")
        if "fontSize" in override:
            body_lines.append(f"set size of object text of {expr} to {numeric_literal(override['fontSize'])}")
        if "color" in override:
            body_lines.append(
                f"set color of object text of {expr} to {{{override['color'][0]}, {override['color'][1]}, {override['color'][2]}}}"
            )
        if "opacity" in override:
            body_lines.append(f"set opacity of {expr} to {numeric_literal(override['opacity'])}")
        if "rotation" in override:
            body_lines.append(f"set rotation of {expr} to {numeric_literal(override['rotation'])}")

    lines: list[str] = [
        f"    -- Slide {slide_number}: {master}",
        "    try",
    ]
    lines.extend(f"      {line}" for line in body_lines)
    lines.extend(
        [
            "    on error errMsg number errNum",
            f"      error {applescript_string(f'Slide {slide_number} ({master}) failed: ')} & errMsg number errNum",
            "    end try",
            "",
        ]
    )
    return lines


def _build_doc_op_applescript(op: dict[str, Any]) -> list[str]:
    """Generate AppleScript lines for a document-level operation."""
    lines: list[str] = []
    kind = op["op"]

    if kind == "duplicate-slide":
        src = op["slide"]
        to = op.get("to")
        if to is not None:
            lines.append(f"      duplicate slide {src} to after slide {to}")
        else:
            lines.append(f"      duplicate slide {src}")

    elif kind == "move-slide":
        src = op["slide"]
        to = op["to"]
        if to == 1:
            lines.append(f"      move slide {src} to beginning")
        else:
            lines.append(f"      move slide {src} to after slide {to - 1}")

    elif kind == "replace-text":
        find_as = applescript_string(op["find"])
        replace_as = applescript_string(op["replace"])
        slide_num = op.get("slide")
        if slide_num is not None:
            lines.append(f"      tell slide {slide_num}")
            lines.append(f"        repeat with ti in text items")
            lines.append(f"          set t to object text of ti")
            lines.append(f"          if t contains {find_as} then")
            lines.append(f"            set object text of ti to my replaceText(t, {find_as}, {replace_as})")
            lines.append(f"          end if")
            lines.append(f"        end repeat")
            lines.append(f"      end tell")
        else:
            lines.append(f"      repeat with s in slides")
            lines.append(f"        repeat with ti in text items of s")
            lines.append(f"          set t to object text of ti")
            lines.append(f"          if t contains {find_as} then")
            lines.append(f"            set object text of ti to my replaceText(t, {find_as}, {replace_as})")
            lines.append(f"          end if")
            lines.append(f"        end repeat")
            lines.append(f"      end repeat")

    elif kind == "add-shape":
        slide_num = op["slide"]
        pos = op["position"]
        sz = op["size"]
        lines.append(f"      tell slide {slide_num}")
        lines.append(f"        set newShape to make new shape")
        lines.append(f"        tell newShape")
        lines.append(f"          set position to {{{numeric_literal(pos[0])}, {numeric_literal(pos[1])}}}")
        lines.append(f"          set width to {numeric_literal(sz[0])}")
        lines.append(f"          set height to {numeric_literal(sz[1])}")
        if "text" in op:
            lines.append(f"          set object text to {applescript_string(op['text'])}")
        if "rotation" in op:
            lines.append(f"          set rotation to {numeric_literal(op['rotation'])}")
        if "opacity" in op:
            lines.append(f"          set opacity to {numeric_literal(op['opacity'])}")
        lines.append(f"        end tell")
        lines.append(f"      end tell")

    elif kind == "set-master":
        slide_num = op["slide"]
        master = op["master"]
        lines.append(f"      set base slide of slide {slide_num} to master slide {applescript_string(master)}")

    elif kind == "set-theme":
        theme = op["theme"]
        lines.append(f"      set document theme to theme {applescript_string(theme)}")

    elif kind == "skip-slide":
        lines.append(f"      set skipped of slide {op['slide']} to true")

    elif kind == "unskip-slide":
        lines.append(f"      set skipped of slide {op['slide']} to false")

    elif kind == "delete-shape":
        lines.append(f"      delete shape {op['index']} of slide {op['slide']}")

    elif kind == "delete-image":
        lines.append(f"      delete image {op['index']} of slide {op['slide']}")

    elif kind == "add-line":
        slide_num = op["slide"]
        from_pt = op["from"]
        to_pt = op["to"]
        lines.append(f"      tell slide {slide_num}")
        lines.append(f"        make new line with properties {{start point:{{{numeric_literal(from_pt[0])}, {numeric_literal(from_pt[1])}}}, end point:{{{numeric_literal(to_pt[0])}, {numeric_literal(to_pt[1])}}}}}")
        lines.append(f"      end tell")

    elif kind == "duplicate-shape":
        lines.append(f"      duplicate shape {op['index']} of slide {op['slide']} to slide {op['to_slide']}")

    elif kind == "set-style":
        slide_num = op["slide"]
        target = op["target"]
        expr = target_to_applescript(target, f"slide {slide_num}")
        lines.append(f"      tell object text of {expr}")
        if "bold" in op:
            lines.append(f"        set bold to {'true' if op['bold'] else 'false'}")
        if "italic" in op:
            lines.append(f"        set italic to {'true' if op['italic'] else 'false'}")
        if "underline" in op:
            lines.append(f"        set underlined to {'true' if op['underline'] else 'false'}")
        lines.append(f"      end tell")

    elif kind == "add-table":
        slide_num = op["slide"]
        rows = op["rows"]
        cols = op["cols"]
        lines.append(f"      tell slide {slide_num}")
        lines.append(f"        set newTable to make new table with properties {{row count:{rows}, column count:{cols}}}")
        if "position" in op:
            pos = op["position"]
            lines.append(f"        set position of newTable to {{{numeric_literal(pos[0])}, {numeric_literal(pos[1])}}}")
        if "size" in op:
            sz = op["size"]
            lines.append(f"        set width of newTable to {numeric_literal(sz[0])}")
            lines.append(f"        set height of newTable to {numeric_literal(sz[1])}")
        lines.append(f"      end tell")

    elif kind == "set-cell":
        slide_num = op["slide"]
        table_idx = op["table"]
        row = op["row"]
        col = op["col"]
        value = op["value"]
        lines.append(f"      set value of cell {col} of row {row} of table {table_idx} of slide {slide_num} to {applescript_string(value)}")

    elif kind == "add-row":
        lines.append(f"      tell table {op['table']} of slide {op['slide']}")
        lines.append(f"        add row above row (row count)")
        lines.append(f"      end tell")

    elif kind == "add-col":
        lines.append(f"      tell table {op['table']} of slide {op['slide']}")
        lines.append(f"        add column above column (column count)")
        lines.append(f"      end tell")

    elif kind == "delete-row":
        lines.append(f"      tell table {op['table']} of slide {op['slide']}")
        lines.append(f"        delete row {op['row']}")
        lines.append(f"      end tell")

    elif kind == "delete-col":
        lines.append(f"      tell table {op['table']} of slide {op['slide']}")
        lines.append(f"        delete column {op['col']}")
        lines.append(f"      end tell")

    elif kind == "set-transition":
        slide_num = op["slide"]
        style = op["style"]
        lines.append(f"      tell transition settings of slide {slide_num}")
        lines.append(f"        set transition effect to {style}")
        if "duration" in op:
            lines.append(f"        set transition duration to {numeric_literal(op['duration'])}")
        lines.append(f"      end tell")

    elif kind == "delete-slides":
        start, end = op["start"], op["end"]
        lines.append(f"      repeat with i from {end} to {start} by -1")
        lines.append(f"        delete slide i")
        lines.append(f"      end repeat")

    return lines


def _needs_replace_text_helper(doc_ops: list[dict[str, Any]]) -> bool:
    return any(op["op"] == "replace-text" for op in doc_ops)


REPLACE_TEXT_HELPER = """\

on replaceText(theText, searchFor, replaceWith)
    set oldDelim to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchFor
    set theItems to text items of theText
    set AppleScript's text item delimiters to replaceWith
    set theText to theItems as text
    set AppleScript's text item delimiters to oldDelim
    return theText
end replaceText"""


def build_build_applescript(
    output_path: Path,
    slides: list[dict[str, Any]],
    *,
    start_slide_number: int = 1,
    doc_ops: list[dict[str, Any]] | None = None,
) -> str:
    lines: list[str] = [
        f"with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds",
        'tell application "Keynote"',
        "  set theDoc to missing value",
        f"  set outputFile to {applescript_posix_file(output_path)}",
        "  try",
        "    set theDoc to open outputFile",
        "    tell theDoc",
        "",
    ]

    for index, slide in enumerate(slides, start=start_slide_number):
        lines.extend(build_slide_applescript(slide, index))

    if doc_ops:
        lines.append("")
        lines.append("    -- Document-level operations")
        for op in doc_ops:
            lines.extend(_build_doc_op_applescript(op))

    lines.extend(
        [
            "      save",
            "    end tell",
            "    close theDoc saving yes",
            "  on error errMsg number errNum",
            "    try",
            "      if theDoc is not missing value then close theDoc saving no",
            "    end try",
            "    error \"Build failed: \" & errMsg number errNum",
            "  end try",
            "end tell",
            "end timeout",
        ]
    )

    if doc_ops and _needs_replace_text_helper(doc_ops):
        lines.append(REPLACE_TEXT_HELPER)

    return "\n".join(lines)


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


# ---------------------------------------------------------------------------
# Inspect
# ---------------------------------------------------------------------------

def build_inspect_applescript(file_path: Path) -> str:
    sep = FIELD_SEP
    return f"""\
with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds
tell application "Keynote"
    set theDoc to open {applescript_posix_file(file_path)}
    try
        tell theDoc
            set slideW to width
            set slideH to height
            set masterNames to {{}}
            repeat with ms in master slides
                set end of masterNames to name of ms
            end repeat
            set slideCount to count of slides
            set resultLines to {{}}
            set end of resultLines to ("DIMENSIONS{sep}" & slideW & "{sep}" & slideH)
            set end of resultLines to ("MASTERS{sep}" & my joinList(masterNames, "{sep}"))
            set end of resultLines to ("SLIDECOUNT{sep}" & slideCount)
            repeat with slideIndex from 1 to slideCount
                set theSlide to slide slideIndex
                set masterName to name of base slide of theSlide
                set notesText to presenter notes of theSlide
                set end of resultLines to ("SLIDE{sep}" & slideIndex & "{sep}" & masterName & "{sep}" & my escapeField(notesText))
                try
                    set tiCount to count of text items of theSlide
                    repeat with tiIndex from 1 to tiCount
                        set ti to text item tiIndex of theSlide
                        set tiText to object text of ti
                        set tiPos to position of ti
                        set tiW to width of ti
                        set tiH to height of ti
                        set end of resultLines to ("TEXTITEM{sep}" & slideIndex & "{sep}" & tiIndex & "{sep}" & my escapeField(tiText) & "{sep}" & (item 1 of tiPos) & "{sep}" & (item 2 of tiPos) & "{sep}" & tiW & "{sep}" & tiH)
                    end repeat
                end try
                try
                    set imgCount to count of images of theSlide
                    repeat with imgIndex from 1 to imgCount
                        set img to image imgIndex of theSlide
                        set imgPos to position of img
                        set imgW to width of img
                        set imgH to height of img
                        set end of resultLines to ("IMAGE{sep}" & slideIndex & "{sep}" & imgIndex & "{sep}" & (item 1 of imgPos) & "{sep}" & (item 2 of imgPos) & "{sep}" & imgW & "{sep}" & imgH)
                    end repeat
                end try
                try
                    set shapeCount to count of shapes of theSlide
                    repeat with shapeIndex from 1 to shapeCount
                        set sh to shape shapeIndex of theSlide
                        set shText to ""
                        try
                            set shText to object text of sh
                        end try
                        set shPos to position of sh
                        set shW to width of sh
                        set shH to height of sh
                        set end of resultLines to ("SHAPE{sep}" & slideIndex & "{sep}" & shapeIndex & "{sep}" & my escapeField(shText) & "{sep}" & (item 1 of shPos) & "{sep}" & (item 2 of shPos) & "{sep}" & shW & "{sep}" & shH)
                    end repeat
                end try
            end repeat
        end tell
        close theDoc saving no
        set AppleScript's text item delimiters to linefeed
        return resultLines as text
    on error errMsg number errNum
        try
            close theDoc saving no
        end try
        error errMsg number errNum
    end try
end tell
end timeout

on escapeField(theText)
    set escaped to ""
    repeat with c in characters of theText
        set c to c as text
        if c is "\\" then
            set escaped to escaped & "\\\\\\\\"
        else if c is return then
            set escaped to escaped & "\\\\n"
        else if c is linefeed then
            set escaped to escaped & "\\\\n"
        else
            set escaped to escaped & c
        end if
    end repeat
    return escaped
end escapeField

on joinList(theList, delim)
    set oldDelim to AppleScript's text item delimiters
    set AppleScript's text item delimiters to delim
    set joined to theList as text
    set AppleScript's text item delimiters to oldDelim
    return joined
end joinList
"""


def filter_text_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen_positions: set[tuple[float, float, float, float]] = set()
    filtered: list[dict[str, Any]] = []
    for item in items:
        w = float(item["size"][0])
        h = float(item["size"][1])
        if w == 0 and h == 0:
            continue
        key = (float(item["position"][0]), float(item["position"][1]), w, h)
        if key in seen_positions:
            continue
        seen_positions.add(key)
        filtered.append(item)
    return filtered


def inspect_file(file_path: Path) -> dict[str, Any]:
    script = build_inspect_applescript(file_path)
    raw = run_osascript(script)

    dimensions: list[int] = [0, 0]
    masters: list[str] = []
    slide_count = 0
    slides_data: dict[int, dict[str, Any]] = {}

    for line in raw.splitlines():
        parts = line.split(FIELD_SEP)
        kind = parts[0]
        if kind == "DIMENSIONS":
            dimensions = [int(float(parts[1])), int(float(parts[2]))]
        elif kind == "MASTERS":
            masters = [p for p in parts[1:] if p]
        elif kind == "SLIDECOUNT":
            slide_count = int(parts[1])
        elif kind == "SLIDE":
            idx = int(parts[1])
            slides_data[idx] = {
                "index": idx,
                "master": parts[2],
                "notes": decode_escaped(parts[3]) if len(parts) > 3 else "",
                "textItems": [],
                "images": [],
                "shapes": [],
            }
        elif kind == "TEXTITEM":
            idx = int(parts[1])
            if idx not in slides_data:
                continue
            slides_data[idx]["textItems"].append({
                "index": int(parts[2]),
                "text": decode_escaped(parts[3]),
                "position": [float(parts[4]), float(parts[5])],
                "size": [float(parts[6]), float(parts[7])],
            })
        elif kind == "IMAGE":
            idx = int(parts[1])
            if idx not in slides_data:
                continue
            slides_data[idx]["images"].append({
                "index": int(parts[2]),
                "position": [float(parts[3]), float(parts[4])],
                "size": [float(parts[5]), float(parts[6])],
            })
        elif kind == "SHAPE":
            idx = int(parts[1])
            if idx not in slides_data:
                continue
            text = decode_escaped(parts[3]) if len(parts) > 3 else ""
            pos = [float(parts[4]), float(parts[5])]
            sz = [float(parts[6]), float(parts[7])]
            if sz[0] > 0 and sz[1] > 0:
                slides_data[idx]["shapes"].append({
                    "index": int(parts[2]),
                    "text": text,
                    "position": pos,
                    "size": sz,
                })

    ordered_slides: list[dict[str, Any]] = []
    for i in range(1, slide_count + 1):
        if i in slides_data:
            slide = slides_data[i]
            slide["textItems"] = filter_text_items(slide["textItems"])
            ordered_slides.append(slide)

    return {
        "file": str(file_path),
        "dimensions": dimensions,
        "slideCount": slide_count,
        "masters": masters,
        "slides": ordered_slides,
    }


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

EXPORT_FORMAT_MAP = {
    "pdf": ("PDF", ".pdf"),
    "png": ("slide images", ".png"),
    "jpeg": ("slide images", ".jpeg"),
    "pptx": ("Microsoft PowerPoint", ".pptx"),
    "html": ("HTML", ".html"),
    "movie": ("QuickTime movie", ".m4v"),
}


def build_export_applescript(input_path: Path, output_path: Path, *, export_format: str = "pdf") -> str:
    fmt_info = EXPORT_FORMAT_MAP.get(export_format)
    if fmt_info is None:
        fail(f"Unknown export format: {export_format!r}")
    as_format = fmt_info[0]

    image_extra = ""
    if export_format in ("png", "jpeg"):
        img_fmt = "PNG" if export_format == "png" else "JPEG"
        image_extra = f" with properties {{export style:IndividualSlides, image format:{img_fmt}}}"

    return f"""\
with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds
tell application "Keynote"
    set theDoc to open {applescript_posix_file(input_path)}
    try
        export theDoc to {applescript_posix_file(output_path)} as {as_format}{image_extra}
        close theDoc saving no
    on error errMsg number errNum
        try
            close theDoc saving no
        end try
        error errMsg number errNum
    end try
end tell
end timeout
"""


def validate_template_masters(template_path: Path, slides: list[dict[str, Any]]) -> dict[str, Any]:
    info = inspect_file(template_path)
    available = set(info.get("masters", []))
    required = {slide["master"] for slide in slides}
    missing = sorted(required - available)
    if missing:
        fail(
            f"Template is missing required master slide(s): {', '.join(repr(name) for name in missing)}. "
            f"Available: {', '.join(repr(name) for name in sorted(available))}"
        )
    return info


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def command_run(args: argparse.Namespace) -> int:
    script_path = Path(args.script).resolve()
    if not script_path.exists():
        fail(f"Script file not found: {script_path}")

    operations = parse_script(script_path)
    plan = _group_operations_into_slides(operations)

    output_path: Path = plan["output"]

    if args.print_applescript:
        script = build_build_applescript(output_path, plan["slides"], doc_ops=plan["doc_ops"])
        print(script)
        return 0

    ensure_runtime_available()

    if args.check_template:
        validate_template_masters(plan["template"], plan["slides"])

    if output_path.exists():
        if plan["force"] or args.force:
            remove_path(output_path)
        else:
            fail(f"Output already exists: {output_path} (use --force or add --force to open command)")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(plan["template"], output_path)

    removed_partial_output = False
    try:
        total_slides = len(plan["slides"])
        for batch_start in range(0, total_slides, KEYNOTE_BUILD_BATCH_SIZE):
            batch_slides = plan["slides"][batch_start: batch_start + KEYNOTE_BUILD_BATCH_SIZE]
            is_last_batch = batch_start + len(batch_slides) >= total_slides
            batch_doc_ops = plan["doc_ops"] if is_last_batch else None
            script = build_build_applescript(
                output_path,
                batch_slides,
                start_slide_number=batch_start + 1,
                doc_ops=batch_doc_ops,
            )
            run_osascript(script)
    except Exception as exc:
        if output_path.exists() and not args.keep_failed_output:
            remove_path(output_path)
            removed_partial_output = True
        message = f"Build failed for {output_path}: {exc}"
        if removed_partial_output:
            message += " Partial output was removed. Re-run with --keep-failed-output to inspect it."
        fail(message)

    print(str(output_path))
    return 0


def command_inspect(args: argparse.Namespace) -> int:
    file_path = Path(args.file).resolve()
    ensure_existing_file(file_path, "Input file", ".key")
    ensure_runtime_available()
    try:
        result = inspect_file(file_path)
    except Exception as exc:
        fail(f"Inspect failed for {file_path}: {exc}")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0


def command_export(args: argparse.Namespace) -> int:
    input_path = Path(args.file).resolve()
    ensure_existing_file(input_path, "Input file", ".key")
    ensure_runtime_available()

    export_format = args.format.lower()
    if export_format not in EXPORT_FORMAT_MAP:
        fail(f"Unknown format: {export_format!r}. Supported: {', '.join(sorted(EXPORT_FORMAT_MAP))}")
    default_suffix = EXPORT_FORMAT_MAP[export_format][1]

    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_path = input_path.with_suffix(default_suffix)

    if output_path.exists():
        if args.force:
            remove_path(output_path)
        else:
            fail(f"Output already exists: {output_path} (use --force to overwrite)")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        run_osascript(build_export_applescript(input_path, output_path, export_format=export_format))
    except Exception as exc:
        fail(f"Export failed for {input_path}: {exc}")
    print(str(output_path))
    return 0


# ---------------------------------------------------------------------------
# Insert equations
# ---------------------------------------------------------------------------

def _build_equation_insert_script(
    slide_num: int, placeholder: str, latex: str, render_timeout: int,
) -> str:
    ph_as = applescript_string(placeholder)
    latex_as = applescript_string(latex)
    return f"""\
tell application "Keynote"
    activate
    tell document 1
        set current slide to slide {slide_num}
    end tell
end tell

delay 0.3

tell application "System Events"
    tell process "Keynote"
        -- Deselect / return to slide level
        key code 53
        delay 0.2

        -- Open Find bar and search for the placeholder
        keystroke "f" using command down
        delay 0.4
        keystroke "a" using command down
        keystroke {ph_as}
        delay 0.2
        key code 36 -- Return: execute find (selects match)
        delay 0.2
        key code 53 -- Escape: close Find bar (selection stays)
        delay 0.2

        -- Open the equation editor
        click menu item "Equation..." of menu "Insert" of menu bar 1

        -- Wait for the equation sheet to appear
        repeat 15 times
            try
                sheet 1 of window 1
                exit repeat
            end try
            delay 0.3
        end repeat
        delay 0.3

        -- Type the LaTeX into the editor
        set theSheet to sheet 1 of window 1
        set textInput to text area 1 of scroll area 1 of splitter group 1 of theSheet
        set value of textInput to {latex_as}

        -- Wait for the renderer – Insert button is greyed-out until done
        repeat {render_timeout} times
            delay 1
            try
                if enabled of button "Insert" of theSheet then
                    click button "Insert" of theSheet
                    delay 0.3
                    return "ok"
                end if
            end try
        end repeat

        -- Timed out – cancel so the dialog doesn't stay open
        click button "Cancel" of theSheet
        return "render_timeout"
    end tell
end tell
"""


def command_insert_equations(args: argparse.Namespace) -> int:
    mappings_path = Path(args.mappings).resolve()
    if not mappings_path.exists():
        fail(f"Mappings file not found: {mappings_path}")
    with open(mappings_path) as f:
        data = json.load(f)
    entries: list[dict[str, Any]] = (
        data if isinstance(data, list) else data.get("equations", [])
    )
    if not entries:
        fail("No equations found in mappings file")
    for i, entry in enumerate(entries):
        for key in ("slide", "placeholder", "latex"):
            if key not in entry:
                fail(f"Entry {i}: missing required key '{key}'")
        if not isinstance(entry["slide"], int) or entry["slide"] < 1:
            fail(f"Entry {i}: 'slide' must be a positive integer")

    render_timeout: int = args.render_timeout
    dry_run: bool = args.dry_run

    succeeded = 0
    failed_entries: list[str] = []

    for i, entry in enumerate(entries):
        slide = entry["slide"]
        placeholder = entry["placeholder"]
        latex = entry["latex"]
        label = entry.get("label", placeholder)

        print(
            f"[{i + 1}/{len(entries)}] Slide {slide} {label} … ",
            end="",
            flush=True,
        )

        if dry_run:
            print("(dry run)")
            succeeded += 1
            continue

        script = _build_equation_insert_script(
            slide, placeholder, latex, render_timeout,
        )
        if args.print_applescript:
            print()
            print(script)
            succeeded += 1
            continue

        try:
            output = run_osascript(script).strip()
        except RuntimeError as exc:
            output = f"error: {exc}"

        if output == "ok":
            print("OK")
            succeeded += 1
        else:
            print(f"FAILED ({output})")
            failed_entries.append(label)

    total = len(entries)
    n_failed = len(failed_entries)
    print(f"\nDone: {succeeded}/{total} inserted, {n_failed} failed")
    if failed_entries:
        for label in failed_entries:
            print(f"  FAILED: {label}")
        return 1
    return 0


# ---------------------------------------------------------------------------
# Insert links
# ---------------------------------------------------------------------------

def _build_url_link_script(slide_num: int, find_text: str, url: str) -> str:
    """Build AppleScript that finds *find_text* on *slide_num* and adds a
    URL hyperlink to it via Cmd+K (GUI scripting)."""
    find_as = applescript_string(find_text)
    url_as = applescript_string(url)
    return f"""\
tell application "Keynote"
    activate
    tell document 1
        set current slide to slide {slide_num}
    end tell
end tell

delay 0.3

tell application "System Events"
    tell process "Keynote"
        -- Deselect / return to slide level
        key code 53
        delay 0.2

        -- Open Find bar and search for the target text
        keystroke "f" using command down
        delay 0.4
        keystroke "a" using command down
        keystroke {find_as}
        delay 0.2
        key code 36 -- Return: execute find (selects match)
        delay 0.2
        key code 53 -- Escape: close Find bar (selection stays)
        delay 0.3

        -- Open Add Link dialog (Cmd+K)
        keystroke "k" using command down
        delay 0.5

        -- Type the URL into the link field
        keystroke "a" using command down
        delay 0.1
        keystroke {url_as}
        delay 0.2

        -- Confirm
        key code 36 -- Return
        delay 0.2

        -- Deselect
        key code 53
        return "ok"
    end tell
end tell
"""


def command_insert_links(args: argparse.Namespace) -> int:
    """Add URL hyperlinks to text in an open Keynote deck.

    Requires: macOS Accessibility permissions for the calling process
    and an open Keynote document as front document.
    """
    mappings_path = Path(args.mappings).resolve()
    if not mappings_path.exists():
        fail(f"Mappings file not found: {mappings_path}")
    with open(mappings_path) as f:
        data = json.load(f)
    entries: list[dict[str, Any]] = (
        data if isinstance(data, list) else data.get("links", [])
    )
    if not entries:
        fail("No links found in mappings file")
    for i, entry in enumerate(entries):
        for key in ("slide", "find", "url"):
            if key not in entry:
                fail(f"Entry {i}: missing required key '{key}'")
        if not isinstance(entry["slide"], int) or entry["slide"] < 1:
            fail(f"Entry {i}: 'slide' must be a positive integer")

    dry_run: bool = args.dry_run
    succeeded = 0
    failed_entries: list[str] = []

    for i, entry in enumerate(entries):
        slide = entry["slide"]
        find_text = entry["find"]
        url = entry["url"]
        label = entry.get("label", find_text)

        print(
            f"[{i + 1}/{len(entries)}] Slide {slide} {label} … ",
            end="",
            flush=True,
        )

        if dry_run:
            print("(dry run)")
            succeeded += 1
            continue

        script = _build_url_link_script(slide, find_text, url)
        if args.print_applescript:
            print()
            print(script)
            succeeded += 1
            continue

        try:
            output = run_osascript(script).strip()
        except RuntimeError as exc:
            output = f"error: {exc}"

        if output == "ok":
            print("OK")
            succeeded += 1
        else:
            print(f"FAILED ({output})")
            failed_entries.append(label)

    total = len(entries)
    n_failed = len(failed_entries)
    print(f"\nDone: {succeeded}/{total} linked, {n_failed} failed")
    if failed_entries:
        for label in failed_entries:
            print(f"  FAILED: {label}")
        return 1
    return 0


def _build_slide_link_script(slide_num: int, shape_index: int, to_slide: int) -> str:
    """Build AppleScript that selects shape *shape_index* on *slide_num*
    and adds a slide navigation link to *to_slide* via the Add Link popover
    (GUI scripting)."""
    return f"""\
tell application "Keynote"
    activate
    tell document 1
        set current slide to slide {slide_num}
    end tell
end tell

delay 0.3

tell application "System Events"
    tell process "Keynote"
        -- Deselect / return to slide level
        key code 53
        delay 0.2
    end tell
end tell

-- Select the target shape via Keynote scripting
tell application "Keynote"
    tell document 1
        tell slide {slide_num}
            set selection to {{shape {shape_index}}}
        end tell
    end tell
end tell

delay 0.3

tell application "System Events"
    tell process "Keynote"
        -- Open Add Link dialog (Cmd+K)
        keystroke "k" using command down
        delay 0.5

        -- The link popover should appear — click the "Slide" tab/option
        -- Look for a pop up button to switch link type
        try
            set linkPopover to window 1
            -- Click the link type selector and choose Slide
            click pop up button 1 of linkPopover
            delay 0.3
            click menu item "Slide" of menu 1 of pop up button 1 of linkPopover
            delay 0.3
        end try

        -- Type the slide number
        keystroke "a" using command down
        delay 0.1
        keystroke "{to_slide}"
        delay 0.2

        -- Confirm
        key code 36 -- Return
        delay 0.2

        -- Deselect
        key code 53
        return "ok"
    end tell
end tell
"""


def command_insert_slide_links(args: argparse.Namespace) -> int:
    """Add slide navigation links to shapes in an open Keynote deck.

    Requires: macOS Accessibility permissions for the calling process
    and an open Keynote document as front document.
    """
    mappings_path = Path(args.mappings).resolve()
    if not mappings_path.exists():
        fail(f"Mappings file not found: {mappings_path}")
    with open(mappings_path) as f:
        data = json.load(f)
    entries: list[dict[str, Any]] = (
        data if isinstance(data, list) else data.get("slide_links", [])
    )
    if not entries:
        fail("No slide links found in mappings file")
    for i, entry in enumerate(entries):
        for key in ("slide", "shape", "to_slide"):
            if key not in entry:
                fail(f"Entry {i}: missing required key '{key}'")
        if not isinstance(entry["slide"], int) or entry["slide"] < 1:
            fail(f"Entry {i}: 'slide' must be a positive integer")
        if not isinstance(entry["shape"], int) or entry["shape"] < 1:
            fail(f"Entry {i}: 'shape' must be a positive integer")
        if not isinstance(entry["to_slide"], int) or entry["to_slide"] < 1:
            fail(f"Entry {i}: 'to_slide' must be a positive integer")

    dry_run: bool = args.dry_run
    succeeded = 0
    failed_entries: list[str] = []

    for i, entry in enumerate(entries):
        slide = entry["slide"]
        shape = entry["shape"]
        to_slide = entry["to_slide"]
        label = entry.get("label", f"shape {shape} → slide {to_slide}")

        print(
            f"[{i + 1}/{len(entries)}] Slide {slide} {label} … ",
            end="",
            flush=True,
        )

        if dry_run:
            print("(dry run)")
            succeeded += 1
            continue

        script = _build_slide_link_script(slide, shape, to_slide)
        if args.print_applescript:
            print()
            print(script)
            succeeded += 1
            continue

        try:
            output = run_osascript(script).strip()
        except RuntimeError as exc:
            output = f"error: {exc}"

        if output == "ok":
            print("OK")
            succeeded += 1
        else:
            print(f"FAILED ({output})")
            failed_entries.append(label)

    total = len(entries)
    n_failed = len(failed_entries)
    print(f"\nDone: {succeeded}/{total} linked, {n_failed} failed")
    if failed_entries:
        for label in failed_entries:
            print(f"  FAILED: {label}")
        return 1
    return 0


# ---------------------------------------------------------------------------
# Present
# ---------------------------------------------------------------------------

def build_present_applescript(file_path: Path, from_slide: int | None = None) -> str:
    lines = [
        f"with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds",
        "tell application \"Keynote\"",
        "  activate",
        f"  set theDoc to open {applescript_posix_file(file_path)}",
    ]
    if from_slide is not None:
        lines.append(f"  tell theDoc to set current slide to slide {from_slide}")
    lines.extend([
        "  start theDoc",
        "end tell",
        "end timeout",
    ])
    return "\n".join(lines)


def command_present(args: argparse.Namespace) -> int:
    file_path = Path(args.file).resolve()
    ensure_existing_file(file_path, "Input file", ".key")
    ensure_runtime_available()
    from_slide = args.from_slide
    if from_slide is not None and from_slide < 1:
        fail("--from must be >= 1")
    try:
        run_osascript(build_present_applescript(file_path, from_slide))
    except Exception as exc:
        fail(f"Present failed for {file_path}: {exc}")
    return 0


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="keynote-cli")
    subparsers = parser.add_subparsers(dest="command", required=True)

    run_parser = subparsers.add_parser("run", help="Run a keynote-cli script file")
    run_parser.add_argument("script", help="Path to script file")
    run_parser.add_argument("--force", action="store_true", help="Overwrite output if it already exists")
    run_parser.add_argument(
        "--keep-failed-output",
        action="store_true",
        help="Do not delete the partially built output file on failure",
    )
    run_parser.add_argument(
        "--print-applescript",
        action="store_true",
        help="Print the generated AppleScript instead of running it",
    )
    run_parser.add_argument(
        "--check-template",
        action="store_true",
        help="Open the template in Keynote and verify that all required master slides exist",
    )
    run_parser.set_defaults(func=command_run)

    inspect_parser = subparsers.add_parser("inspect", help="Inspect a .key file and print JSON")
    inspect_parser.add_argument("file", help="Path to .key file")
    inspect_parser.set_defaults(func=command_inspect)

    export_parser = subparsers.add_parser("export", help="Export a .key file")
    export_parser.add_argument("file", help="Path to .key file")
    export_parser.add_argument("--output", help="Output path")
    export_parser.add_argument("--format", default="pdf", help="Export format: pdf, png, jpeg, pptx, html, movie (default: pdf)")
    export_parser.add_argument("--force", action="store_true", help="Overwrite output if it already exists")
    export_parser.set_defaults(func=command_export)

    eq_parser = subparsers.add_parser(
        "insert-equations",
        help="Replace placeholder tokens in an open Keynote deck with rendered LaTeX equations",
    )
    eq_parser.add_argument(
        "mappings",
        help="JSON file: array of {slide, placeholder, latex} objects",
    )
    eq_parser.add_argument(
        "--render-timeout",
        type=int,
        default=15,
        metavar="SEC",
        help="Max seconds to wait for each equation to render (default: 15)",
    )
    eq_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate the mappings file without touching Keynote",
    )
    eq_parser.add_argument(
        "--print-applescript",
        action="store_true",
        help="Print generated AppleScript for each equation instead of running it",
    )
    eq_parser.set_defaults(func=command_insert_equations)

    link_parser = subparsers.add_parser(
        "insert-links",
        help="Add URL hyperlinks to text in an open Keynote deck (GUI scripting)",
    )
    link_parser.add_argument(
        "mappings",
        help='JSON file: array of {slide, find, url} objects',
    )
    link_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate the mappings file without touching Keynote",
    )
    link_parser.add_argument(
        "--print-applescript",
        action="store_true",
        help="Print generated AppleScript for each link instead of running it",
    )
    link_parser.set_defaults(func=command_insert_links)

    slide_link_parser = subparsers.add_parser(
        "insert-slide-links",
        help="Add slide navigation links to shapes in an open Keynote deck (GUI scripting)",
    )
    slide_link_parser.add_argument(
        "mappings",
        help='JSON file: array of {slide, shape, to_slide} objects',
    )
    slide_link_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate the mappings file without touching Keynote",
    )
    slide_link_parser.add_argument(
        "--print-applescript",
        action="store_true",
        help="Print generated AppleScript for each link instead of running it",
    )
    slide_link_parser.set_defaults(func=command_insert_slide_links)

    present_parser = subparsers.add_parser("present", help="Start a slideshow")
    present_parser.add_argument("file", help="Path to .key file")
    present_parser.add_argument("--from", type=int, dest="from_slide", help="Start from slide N (1-based)")
    present_parser.set_defaults(func=command_present)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return int(args.func(args))
    except KeynoteCLIError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Error: Interrupted", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
