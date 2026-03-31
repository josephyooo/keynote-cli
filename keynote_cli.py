#!/usr/bin/env python3
from __future__ import annotations

import argparse
import codecs
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, NoReturn


APPLESCRIPT_LONG_TIMEOUT_SECONDS = 7200
KEYNOTE_BUILD_BATCH_SIZE = 20

LAYOUT_SPECS: dict[str, dict[str, Any]] = {
    "Title": {
        "master": "Title",
        "content_order": ["title", "subtitle"],
        "accessors": {
            "title": "default title item",
            "subtitle": "default body item",
        },
    },
    "Header-Body": {
        "master": "Title & Bullets",
        "content_order": ["header", "body"],
        "accessors": {
            "header": "default title item",
            "body": "default body item",
        },
    },
    "Header-TwoCol": {
        "master": "Title & Bullets Two-Column",
        "content_order": ["header", "left", "right"],
        "accessors": {
            "header": "default title item",
            "left": "text item 3",
            "right": "text item 2",
        },
    },
    "Header-Body-TwoCol": {
        "master": "Title & Bullets Body over Two-Column",
        "content_order": ["header", "body", "left", "right"],
        "accessors": {
            "header": "default title item",
            "body": "text item 3",
            "left": "text item 4",
            "right": "text item 2",
        },
    },
    "Header-TwoCol-Body": {
        "master": "Title & Bullets Body under Two-Column",
        "content_order": ["header", "left", "right", "body"],
        "accessors": {
            "header": "default title item",
            "left": "text item 4",
            "right": "text item 2",
            "body": "text item 3",
        },
    },
    "Header-Body-TwoCol-Body": {
        "master": "Title & Bullets Body around Two-Column",
        "content_order": ["header", "body_top", "left", "right", "body_bottom"],
        "accessors": {
            "header": "default title item",
            "body_top": "text item 3",
            "left": "text item 4",
            "right": "text item 2",
            "body_bottom": "text item 5",
        },
    },
}

DEFAULT_TEXTBOX_STYLE = {
    "font": "HelveticaNeue",
    "fontSize": 50,
    "color": [0, 0, 0],
}

FIELD_SEP = chr(31)
ROOT_KEYS = {"template", "output", "slides"}
SLIDE_KEYS = {"layout", "content", "images", "text_boxes", "textBoxes", "overrides", "notes"}
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
TEXT_CAPABLE_OVERRIDE_TARGET_PREFIXES = ("content:", "textItem:", "shape:")
LAYOUTS_WITH_REAL_DEFAULT_BODY = {"Title", "Header-Body"}


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


def validate_override_target(target: Any, layout: str, field_name: str) -> str:
    if not isinstance(target, str) or not target:
        fail(f"{field_name} must be a non-empty string")

    if target == "defaultTitleItem":
        return target

    if target == "defaultBodyItem":
        if layout not in LAYOUTS_WITH_REAL_DEFAULT_BODY:
            fail(
                f"{field_name} uses defaultBodyItem on layout {layout!r}, but that layout's default body placeholder is hidden; "
                f"use content:<key> or textItem:<n> instead"
            )
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

    if target.startswith("content:"):
        content_key = target.split(":", 1)[1]
        if content_key not in LAYOUT_SPECS[layout]["accessors"]:
            fail(f"{field_name} references unknown content key {content_key!r} for layout {layout!r}")
        return target

    fail(
        f"{field_name} has invalid target {target!r}. Use one of: defaultTitleItem, defaultBodyItem, "
        f"content:<key>, textItem:<n>, image:<n>, shape:<n>"
    )


def slide_item_expression(accessor: str, slide_var: str = "newSlide") -> str:
    return f"{accessor} of {slide_var}"


def override_target_expression(target: str, layout: str, slide_var: str = "newSlide") -> str:
    target = str(target)
    if target == "defaultTitleItem":
        return f"default title item of {slide_var}"
    if target == "defaultBodyItem":
        return f"default body item of {slide_var}"
    if target.startswith("textItem:"):
        try:
            index = int(target.split(":", 1)[1])
        except ValueError:
            fail(f"Invalid override target: {target}")
        return f"text item {index} of {slide_var}"
    if target.startswith("image:"):
        try:
            index = int(target.split(":", 1)[1])
        except ValueError:
            fail(f"Invalid override target: {target}")
        return f"image {index} of {slide_var}"
    if target.startswith("shape:"):
        try:
            index = int(target.split(":", 1)[1])
        except ValueError:
            fail(f"Invalid override target: {target}")
        return f"shape {index} of {slide_var}"
    if target.startswith("content:"):
        content_key = target.split(":", 1)[1]
        accessor = LAYOUT_SPECS[layout]["accessors"].get(content_key)
        if accessor is None:
            fail(f"Unknown content key for override target {target!r} on layout {layout!r}")
        return slide_item_expression(accessor, slide_var)
    fail(
        "Invalid override target. Use one of: defaultTitleItem, defaultBodyItem, "
        "content:<key>, textItem:<n>, image:<n>, shape:<n>"
    )


def _validate_content_value(raw: Any, field: str) -> dict[str, Any]:
    """Normalise a content value into ``{"text": str, "indents": list[int] | None}``.

    Accepted input forms:
    - ``null`` / ``""`` / plain string  (no indent info)
    - An array of ``{"text": str, "indent": int}`` objects  (explicit indents)
    - An array of plain strings  (treated as indent 0 each)
    """
    if raw is None or (isinstance(raw, str) and raw == ""):
        return {"text": "", "indents": None}
    if isinstance(raw, str):
        return {"text": raw, "indents": None}
    if isinstance(raw, list):
        texts: list[str] = []
        indents: list[int] = []
        has_indent = False
        for j, item in enumerate(raw):
            item_field = f"{field}[{j}]"
            if isinstance(item, str):
                texts.append(item)
                indents.append(0)
            elif isinstance(item, dict):
                if "text" not in item:
                    fail(f"{item_field} must contain a 'text' key")
                texts.append(str(item["text"]))
                lvl = item.get("indent", 0)
                if not isinstance(lvl, int) or lvl < 0:
                    fail(f"{item_field}.indent must be a non-negative integer")
                indents.append(lvl)
                if lvl != 0:
                    has_indent = True
            else:
                fail(f"{item_field} must be a string or {{text, indent}} object")
        return {
            "text": "\n".join(texts),
            "indents": indents if has_indent else None,
        }
    fail(f"{field} must be a string or array of paragraphs")


def validate_instructions(data: Any, instructions_path: Path) -> dict[str, Any]:
    if not isinstance(data, dict):
        fail("Instruction file must contain a JSON object at the top level")

    ensure_allowed_keys(data, ROOT_KEYS, "instruction file")
    ensure_required_keys(data, ROOT_KEYS, "instruction file")

    base_dir = instructions_path.parent.resolve()

    template = resolve_path(ensure_non_empty_string(data["template"], "template"), base_dir)
    output = resolve_path(ensure_non_empty_string(data["output"], "output"), base_dir)

    ensure_existing_file(template, "template", ".key")
    ensure_output_suffix(output, ".key", "output")

    if template == output:
        fail("output must be different from template")

    slides = data["slides"]
    if not isinstance(slides, list) or not slides:
        fail("slides must be a non-empty array")

    validated_slides: list[dict[str, Any]] = []

    for slide_index, slide in enumerate(slides, start=1):
        slide_field = f"slides[{slide_index}]"
        if not isinstance(slide, dict):
            fail(f"{slide_field} must be an object")
        ensure_allowed_keys(slide, SLIDE_KEYS, slide_field)
        ensure_required_keys(slide, {"layout", "content"}, slide_field)

        if "text_boxes" in slide and "textBoxes" in slide:
            fail(f"{slide_field} cannot contain both 'text_boxes' and 'textBoxes'; use only one")

        layout = slide.get("layout")
        if not isinstance(layout, str) or layout not in LAYOUT_SPECS:
            fail(
                f"{slide_field}.layout must be one of: "
                + ", ".join(sorted(LAYOUT_SPECS.keys()))
            )

        content_field = f"{slide_field}.content"
        content = slide["content"]
        if not isinstance(content, dict):
            fail(f"{content_field} must be an object")

        required_content_keys = set(LAYOUT_SPECS[layout]["content_order"])
        ensure_allowed_keys(content, required_content_keys, content_field)
        ensure_required_keys(content, required_content_keys, content_field)
        validated_content: dict[str, dict[str, Any]] = {}
        for key in LAYOUT_SPECS[layout]["content_order"]:
            raw = content[key]
            validated_content[key] = _validate_content_value(
                raw, f"{content_field}.{key}",
            )

        validated_images: list[dict[str, Any]] = []
        images = slide.get("images", [])
        if not isinstance(images, list):
            fail(f"{slide_field}.images must be an array")
        for image_index, image in enumerate(images, start=1):
            image_field = f"{slide_field}.images[{image_index}]"
            if not isinstance(image, dict):
                fail(f"{image_field} must be an object")
            ensure_allowed_keys(image, IMAGE_KEYS, image_field)
            ensure_required_keys(image, {"file"}, image_field)
            image_path = resolve_path(ensure_non_empty_string(image["file"], f"{image_field}.file"), base_dir)
            ensure_existing_file(image_path, f"{image_field}.file")
            position = validate_point(image.get("position", [0, 0]), f"{image_field}.position")
            size = None
            if "size" in image:
                size = validate_size(image["size"], f"{image_field}.size")
            validated_images.append({
                "file": image_path,
                "position": position,
                "size": size,
            })

        validated_text_boxes: list[dict[str, Any]] = []
        text_boxes = slide.get("text_boxes", slide.get("textBoxes", []))
        if not isinstance(text_boxes, list):
            fail(f"{slide_field}.text_boxes must be an array")
        for box_index, box in enumerate(text_boxes, start=1):
            box_field = f"{slide_field}.text_boxes[{box_index}]"
            if not isinstance(box, dict):
                fail(f"{box_field} must be an object")
            ensure_allowed_keys(box, TEXT_BOX_KEYS, box_field)
            ensure_required_keys(box, {"text", "position", "size"}, box_field)
            position = validate_point(box["position"], f"{box_field}.position")
            size = validate_size(box["size"], f"{box_field}.size")
            font = ensure_non_empty_string(box.get("font", DEFAULT_TEXTBOX_STYLE["font"]), f"{box_field}.font")
            font_size = ensure_number(
                box.get("fontSize", box.get("font_size", DEFAULT_TEXTBOX_STYLE["fontSize"])),
                f"{box_field}.fontSize",
                minimum=0.01,
            )
            color = normalize_color(box.get("color", DEFAULT_TEXTBOX_STYLE["color"]), f"{box_field}.color")
            validated_text_boxes.append({
                "text": "" if box.get("text") is None else str(box.get("text")),
                "position": position,
                "size": size,
                "font": font,
                "fontSize": float(font_size),
                "color": color,
            })

        validated_overrides: list[dict[str, Any]] = []
        overrides = slide.get("overrides", [])
        if not isinstance(overrides, list):
            fail(f"{slide_field}.overrides must be an array")
        for override_index, override in enumerate(overrides, start=1):
            override_field = f"{slide_field}.overrides[{override_index}]"
            if not isinstance(override, dict):
                fail(f"{override_field} must be an object")
            ensure_allowed_keys(override, OVERRIDE_KEYS, override_field)
            ensure_required_keys(override, {"target"}, override_field)
            if len(set(override) - {"target"}) == 0:
                fail(f"{override_field} must include at least one change besides 'target'")

            target = validate_override_target(override["target"], layout, f"{override_field}.target")
            validated_override: dict[str, Any] = {"target": target}

            if target.startswith("image:"):
                illegal_fields = sorted({"text", "font", "fontSize", "font_size", "color"} & set(override))
                if illegal_fields:
                    fail(
                        f"{override_field} targets an image, so it cannot set text properties: "
                        + ", ".join(repr(field) for field in illegal_fields)
                    )

            if "text" in override:
                validated_override["text"] = "" if override["text"] is None else str(override["text"])
            if "position" in override:
                validated_override["position"] = validate_point(override["position"], f"{override_field}.position")
            if "size" in override:
                validated_override["size"] = validate_size(override["size"], f"{override_field}.size")
            if "font" in override:
                validated_override["font"] = ensure_non_empty_string(override["font"], f"{override_field}.font")
            if "fontSize" in override or "font_size" in override:
                font_size = override.get("fontSize", override.get("font_size"))
                validated_override["fontSize"] = ensure_number(
                    font_size,
                    f"{override_field}.fontSize",
                    minimum=0.01,
                )
            if "color" in override:
                validated_override["color"] = normalize_color(override["color"], f"{override_field}.color")
            if "opacity" in override:
                validated_override["opacity"] = ensure_number(
                    override["opacity"],
                    f"{override_field}.opacity",
                    minimum=0,
                    maximum=100,
                )
            if "rotation" in override:
                validated_override["rotation"] = ensure_number(
                    override["rotation"],
                    f"{override_field}.rotation",
                )
            validated_overrides.append(validated_override)

        notes = slide.get("notes")
        if notes is not None:
            notes = str(notes)

        validated_slides.append({
            "layout": layout,
            "content": validated_content,
            "images": validated_images,
            "text_boxes": validated_text_boxes,
            "overrides": validated_overrides,
            "notes": notes,
        })

    return {
        "template": template,
        "output": output,
        "slides": validated_slides,
    }


def build_slide_applescript(slide: dict[str, Any], slide_number: int) -> list[str]:
    layout = slide["layout"]
    spec = LAYOUT_SPECS[layout]
    body_lines: list[str] = []
    body_lines.append(
        f"set newSlide to make new slide with properties {{base slide: master slide {applescript_string(spec['master'])}}}"
    )

    for key in spec["content_order"]:
        accessor = spec["accessors"][key]
        expr = slide_item_expression(accessor)
        content_val = slide["content"].get(key, {"text": "", "indents": None})
        if isinstance(content_val, str):
            # Legacy plain-string path (in case callers bypass validation)
            body_lines.append(f"set object text of {expr} to {applescript_string(content_val)}")
        else:
            body_lines.append(f"set object text of {expr} to {applescript_string(content_val['text'])}")
            if content_val.get("indents"):
                body_lines.append(f"tell object text of {expr}")
                for para_idx, indent_lvl in enumerate(content_val["indents"], start=1):
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
        target_expr = override_target_expression(override["target"], layout)
        if "text" in override:
            body_lines.append(f"set object text of {target_expr} to {applescript_string(override['text'])}")
        if "position" in override:
            body_lines.append(
                f"set position of {target_expr} to {{{numeric_literal(override['position'][0])}, {numeric_literal(override['position'][1])}}}"
            )
        if "size" in override:
            body_lines.append(f"set width of {target_expr} to {numeric_literal(override['size'][0])}")
            body_lines.append(f"set height of {target_expr} to {numeric_literal(override['size'][1])}")
        if "font" in override:
            body_lines.append(f"set font of object text of {target_expr} to {applescript_string(override['font'])}")
        if "fontSize" in override:
            body_lines.append(f"set size of object text of {target_expr} to {numeric_literal(override['fontSize'])}")
        if "color" in override:
            body_lines.append(
                f"set color of object text of {target_expr} to {{{override['color'][0]}, {override['color'][1]}, {override['color'][2]}}}"
            )
        if "opacity" in override:
            body_lines.append(f"set opacity of {target_expr} to {numeric_literal(override['opacity'])}")
        if "rotation" in override:
            body_lines.append(f"set rotation of {target_expr} to {numeric_literal(override['rotation'])}")

    lines: list[str] = [
        f"    -- Slide {slide_number}: {layout}",
        "    try",
    ]
    lines.extend(f"      {line}" for line in body_lines)
    lines.extend(
        [
            "    on error errMsg number errNum",
            f"      error {applescript_string(f'Slide {slide_number} ({layout}) failed: ')} & errMsg number errNum",
            "    end try",
            "",
        ]
    )
    return lines


def build_build_applescript(
    output_path: Path,
    slides: list[dict[str, Any]],
    *,
    start_slide_number: int = 1,
    delete_template_slide_count: int | None = None,
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

    if delete_template_slide_count is not None:
        lines.extend(
            [
                f"      repeat with i from {delete_template_slide_count} to 1 by -1",
                "        delete slide i",
                "      end repeat",
            ]
        )

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
    return "\n".join(lines)


def run_osascript(script: str) -> str:
    result = subprocess.run(
        ["osascript", "-"],
        input=script,
        text=True,
        capture_output=True,
        encoding="utf-8",
    )
    if result.returncode != 0:
        error_parts = []
        if result.stderr and result.stderr.strip():
            error_parts.append(result.stderr.strip())
        if result.stdout and result.stdout.strip():
            error_parts.append(result.stdout.strip())
        error_text = "\n".join(error_parts).strip()
        if not error_text:
            error_text = f"osascript failed with exit code {result.returncode}"
        raise RuntimeError(error_text)
    return result.stdout


def decode_escaped(value: str) -> str:
    return codecs.decode(value, "unicode_escape")


def build_inspect_applescript(file_path: Path) -> str:
    backslash = applescript_string("\\")
    double_backslash = applescript_string("\\\\")
    escaped_field_sep = applescript_string("\\u001f")
    escaped_newline = applescript_string("\\n")

    return f'''on replace_text(find_text, replace_text, source_text)
  set AppleScript's text item delimiters to find_text
  set parts to text items of source_text
  set AppleScript's text item delimiters to replace_text
  set result_text to parts as text
  set AppleScript's text item delimiters to ""
  return result_text
end replace_text

on escape_text(source_text, field_sep)
  set t to source_text as text
  set t to my replace_text({backslash}, {double_backslash}, t)
  set t to my replace_text(field_sep, {escaped_field_sep}, t)
  set t to my replace_text(return, {escaped_newline}, t)
  set t to my replace_text(linefeed, {escaped_newline}, t)
  return t
end escape_text

on join_lines(line_list)
  set AppleScript's text item delimiters to linefeed
  set joined_text to line_list as text
  set AppleScript's text item delimiters to ""
  return joined_text
end join_lines

set fieldSep to ASCII character 31
set outLines to {{}}

tell application "Keynote"
  set inputFile to {applescript_posix_file(file_path)}
  set theDoc to open inputFile
  tell theDoc
    set end of outLines to "DOC" & fieldSep & width & fieldSep & height & fieldSep & (count of slides)
    repeat with i from 1 to count of master slides
      set end of outLines to "MASTER" & fieldSep & i & fieldSep & my escape_text(name of master slide i, fieldSep)
    end repeat
    repeat with i from 1 to count of slides
      set s to slide i
      set masterName to ""
      set notesText to ""
      try
        set masterName to name of base slide of s
      end try
      try
        set notesText to presenter notes of s
      end try
      set end of outLines to "SLIDE" & fieldSep & i & fieldSep & my escape_text(masterName, fieldSep)
      set end of outLines to "NOTES" & fieldSep & i & fieldSep & my escape_text(notesText, fieldSep)
      repeat with j from 1 to count of text items of s
        set ti to text item j of s
        set tiPos to position of ti
        set end of outLines to "TEXT" & fieldSep & i & fieldSep & j & fieldSep & my escape_text(object text of ti, fieldSep) & fieldSep & item 1 of tiPos & fieldSep & item 2 of tiPos & fieldSep & width of ti & fieldSep & height of ti
      end repeat
      repeat with j from 1 to count of images of s
        set img to image j of s
        set imgPos to position of img
        set end of outLines to "IMAGE" & fieldSep & i & fieldSep & j & fieldSep & item 1 of imgPos & fieldSep & item 2 of imgPos & fieldSep & width of img & fieldSep & height of img
      end repeat
      repeat with j from 1 to count of shapes of s
        set sh to shape j of s
        set shPos to position of sh
        set shText to ""
        try
          set shText to object text of sh
        end try
        set end of outLines to "SHAPE" & fieldSep & i & fieldSep & j & fieldSep & my escape_text(shText, fieldSep) & fieldSep & item 1 of shPos & fieldSep & item 2 of shPos & fieldSep & width of sh & fieldSep & height of sh
      end repeat
    end repeat
  end tell
  close theDoc saving no
end tell

return my join_lines(outLines)
'''


def filter_text_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    filtered: list[dict[str, Any]] = []
    seen_positions: set[tuple[float, float, float, float]] = set()
    for item in items:
        width, height = item["size"]
        if width <= 0 or height <= 0:
            continue
        key = (*item["position"], *item["size"])
        if key in seen_positions:
            continue
        seen_positions.add(key)
        filtered.append(item)
    return filtered


def inspect_file(file_path: Path) -> dict[str, Any]:
    output = run_osascript(build_inspect_applescript(file_path))
    slides: dict[int, dict[str, Any]] = {}
    masters: list[str] = []
    dimensions: list[int] = [0, 0]
    slide_count = 0

    for raw_line in output.splitlines():
        if not raw_line:
            continue
        parts = raw_line.split(FIELD_SEP)
        record_type = parts[0]

        if record_type == "DOC":
            dimensions = [int(float(parts[1])), int(float(parts[2]))]
            slide_count = int(parts[3])
        elif record_type == "MASTER":
            masters.append(decode_escaped(parts[2]))
        elif record_type == "SLIDE":
            slide_index = int(parts[1])
            slides[slide_index] = {
                "index": slide_index,
                "master": decode_escaped(parts[2]),
                "notes": "",
                "textItems": [],
                "images": [],
                "shapes": [],
            }
        elif record_type == "NOTES":
            slide_index = int(parts[1])
            slides.setdefault(slide_index, {
                "index": slide_index,
                "master": "",
                "notes": "",
                "textItems": [],
                "images": [],
                "shapes": [],
            })["notes"] = decode_escaped(parts[2])
        elif record_type == "TEXT":
            slide_index = int(parts[1])
            slides[slide_index]["textItems"].append({
                "index": int(parts[2]),
                "text": decode_escaped(parts[3]),
                "position": [float(parts[4]), float(parts[5])],
                "size": [float(parts[6]), float(parts[7])],
            })
        elif record_type == "IMAGE":
            slide_index = int(parts[1])
            slides[slide_index]["images"].append({
                "index": int(parts[2]),
                "position": [float(parts[3]), float(parts[4])],
                "size": [float(parts[5]), float(parts[6])],
            })
        elif record_type == "SHAPE":
            slide_index = int(parts[1])
            slides[slide_index]["shapes"].append({
                "index": int(parts[2]),
                "text": decode_escaped(parts[3]),
                "position": [float(parts[4]), float(parts[5])],
                "size": [float(parts[6]), float(parts[7])],
            })

    ordered_slides = [slides[index] for index in sorted(slides)]
    for slide in ordered_slides:
        slide["textItems"] = filter_text_items(slide["textItems"])
        slide["shapes"] = [
            shape for shape in slide["shapes"] if shape["size"][0] > 0 and shape["size"][1] > 0
        ]

    return {
        "file": str(file_path),
        "dimensions": dimensions,
        "slideCount": slide_count,
        "masters": masters,
        "slides": ordered_slides,
    }


def build_export_applescript(input_path: Path, output_path: Path) -> str:
    return f'''with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds
tell application "Keynote"
  set theDoc to missing value
  try
    set inputFile to {applescript_posix_file(input_path)}
    set outputFile to {applescript_posix_file(output_path)}
    set theDoc to open inputFile
    tell theDoc
      export to outputFile as PDF
    end tell
    close theDoc saving no
  on error errMsg number errNum
    try
      if theDoc is not missing value then close theDoc saving no
    end try
    error "Export failed: " & errMsg number errNum
  end try
end tell
end timeout
'''


def load_instruction_json(instructions_path: Path) -> Any:
    ensure_existing_file(instructions_path, "Instruction file")
    try:
        with instructions_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        fail(
            f"Invalid JSON in {instructions_path} at line {exc.lineno}, column {exc.colno}: {exc.msg}"
        )


def validate_template_masters(template_path: Path, slides: list[dict[str, Any]]) -> dict[str, Any]:
    required_masters = {LAYOUT_SPECS[slide['layout']]['master'] for slide in slides}
    template_info = inspect_file(template_path)
    available_masters = set(template_info["masters"])
    missing = sorted(required_masters - available_masters)
    if missing:
        fail(
            f"Template is missing required master slide(s): {', '.join(repr(name) for name in missing)}"
        )
    return template_info


def command_build(args: argparse.Namespace) -> int:
    instructions_path = Path(args.instructions).resolve()
    raw_data = load_instruction_json(instructions_path)
    instructions = validate_instructions(raw_data, instructions_path)
    output_path: Path = instructions["output"]
    if args.print_applescript:
        print(build_build_applescript(output_path, instructions["slides"]))
        return 0

    ensure_runtime_available()
    template_info = validate_template_masters(instructions["template"], instructions["slides"])
    template_slide_count = int(template_info["slideCount"])

    if output_path.exists():
        if args.force:
            remove_path(output_path)
        else:
            fail(f"Output already exists: {output_path} (use --force to overwrite)")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(instructions["template"], output_path)

    removed_partial_output = False
    try:
        total_slides = len(instructions["slides"])
        for batch_start in range(0, total_slides, KEYNOTE_BUILD_BATCH_SIZE):
            batch_slides = instructions["slides"][batch_start: batch_start + KEYNOTE_BUILD_BATCH_SIZE]
            delete_template_slide_count = template_slide_count if batch_start + len(batch_slides) >= total_slides else None
            script = build_build_applescript(
                output_path,
                batch_slides,
                start_slide_number=batch_start + 1,
                delete_template_slide_count=delete_template_slide_count,
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

    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_path = input_path.with_suffix(".pdf")
    ensure_output_suffix(output_path, ".pdf", "Export output")

    if output_path.exists():
        if args.force:
            remove_path(output_path)
        else:
            fail(f"Output already exists: {output_path} (use --force to overwrite)")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        run_osascript(build_export_applescript(input_path, output_path))
    except Exception as exc:
        fail(f"Export failed for {input_path}: {exc}")
    print(str(output_path))
    return 0


def command_validate(args: argparse.Namespace) -> int:
    instructions_path = Path(args.instructions).resolve()
    raw_data = load_instruction_json(instructions_path)
    instructions = validate_instructions(raw_data, instructions_path)
    if args.check_template:
        ensure_runtime_available()
        validate_template_masters(instructions["template"], instructions["slides"])
    print(json.dumps({
        "ok": True,
        "template": str(instructions["template"]),
        "output": str(instructions["output"]),
        "slides": len(instructions["slides"]),
        "checkedTemplateMasters": bool(args.check_template),
    }, indent=2))
    return 0


def _build_equation_insert_script(
    slide_num: int, placeholder: str, latex: str, render_timeout: int,
) -> str:
    """Build AppleScript that finds *placeholder* on *slide_num* and replaces
    it with a rendered LaTeX equation via GUI scripting (System Events)."""
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
    """Replace [PLACEHOLDER] tokens in a Keynote slide with rendered LaTeX
    equations using the Insert > Equation GUI.

    Requires: macOS Accessibility permissions for the calling process
    (Terminal / osascript) and an open Keynote document as front document.
    """
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="keynote-cli")
    subparsers = parser.add_subparsers(dest="command", required=True)

    build_parser_ = subparsers.add_parser("build", help="Build a presentation from an instruction JSON file")
    build_parser_.add_argument("instructions", help="Path to instructions.json")
    build_parser_.add_argument("--force", action="store_true", help="Overwrite output if it already exists")
    build_parser_.add_argument(
        "--keep-failed-output",
        action="store_true",
        help="Do not delete the partially built output file if Keynote fails",
    )
    build_parser_.add_argument(
        "--print-applescript",
        action="store_true",
        help="Print the generated AppleScript instead of running it",
    )
    build_parser_.set_defaults(func=command_build)

    validate_parser = subparsers.add_parser("validate", help="Validate an instruction JSON file")
    validate_parser.add_argument("instructions", help="Path to instructions.json")
    validate_parser.add_argument(
        "--check-template",
        action="store_true",
        help="Also open the template in Keynote and verify that all required master slides exist",
    )
    validate_parser.set_defaults(func=command_validate)

    inspect_parser = subparsers.add_parser("inspect", help="Inspect a .key file and print JSON")
    inspect_parser.add_argument("file", help="Path to .key file")
    inspect_parser.set_defaults(func=command_inspect)

    export_parser = subparsers.add_parser("export", help="Export a .key file to PDF")
    export_parser.add_argument("file", help="Path to .key file")
    export_parser.add_argument("--output", help="Output PDF path")
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
