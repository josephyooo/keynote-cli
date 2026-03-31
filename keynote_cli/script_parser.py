from __future__ import annotations

import argparse
from pathlib import Path
import shlex
from typing import Any

from keynote_cli.common import (
    DEFAULT_TEXTBOX_STYLE,
    ensure_existing_file,
    ensure_output_suffix,
    fail,
    normalize_color,
    resolve_path,
    validate_target,
    validate_text_target,
)


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
        fail(f"Color must be R,G,B - got {raw!r}")
    try:
        return normalize_color([int(p) for p in parts], "color")
    except ValueError:
        fail(f"Invalid color: {raw!r}")


def _unescape_script_text(text: str) -> str:
    placeholder = "\x00"
    return (
        text.replace("\\\\", placeholder)
        .replace("\\n", "\n")
        .replace("\\t", "\t")
        .replace(placeholder, "\\")
    )


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
        validate_text_target(args.target, f"line {line_num} target")
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
        if args.to_pos is not None and args.to_pos < 1:
            fail(f"Line {line_num}: --to must be >= 1")
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
        validate_text_target(args.target, f"line {line_num} target")
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
