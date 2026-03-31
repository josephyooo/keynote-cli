from __future__ import annotations

from pathlib import Path
from typing import Any

from keynote_cli.common import (
    APPLESCRIPT_LONG_TIMEOUT_SECONDS,
    applescript_posix_file,
    applescript_string,
    numeric_literal,
    target_to_applescript,
)


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
        lines.append(f"        add row below row (row count)")
        lines.append(f"      end tell")

    elif kind == "add-col":
        lines.append(f"      tell table {op['table']} of slide {op['slide']}")
        lines.append(f"        add column after column (column count)")
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
        effect = "no transition effect" if style == "none" else style
        props = [f"transition effect:{effect}"]
        if "duration" in op:
            props.append(f"transition duration:{numeric_literal(op['duration'])}")
        lines.append(f"      set transition properties of slide {slide_num} to {{{', '.join(props)}}}")

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
