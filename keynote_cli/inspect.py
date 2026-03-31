from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from keynote_cli.common import (
    APPLESCRIPT_LONG_TIMEOUT_SECONDS,
    FIELD_SEP,
    applescript_posix_file,
    decode_escaped,
    ensure_existing_file,
    ensure_runtime_available,
    fail,
    run_osascript,
)


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
        if c is "\\\\" then
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


def build_inspect_masters_applescript(file_path: Path) -> str:
    sep = FIELD_SEP
    return f"""\
with timeout of {APPLESCRIPT_LONG_TIMEOUT_SECONDS} seconds
tell application "Keynote"
    set theDoc to open {applescript_posix_file(file_path)}
    try
        tell theDoc
            set originalCount to count of slides
            set masterNames to {{}}
            repeat with ms in master slides
                set end of masterNames to name of ms
            end repeat

            set resultLines to {{}}
            repeat with masterIdx from 1 to count of masterNames
                set masterName to item masterIdx of masterNames
                set newSlide to make new slide with properties {{base slide: master slide masterName}}
                set end of resultLines to ("MASTER{sep}" & masterName)

                -- defaultTitleItem
                try
                    set dti to default title item of newSlide
                    set dtiPos to position of dti
                    set dtiW to width of dti
                    set dtiH to height of dti
                    set end of resultLines to ("DTI{sep}" & (item 1 of dtiPos) & "{sep}" & (item 2 of dtiPos) & "{sep}" & dtiW & "{sep}" & dtiH)
                on error
                    set end of resultLines to ("DTI{sep}NONE")
                end try

                -- defaultBodyItem
                try
                    set dbi to default body item of newSlide
                    set dbiPos to position of dbi
                    set dbiW to width of dbi
                    set dbiH to height of dbi
                    set end of resultLines to ("DBI{sep}" & (item 1 of dbiPos) & "{sep}" & (item 2 of dbiPos) & "{sep}" & dbiW & "{sep}" & dbiH)
                on error
                    set end of resultLines to ("DBI{sep}NONE")
                end try

                -- all text items
                try
                    set tiCount to count of text items of newSlide
                    repeat with tiIdx from 1 to tiCount
                        set ti to text item tiIdx of newSlide
                        set tiPos to position of ti
                        set tiW to width of ti
                        set tiH to height of ti
                        set end of resultLines to ("TI{sep}" & tiIdx & "{sep}" & (item 1 of tiPos) & "{sep}" & (item 2 of tiPos) & "{sep}" & tiW & "{sep}" & tiH)
                    end repeat
                end try

                delete slide (originalCount + 1)
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
"""


def inspect_masters_file(file_path: Path) -> list[dict[str, Any]]:
    script = build_inspect_masters_applescript(file_path)
    raw = run_osascript(script)

    masters: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    for line in raw.splitlines():
        parts = line.split(FIELD_SEP)
        kind = parts[0]
        if kind == "MASTER":
            current = {"master": parts[1], "defaultTitleItem": None, "defaultBodyItem": None, "textItems": []}
            masters.append(current)
        elif kind == "DTI" and current is not None:
            if parts[1] != "NONE":
                sz = [float(parts[3]), float(parts[4])]
                if sz[0] > 0 and sz[1] > 0:
                    current["defaultTitleItem"] = {
                        "position": [float(parts[1]), float(parts[2])],
                        "size": sz,
                    }
        elif kind == "DBI" and current is not None:
            if parts[1] != "NONE":
                sz = [float(parts[3]), float(parts[4])]
                if sz[0] > 0 and sz[1] > 0:
                    current["defaultBodyItem"] = {
                        "position": [float(parts[1]), float(parts[2])],
                        "size": sz,
                    }
        elif kind == "TI" and current is not None:
            current["textItems"].append({
                "index": int(parts[1]),
                "position": [float(parts[2]), float(parts[3])],
                "size": [float(parts[4]), float(parts[5])],
            })

    # Annotate text items with target notation
    for master in masters:
        dti = master["defaultTitleItem"]
        dbi = master["defaultBodyItem"]
        visible_items: list[dict[str, Any]] = []
        for ti in master["textItems"]:
            w, h = ti["size"]
            if w == 0 and h == 0:
                ti["target"] = None
                ti["hidden"] = True
                continue
            ti["hidden"] = False
            # Match against defaultTitleItem/defaultBodyItem by position
            if dti and ti["position"] == dti["position"] and ti["size"] == dti["size"]:
                ti["target"] = "defaultTitleItem"
            elif dbi and ti["position"] == dbi["position"] and ti["size"] == dbi["size"]:
                ti["target"] = "defaultBodyItem"
            else:
                ti["target"] = f"textItem:{ti['index']}"
            visible_items.append(ti)
        # Filter out duplicates (same position+size)
        seen: set[tuple[float, float, float, float]] = set()
        deduped: list[dict[str, Any]] = []
        for ti in visible_items:
            key = (ti["position"][0], ti["position"][1], ti["size"][0], ti["size"][1])
            if key in seen:
                continue
            seen.add(key)
            deduped.append(ti)
        master["textItems"] = deduped

    return masters


def command_inspect_masters(args: argparse.Namespace) -> int:
    file_path = Path(args.file).resolve()
    ensure_existing_file(file_path, "Input file", ".key")
    ensure_runtime_available()
    try:
        result = inspect_masters_file(file_path)
    except Exception as exc:
        fail(f"Inspect masters failed for {file_path}: {exc}")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0
