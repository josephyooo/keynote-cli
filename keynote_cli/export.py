from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from keynote_cli.common import (
    APPLESCRIPT_LONG_TIMEOUT_SECONDS,
    applescript_posix_file,
    applescript_string,
    ensure_existing_file,
    ensure_runtime_available,
    fail,
    remove_path,
    run_osascript,
)


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

    if output_path == input_path:
        fail(f"Output path is the same as input: {output_path}")

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
