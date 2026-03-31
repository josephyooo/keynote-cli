from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
from typing import Any

from keynote_cli.common import (
    KEYNOTE_BUILD_BATCH_SIZE,
    KeynoteCLIError,
    ensure_runtime_available,
    fail,
    remove_path,
    run_osascript,
)
from keynote_cli.script_parser import parse_script, _group_operations_into_slides
from keynote_cli.build import build_build_applescript
from keynote_cli.inspect import command_inspect, command_inspect_masters, inspect_file
from keynote_cli.export import command_export, command_present
from keynote_cli.gui import command_insert_equations, command_insert_links, command_insert_slide_links


def validate_template_masters(
    template_path: Path,
    slides: list[dict[str, Any]],
    doc_ops: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    info = inspect_file(template_path)
    available = set(info.get("masters", []))
    required = {slide["master"] for slide in slides}
    if doc_ops:
        required |= {op["master"] for op in doc_ops if op["op"] == "set-master"}
    missing = sorted(required - available)
    if missing:
        fail(
            f"Template is missing required master slide(s): {', '.join(repr(name) for name in missing)}. "
            f"Available: {', '.join(repr(name) for name in sorted(available))}"
        )
    return info


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
        validate_template_masters(plan["template"], plan["slides"], plan["doc_ops"])

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
        if total_slides == 0 and plan["doc_ops"]:
            script = build_build_applescript(output_path, [], doc_ops=plan["doc_ops"])
            run_osascript(script)
        else:
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

    inspect_masters_parser = subparsers.add_parser("inspect-masters", help="Inspect master slide text item layout")
    inspect_masters_parser.add_argument("file", help="Path to .key file")
    inspect_masters_parser.set_defaults(func=command_inspect_masters)

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
