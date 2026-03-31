from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from keynote_cli.common import (
    APPLESCRIPT_LONG_TIMEOUT_SECONDS,
    applescript_posix_file,
    applescript_string,
    ensure_runtime_available,
    fail,
    run_osascript,
)


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
    ensure_runtime_available()
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
    ensure_runtime_available()
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
    ensure_runtime_available()
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
