# keynote-cli

CLI for automating Keynote presentations on macOS via AppleScript.

## Prerequisites

- macOS with Keynote installed
- Python 3
- For `insert-equations`: Terminal must have Accessibility permissions (System Settings > Privacy & Security > Accessibility)

## Commands

```bash
keynote-cli run script.txt                               # Run a script file
keynote-cli run script.txt --check-template              # Also verify masters exist
keynote-cli run script.txt --force                       # Overwrite existing output
keynote-cli run script.txt --print-applescript            # Print generated AppleScript
keynote-cli inspect file.key                             # Dump slide structure as JSON
keynote-cli export file.key --output file.pdf            # Export to PDF (default)
keynote-cli export file.key --format png --output slides/# Export as PNG images
keynote-cli export file.key --format pptx                # Export as PowerPoint
keynote-cli export file.key --format movie               # Export as movie
keynote-cli present file.key                             # Start slideshow
keynote-cli present file.key --from 5                    # Start from slide 5
keynote-cli insert-links links.json                      # Add URL hyperlinks (GUI scripting)
keynote-cli insert-slide-links nav.json                  # Add slide navigation links (GUI scripting)
keynote-cli insert-equations equations.json              # Insert LaTeX equations (GUI scripting)
```

## Script format

A script file is a newline-delimited sequence of commands:

```
# Build a presentation from a template
open template.key --output out.key

# Create slides from master layouts
add-slide --master "Title"
set-text --slide 1 --target defaultTitleItem "Presentation Title"
set-text --slide 1 --target defaultBodyItem "Author Name"

add-slide --master "Title & Bullets"
set-text --slide 2 --target defaultTitleItem "Section Header"
set-text --slide 2 --target defaultBodyItem "Point 1\nSub-point A" --indents 0,1
add-image --slide 2 --file figure.png --position 960,300 --size 800,500

# Clean up and save
delete-slides 1-7
save
```

Lines starting with `#` are comments. Blank lines are ignored.

## Script commands

| Command | Description |
|---------|-------------|
| `open TEMPLATE --output OUTPUT [--force]` | Copy template to output path, open in Keynote |
| `add-slide --master NAME` | Create a new slide from the named master |
| `set-text --slide N --target TARGET TEXT [--indents 0,1,2]` | Set text on a slide item |
| `set-notes --slide N TEXT` | Set presenter notes |
| `add-image --slide N --file PATH --position X,Y [--size W,H]` | Insert image |
| `add-text-box --slide N --text TEXT --position X,Y --size W,H [--font F] [--font-size S] [--color R,G,B]` | Insert free text box |
| `override --slide N --target TARGET [--text T] [--position X,Y] [--size W,H] [--font F] [--font-size S] [--color R,G,B] [--opacity O] [--rotation R]` | Modify existing element |
| `duplicate-slide --slide N [--to M]` | Duplicate a slide (optionally to after slide M) |
| `move-slide --slide N --to M` | Move slide N to position M |
| `skip-slide --slide N` | Hide slide from presentation |
| `unskip-slide --slide N` | Unhide slide |
| `replace-text --find "X" --replace "Y" [--slide N]` | Find/replace text across slides |
| `set-style --slide N --target TARGET [--bold] [--italic] [--underline]` | Set text style (use `--no-bold` etc. to unset) |
| `add-shape --slide N --position X,Y --size W,H [--text T] [--rotation D] [--opacity O]` | Add a shape |
| `add-line --slide N --from X,Y --to X,Y` | Add a line |
| `duplicate-shape --slide N --index I --to-slide M` | Copy a shape to another slide |
| `delete-shape --slide N --index I` | Delete a shape |
| `delete-image --slide N --index I` | Delete an image |
| `set-master --slide N --master NAME` | Change a slide's master (base slide) |
| `add-table --slide N --rows R --cols C [--position X,Y] [--size W,H]` | Add a table |
| `set-cell --slide N --row R --col C VALUE [--table I]` | Set a table cell value |
| `add-row --slide N [--table I]` | Add a row to a table |
| `add-col --slide N [--table I]` | Add a column to a table |
| `delete-row --slide N --row R [--table I]` | Delete a table row |
| `delete-col --slide N --col C [--table I]` | Delete a table column |
| `set-transition --slide N --style TYPE [--duration S]` | Set slide transition |
| `set-theme --theme NAME` | Apply a theme to the entire document |
| `delete-slides RANGE` | Delete slides (e.g. `1-7` or `5`) |
| `save` | Save and close the document |

### Target notation

- `defaultTitleItem` — the slide's default title placeholder
- `defaultBodyItem` — the slide's default body placeholder
- `textItem:N` — text item by 1-based index
- `image:N` — image by 1-based index
- `shape:N` — shape by 1-based index

### Text escaping in scripts

Use `\n` for newlines, `\t` for tabs, `\\` for literal backslashes within text arguments. Quote arguments with spaces using shell quoting.

## Performance

`keynote-cli run` compiles the entire script into batched AppleScript (20 slides per `osascript` call) for performance. Individual commands are not executed one at a time.

## insert-equations

See `AGENTS.md` for the input format. Replaces `[PLACEHOLDER]` tokens in an already-open Keynote deck with rendered LaTeX equations via the Insert > Equation GUI.

## Notes

- Run one command at a time — Keynote scripting is not concurrency-safe.
- Build failures show the failing slide number and master name.
- Source code: `keynote-cli` (entry point) imports `keynote_cli.py`.
