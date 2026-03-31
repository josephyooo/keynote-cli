# keynote-cli

CLI for building Keynote presentations from a JSON instruction file and a template `.key` file.

## Prerequisites

- macOS with Keynote installed
- Python 3
- For `insert-equations`: Terminal must have Accessibility permissions (System Settings > Privacy & Security > Accessibility)

## Commands

```bash
keynote-cli validate instructions.json                   # Validate JSON structure
keynote-cli validate instructions.json --check-template  # Also verify masters in Keynote
keynote-cli build instructions.json                      # Build .key from instructions
keynote-cli build instructions.json --force              # Overwrite existing output
keynote-cli build instructions.json --keep-failed-output # Keep partial build on failure
keynote-cli inspect file.key                             # Dump slide structure as JSON
keynote-cli export file.key --output file.pdf            # Export to PDF
keynote-cli insert-equations equations.json              # Insert LaTeX equations via GUI
```

## Instruction format

Instructions are a JSON file with three required top-level fields:

```json
{
  "template": "template.key",
  "output": "out.key",
  "slides": [
    {
      "layout": "Header-Body",
      "content": {
        "header": "Introduction",
        "body": "Body text here"
      }
    }
  ]
}
```

Paths are resolved relative to the instruction file.

## Layouts

| Layout | Content keys |
|--------|-------------|
| `Title` | `title`, `subtitle` |
| `Header-Body` | `header`, `body` |
| `Header-TwoCol` | `header`, `left`, `right` |
| `Header-Body-TwoCol` | `header`, `body`, `left`, `right` |
| `Header-TwoCol-Body` | `header`, `left`, `right`, `body` |
| `Header-Body-TwoCol-Body` | `header`, `body_top`, `left`, `right`, `body_bottom` |

Slides can also include `images`, `text_boxes`, `overrides`, and `notes`. See `AGENTS.md` for the full schema.

## insert-equations

Replaces `[PLACEHOLDER]` tokens in an already-open Keynote deck with rendered LaTeX equations via the Insert > Equation GUI.

```bash
keynote-cli insert-equations equations.json
keynote-cli insert-equations equations.json --dry-run
keynote-cli insert-equations equations.json --render-timeout 30
```

Input is a JSON array:

```json
[
  {
    "slide": 24,
    "placeholder": "[EQ1]",
    "latex": "C(u,v) = (u^{-\\theta} + v^{-\\theta} - 1)^{-1/\\theta}",
    "label": "Clayton CDF"
  }
]
```

Limitations: Keynote must be open with the target document as front document. Multi-line environments (`\begin{array}`, `\begin{aligned}`, etc.) are not supported.

## Notes

- Run one command at a time — Keynote scripting is not concurrency-safe.
- The template (`template.key`) is authoritative for styling.
- Build failures show the failing slide number and layout.
- Source code: `keynote-cli` (entry point) imports `keynote_cli.py`.
