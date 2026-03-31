# keynote-cli

CLI for building Keynote presentations from a template `.key` file.

## Source code

- Entry point: `./keynote-cli`
- Main source: `./keynote_cli.py`

`keynote-cli` is a thin launcher that imports `keynote_cli.py`.

## Commands

```bash
./keynote-cli validate instructions.json
./keynote-cli validate instructions.json --check-template
./keynote-cli build instructions.json
./keynote-cli inspect file.key
./keynote-cli export file.key --output file.pdf
./keynote-cli insert-equations equations.json
```

## insert-equations

Replace `[PLACEHOLDER]` tokens in an **already-open** Keynote deck with
rendered LaTeX equations via the Insert > Equation GUI.

This uses macOS GUI scripting (System Events), so the calling process
(Terminal, etc.) must be listed under **System Settings > Privacy &
Security > Accessibility**.

### Input format

A JSON array of objects, each with:

| Key | Type | Description |
|---|---|---|
| `slide` | int | 1-based slide number |
| `placeholder` | string | Exact text to find and replace, e.g. `[EQ1]` |
| `latex` | string | LaTeX expression (standard JSON escaping for `\`) |
| `label` | string | *(optional)* Human-readable label for progress output |

Example `equations.json`:
```json
[
  {
    "slide": 24,
    "placeholder": "[EQ1]",
    "latex": "C(u,v) = (u^{-\\theta} + v^{-\\theta} - 1)^{-1/\\theta}",
    "label": "Clayton CDF"
  },
  {
    "slide": 25,
    "placeholder": "[EQ5]",
    "latex": "\\tau = 1 - 1/\\theta \\Rightarrow \\theta = \\frac{1}{1-\\tau}",
    "label": "Gumbel tau-theta"
  }
]
```

### Options

- `--render-timeout SEC` — Max seconds to wait for each equation to render
  (default: 15). Increase for complex expressions.
- `--dry-run` — Validate the JSON without touching Keynote.
- `--print-applescript` — Print the generated AppleScript instead of running it.

### Limitations

- Keynote must be open with the target document as the front document.
- The equation editor does **not** support multi-line environments
  (`\begin{array}`, `\begin{aligned}`, etc.). Use one equation per
  placeholder.
- If a placeholder is not found on the specified slide, the equation may
  be inserted as a floating object instead of inline.

## Layout names

Use these `layout` values in `instructions.json`:

- `Title`
- `Header-Body`
- `Header-TwoCol`
- `Header-Body-TwoCol`
- `Header-TwoCol-Body`
- `Header-Body-TwoCol-Body`

These map to master slides in `template.key`.

## Instruction format

```json
{
  "template": "template.key",
  "output": "out.key",
  "slides": [
    {
      "layout": "Title",
      "content": {
        "title": "Climate Risk and Intelligent Adaptation",
        "subtitle": "Acme Corp — Strategy Team"
      }
    },
    {
      "layout": "Header-Body",
      "content": {
        "header": "Introduction",
        "body": "Body text here"
      }
    },
    {
      "layout": "Header-TwoCol",
      "content": {
        "header": "Comparison",
        "left": "Left column",
        "right": "Right column"
      },
      "images": [
        {
          "file": "figure.png",
          "position": [1200, 700],
          "size": [300, 200]
        }
      ]
    }
  ]
}
```

Relative paths are resolved relative to the instruction file.

## Required top-level fields

- `template` — must point to an existing `.key` file
- `output` — must end with `.key`
- `slides` — non-empty array

## Required `content` keys by layout

All required content keys must be present, even if the value is an empty string.

- `Title` → `title`, `subtitle`
- `Header-Body` → `header`, `body`
- `Header-TwoCol` → `header`, `left`, `right`
- `Header-Body-TwoCol` → `header`, `body`, `left`, `right`
- `Header-TwoCol-Body` → `header`, `left`, `right`, `body`
- `Header-Body-TwoCol-Body` → `header`, `body_top`, `left`, `right`, `body_bottom`

Each content value can be:

- A **string** (plain text, newlines separate paragraphs — no indent info)
- An **array of paragraph objects** with explicit indent levels:

```json
"body": [
  {"text": "Main Point 1", "indent": 0},
  {"text": "Sub-point A", "indent": 1},
  {"text": "Sub-sub detail", "indent": 2},
  {"text": "Sub-point B", "indent": 1},
  {"text": "Main Point 2", "indent": 0}
]
```

Indent levels are 0-based (`0` = top-level bullet, `1` = first sub-bullet,
etc.). The visual bullet style and font size per level are controlled by the
slide layout/theme. Text is always set first, then indent levels are applied
— this is required by Keynote's paragraph model.

An array of plain strings (no `indent` key) is also accepted and treated as
all indent-0.

## Optional slide fields

### `notes`
Presenter notes for the slide.

### `images`
Insert images on the slide.

```json
{
  "file": "figure.png",
  "position": [960, 300],
  "size": [800, 500]
}
```

Rules:
- `file` is required and must exist
- `position` defaults to `[0, 0]`
- `size` is optional; if omitted, Keynote uses the natural size
- if present, `size` values must both be `> 0`

### `text_boxes`
Insert free text boxes.

```json
{
  "text": "Source: IPCC 2023",
  "position": [100, 1000],
  "size": [400, 30],
  "font": "HelveticaNeue",
  "fontSize": 20,
  "color": [128, 128, 128]
}
```

Rules:
- `text`, `position`, and `size` are required
- `font` must be a non-empty string
- `fontSize` must be `> 0`
- `color` accepts either 0-255 RGB or 0-65535 AppleScript RGB
- if omitted, free text boxes default to HelveticaNeue, 50pt, black

### `overrides`
Modify existing items on the created slide.

Supported `target` values:

- `defaultTitleItem`
- `defaultBodyItem`
- `content:<key>`
- `textItem:<n>`
- `image:<n>`
- `shape:<n>`

Example:

```json
{
  "target": "content:body",
  "position": [180, 920],
  "size": [1500, 120]
}
```

Supported override fields:

- `text`
- `position`
- `size`
- `font`
- `fontSize`
- `color`
- `opacity`
- `rotation`

Rules:
- every override must include `target`
- every override must change at least one property besides `target`
- `defaultBodyItem` is only valid on `Title` and `Header-Body`
- image targets cannot set text/font/color fields
- `opacity` must be between `0` and `100`
- `size` values must both be `> 0`

## Validation

`validate` checks JSON structure, required fields, layout/content compatibility, paths, and optional template master availability.

```bash
./keynote-cli validate example-instructions.json
./keynote-cli validate example-instructions.json --check-template
```

With `--check-template`, `keynote-cli` opens the template in Keynote and verifies that all master slides required by the instructions exist.

## Better runtime errors

`build` now wraps each slide in AppleScript error handling. If Keynote fails while building a slide, the error message includes the slide number and layout, e.g.:

```text
Slide 4 (Header-TwoCol) failed: ...
```

If a build fails after output creation, the partial `.key` file is removed unless you pass:

```bash
./keynote-cli build instructions.json --keep-failed-output
```

## Inspect output

`inspect` prints JSON with:

- slide dimensions
- master slide names
- slide count
- each slide's master name
- filtered text items
- images
- shapes
- presenter notes

Example:

```bash
./keynote-cli inspect out.key | jq '.slides[0]'
```

## Export

```bash
./keynote-cli export out.key --output out.pdf
```

If `--output` is omitted, the PDF is written next to the `.key` file with the same basename.

## Notes

- Uses AppleScript via `osascript`, so it only works on macOS with Keynote installed.
- Run one `keynote-cli` command at a time. Keynote scripting is not reliably concurrency-safe.
- Layout placeholder labels visible in Keynote's slide-layout editor are not accessible via AppleScript. The CLI targets items by fixed layout mappings and item indices.
- For the two-column layouts, the underlying Keynote item order is counterintuitive: right column comes before left column internally. The CLI hides this behind `left` and `right` content keys.
