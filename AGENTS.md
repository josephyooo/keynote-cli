# keynote-cli — Technical Reference

## Script command reference

### open

```
open TEMPLATE --output OUTPUT [--force]
```

- `TEMPLATE` must be an existing `.key` file.
- `OUTPUT` must end with `.key` and be different from `TEMPLATE`.
- `--force` allows overwriting existing output.
- Paths are resolved relative to the script file's directory.

### add-slide

```
add-slide --master NAME
```

- `NAME` is the exact Keynote master slide name (e.g. `"Title & Bullets"`).
- Slides are numbered sequentially starting from 1 in the order they are added.

### set-text

```
set-text --slide N --target TARGET TEXT [--indents I1,I2,...]
```

- `N` is 1-based slide number (among added slides).
- `TARGET`: `defaultTitleItem`, `defaultBodyItem`, or `textItem:N`.
- `TEXT`: use `\n` for newlines, `\t` for tabs, `\\` for backslashes.
- `--indents`: comma-separated 0-based indent levels, one per line in `TEXT`. Length must match the number of `\n`-delimited lines.

### set-notes

```
set-notes --slide N TEXT
```

### add-image

```
add-image --slide N --file PATH --position X,Y [--size W,H]
```

- `PATH` must exist (resolved relative to script file).
- `--position` defaults to `0,0` if omitted.
- `--size` is optional; if omitted, Keynote uses the image's natural size.
- If provided, both width and height must be > 0.

### add-text-box

```
add-text-box --slide N --text TEXT --position X,Y --size W,H [--font F] [--font-size S] [--color R,G,B]
```

- `--text`, `--position`, `--size` are required.
- Defaults: font `HelveticaNeue`, size `50`, color `0,0,0` (black).
- Color accepts 0-255 RGB (auto-scaled to 0-65535).

### override

```
override --slide N --target TARGET [--text T] [--position X,Y] [--size W,H] [--font F] [--font-size S] [--color R,G,B] [--opacity O] [--rotation R]
```

- `TARGET`: `defaultTitleItem`, `defaultBodyItem`, `textItem:N`, `image:N`, `shape:N`.
- At least one property must be specified besides `--target`.
- Image targets (`image:N`) cannot set `--text`, `--font`, `--font-size`, or `--color`.
- `--opacity`: 0 to 100.
- `--size` values must both be > 0.

### delete-slides

```
delete-slides RANGE
```

- `RANGE` is `START-END` (e.g. `1-7`) or a single number (e.g. `5`).
- 1-based. Both start and end are inclusive.

### duplicate-slide

```
duplicate-slide --slide N [--to M]
```

- `N` is 1-based slide index in the document (including template slides).
- `--to M`: insert the duplicate after slide M. If omitted, it goes immediately after the source.

### move-slide

```
move-slide --slide N --to M
```

- Moves slide N to position M (1-based).
- `--to 1` moves to the beginning.

### replace-text

```
replace-text --find "X" --replace "Y" [--slide N]
```

- Finds and replaces text across all text items on all slides.
- `--slide N`: limit replacement to a single slide.
- Uses AppleScript `text item delimiters` for substring replacement.

### add-shape

```
add-shape --slide N --position X,Y --size W,H [--text T] [--rotation D] [--opacity O]
```

- Creates a new shape on the specified slide.
- `--position` and `--size` are required. Size values must be > 0.
- Fill color is not settable via AppleScript (shapes get Keynote's default styling).
- `--text`: optional text inside the shape.
- `--rotation`: degrees (0-359).
- `--opacity`: 0-100.

### skip-slide

```
skip-slide --slide N
```

- Hides slide N from presentation playback (`set skipped to true`).

### unskip-slide

```
unskip-slide --slide N
```

- Unhides slide N (`set skipped to false`).

### set-style

```
set-style --slide N --target TARGET [--bold|--no-bold] [--italic|--no-italic] [--underline|--no-underline]
```

- Sets bold/italic/underline on the entire text of the target element.
- Use `--no-bold`, `--no-italic`, `--no-underline` to explicitly unset.
- At least one style flag is required.

### add-line

```
add-line --slide N --from X,Y --to X,Y
```

- Creates a line from the start point to the end point on slide N.

### duplicate-shape

```
duplicate-shape --slide N --index I --to-slide M
```

- Copies shape I from slide N to slide M.
- Useful as a workaround for styled shapes: pre-style a shape in your template, then duplicate it.

### delete-shape

```
delete-shape --slide N --index I
```

- Deletes shape I (1-based) from slide N.

### delete-image

```
delete-image --slide N --index I
```

- Deletes image I (1-based) from slide N.
- Since image source is read-only, delete + re-insert is the way to swap images.

### set-master

```
set-master --slide N --master NAME
```

- Changes the master (base slide) of slide N.
- This is the scriptable way to change slide backgrounds — use a master with the desired background.

### set-theme

```
set-theme --theme NAME
```

- Applies a theme to the entire document, changing all master slides at once.

### delete-slides

```
delete-slides RANGE
```

- `RANGE` is `START-END` (e.g. `1-7`) or a single number (e.g. `5`).
- 1-based. Both start and end are inclusive.

### save

```
save
```

Saves and closes the document.

## Execution order

Slide-creation commands (`add-slide`, `set-text`, `add-image`, etc.) are batched and executed first. Document-level commands (`duplicate-slide`, `move-slide`, `replace-text`, `add-shape`, `set-master`, `delete-slides`) execute after all slide creation, in script order. This means document-level commands should reference slide indices as they will exist after all new slides have been added.

## Target notation

| Notation | AppleScript equivalent |
|----------|----------------------|
| `defaultTitleItem` | `default title item of slide` |
| `defaultBodyItem` | `default body item of slide` |
| `textItem:N` | `text item N of slide` |
| `image:N` | `image N of slide` |
| `shape:N` | `shape N of slide` |

## inspect output format

```json
{
  "file": "/path/to/file.key",
  "dimensions": [1920, 1080],
  "slideCount": 5,
  "masters": ["Title", "Title & Bullets"],
  "slides": [
    {
      "index": 1,
      "master": "Title",
      "notes": "",
      "textItems": [
        {"index": 1, "text": "...", "position": [140, 181], "size": [1640, 366]}
      ],
      "images": [],
      "shapes": []
    }
  ]
}
```

## Export formats

`keynote-cli export file.key --format FORMAT [--output PATH]`

| Format | Keynote export type | Default suffix | Notes |
|--------|-------------------|----------------|-------|
| `pdf` | PDF | `.pdf` | Default |
| `png` | slide images (PNG) | `.png` | Individual slide images |
| `jpeg` | slide images (JPEG) | `.jpeg` | Individual slide images |
| `pptx` | Microsoft PowerPoint | `.pptx` | |
| `html` | HTML | `.html` | |

## insert-equations input format

A JSON array of objects:

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `slide` | int | yes | 1-based slide number |
| `placeholder` | string | yes | Exact text to find and replace, e.g. `[EQ1]` |
| `latex` | string | yes | LaTeX expression |
| `label` | string | no | Human-readable label for progress output |

Options: `--render-timeout SEC` (default 15), `--dry-run`, `--print-applescript`.

Limitations: Keynote must be open with target document as front document. Multi-line environments (`\begin{array}`, `\begin{aligned}`, etc.) are not supported.

## Performance

`keynote-cli run` batches operations into groups of 20 slides per `osascript` call. The `delete-slides` command is appended to the last batch. Each batch is wrapped in error handling that reports the failing slide number and master name.
