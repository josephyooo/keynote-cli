# keynote-cli — Technical Reference

This document contains the schemas, mappings, and rules needed to generate valid `keynote-cli` instruction files.

## Template specifications

Document size: **1920 x 1080**

### Design system

| Role | Font | Size | Color (RGB 0-65535) | Color (hex) |
|------|------|------|---------------------|-------------|
| Title (cover) | HelveticaNeue-Medium | 112pt | 42010, 37849, 22169 | `#a49356` |
| Subtitle (cover) | HelveticaNeue | 54pt | 0, 0, 0 | black |
| Header (content) | HelveticaNeue-Medium | 84pt | 42010, 37849, 22169 | `#a49356` |
| Body text | HelveticaNeue | 50pt | 0, 0, 0 | black |

### Layout name to Keynote master slide mapping

| JSON `layout` value | Keynote master slide name |
|---|---|
| `Title` | `Title` |
| `Header-Body` | `Title & Bullets` |
| `Header-TwoCol` | `Title & Bullets Two-Column` |
| `Header-Body-TwoCol` | `Title & Bullets Body over Two-Column` |
| `Header-TwoCol-Body` | `Title & Bullets Body under Two-Column` |
| `Header-Body-TwoCol-Body` | `Title & Bullets Body around Two-Column` |

### Layout text item index map

When a slide is created from a master, text items appear in a fixed order. Items after the documented content items are master-inherited duplicates or hidden zero-size placeholders — ignore them.

**Title** (use first 2 of 4 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Title | 140, 181 | 1640 x 366 |
| 2 | Subtitle | 140, 557 | 1640 x 125 |

Access: `defaultTitleItem` → title, `defaultBodyItem` → subtitle

**Header-Body** (use first 2 of 4 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654 x 180 |
| 2 | Body | 133, 190 | 1654 x 732 |

Access: `defaultTitleItem` → header, `defaultBodyItem` → body

**Header-TwoCol** (use first 3 of 5 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654 x 180 |
| 2 | Right column | 960, 190 | 827 x 732 |
| 3 | Left column | 133, 190 | 827 x 732 |

Access: `defaultTitleItem` → header, `text item 3` → left, `text item 2` → right

**Header-Body-TwoCol** (use first 4 of 6 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654 x 180 |
| 2 | Right column | 960, 268 | 827 x 732 |
| 3 | Body (above cols) | 133, 190 | 1654 x 76 |
| 4 | Left column | 133, 268 | 827 x 732 |

Access: `defaultTitleItem` → header, `text item 3` → body, `text item 4` → left, `text item 2` → right

**Header-TwoCol-Body** (use first 4 of 6 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654 x 180 |
| 2 | Right column | 960, 190 | 827 x 749 |
| 3 | Body (below cols) | 133, 941 | 1654 x 132 |
| 4 | Left column | 133, 190 | 827 x 749 |

Access: `defaultTitleItem` → header, `text item 4` → left, `text item 2` → right, `text item 3` → body

**Header-Body-TwoCol-Body** (use first 5 of 7 text items):

| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654 x 180 |
| 2 | Right column | 960, 268 | 827 x 671 |
| 3 | Body top (above cols) | 133, 190 | 1654 x 76 |
| 4 | Left column | 133, 268 | 827 x 671 |
| 5 | Body bottom (below cols) | 133, 941 | 1654 x 132 |

Access: `defaultTitleItem` → header, `text item 3` → body_top, `text item 4` → left, `text item 2` → right, `text item 5` → body_bottom

## Instruction JSON schema

### Top-level fields (all required)

| Field | Type | Rules |
|-------|------|-------|
| `template` | string | Must point to an existing `.key` file |
| `output` | string | Must end with `.key` |
| `slides` | array | Non-empty |

### Slide object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `layout` | string | yes | One of the 6 layout names |
| `content` | object | yes | Keys depend on layout (see below) |
| `notes` | string | no | Presenter notes |
| `images` | array | no | Image objects |
| `text_boxes` | array | no | Free text box objects |
| `overrides` | array | no | Element override objects |

### Required content keys by layout

All required keys must be present, even if the value is an empty string.

| Layout | Required keys |
|--------|--------------|
| `Title` | `title`, `subtitle` |
| `Header-Body` | `header`, `body` |
| `Header-TwoCol` | `header`, `left`, `right` |
| `Header-Body-TwoCol` | `header`, `body`, `left`, `right` |
| `Header-TwoCol-Body` | `header`, `left`, `right`, `body` |
| `Header-Body-TwoCol-Body` | `header`, `body_top`, `left`, `right`, `body_bottom` |

### Content value format

Each content value can be either:

**A string** — newlines separate paragraphs, no indent control:

```json
"body": "Main point\nSub-point\nAnother point"
```

**An array of paragraph objects** — explicit indent levels:

```json
"body": [
  {"text": "Main Point 1", "indent": 0},
  {"text": "Sub-point A", "indent": 1},
  {"text": "Sub-sub detail", "indent": 2},
  {"text": "Main Point 2", "indent": 0}
]
```

Indent levels are 0-based. Visual bullet style and font size per level are controlled by the template. An array of plain strings (no `indent` key) is accepted and treated as all indent-0.

### Image object

| Field | Type | Required | Rules |
|-------|------|----------|-------|
| `file` | string | yes | Must exist (resolved relative to instruction file) |
| `position` | [x, y] | no | Defaults to `[0, 0]` |
| `size` | [w, h] | no | If present, both values must be > 0. If omitted, Keynote uses natural size. |

### Text box object

| Field | Type | Required | Default |
|-------|------|----------|---------|
| `text` | string | yes | |
| `position` | [x, y] | yes | |
| `size` | [w, h] | yes | |
| `font` | string | no | `HelveticaNeue` |
| `fontSize` | number | no | `50` |
| `color` | [r, g, b] | no | `[0, 0, 0]` (black) |

- `font` must be a non-empty string
- `fontSize` must be > 0
- `color` accepts either 0-255 RGB or 0-65535 AppleScript RGB

### Override object

| Field | Type | Required |
|-------|------|----------|
| `target` | string | yes |
| `text` | string | no |
| `position` | [x, y] | no |
| `size` | [w, h] | no |
| `font` | string | no |
| `fontSize` | number | no |
| `color` | [r, g, b] | no |
| `opacity` | number | no |
| `rotation` | number | no |

Rules:
- Every override must include `target` and change at least one other property.
- `size` values must both be > 0.
- `opacity` must be between 0 and 100.
- `defaultBodyItem` target is only valid on `Title` and `Header-Body` layouts.
- Image targets (`image:N`) cannot set text/font/color fields.

Supported `target` values:
- `defaultTitleItem`
- `defaultBodyItem`
- `content:<key>` (e.g. `content:body`, `content:left`)
- `textItem:<n>` (1-based)
- `image:<n>` (1-based)
- `shape:<n>` (1-based)

### insert-equations input

A JSON array of objects:

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `slide` | int | yes | 1-based slide number |
| `placeholder` | string | yes | Exact text to find and replace, e.g. `[EQ1]` |
| `latex` | string | yes | LaTeX expression |
| `label` | string | no | Human-readable label for progress output |

Options: `--render-timeout SEC` (default 15), `--dry-run`, `--print-applescript`.

## inspect output format

```json
{
  "file": "/path/to/file.key",
  "dimensions": [1920, 1080],
  "slideCount": 5,
  "slides": [
    {
      "index": 1,
      "master": "Title",
      "textItems": [
        {"index": 1, "text": "...", "position": [140, 181], "size": [1640, 366]}
      ],
      "images": [],
      "shapes": [],
      "presenterNotes": ""
    }
  ]
}
```

## Template slides 1-7

The template contains 7 example/reference slides. They are automatically deleted from output files after new slides are created.
