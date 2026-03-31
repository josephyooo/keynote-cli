# Keynote CLI Tool — Plan

> **Status:** Historical implementation/background document. The `keynote-cli` MVP described here is already implemented. For current repo-wide priorities, see `../ACTION_PLAN.md`.

## Goal
A CLI tool that programmatically creates Keynote presentations from a template `.key` file. It reads a JSON instruction file describing slides, creates them from the template's master layouts, populates text, inserts images, and saves the result.

---

## Architecture: AppleScript via `osascript`

**AppleScript called from a shell wrapper**, not JXA or Swift.

- **JXA is disqualified**: `masterSlides` is completely broken (`"Can't convert types"` on every access). Cannot list, select, or create slides from specific layouts.
- **Swift deferred**: No benefit at this stage. AppleScript handles everything needed. No compilation step.

### Confirmed Capabilities
| Feature | Status | Notes |
|---------|--------|-------|
| Open/save `.key` files | ✅ | |
| List master slide names | ✅ | |
| Create slide from master layout | ✅ | `make new slide with properties {base slide: master slide "X"}` |
| Set text via `defaultTitleItem` | ✅ | |
| Set text via `defaultBodyItem` | ✅ | Only works for Title and Title & Bullets layouts |
| Set text via `text item N` | ✅ | All items are writable |
| Add free text items | ✅ | Position, size, font, color |
| Set font/size/color on text | ✅ | Must use font variant names (e.g. `"Helvetica-Bold"`) |
| Add images | ✅ | `make new image with properties {file:..., position:..., width:..., height:...}` |
| Position/resize elements | ✅ | |
| Duplicate/delete slides | ✅ | |
| Export to PDF | ✅ | |
| Set text alignment | ❌ | Must be set in master slide |
| Rename/delete master slides | ❌ | Must be preset in template |

---

## Template (`template.key`)

### Document: 1920×1080

### Design System
| Role | Font | Size | Color (AppleScript RGB) |
|------|------|------|------------------------|
| Title (cover slide) | HelveticaNeue-Medium | 112pt | 42010, 37849, 22169 (gold `#a49356`) |
| Subtitle (cover slide) | HelveticaNeue | 54pt | 0, 0, 0 (black) |
| Header (content slides) | HelveticaNeue-Medium | 84pt | 42010, 37849, 22169 (gold) |
| Body text | HelveticaNeue | 50pt | 0, 0, 0 (black) |

### Master Slides Available
| Master Name | Used For |
|---|---|
| `Title` | Cover, section titles, Questions slide |
| `Title & Bullets` | Header + full-width body, References |
| `Title & Bullets Two-Column` | Header + two columns |
| `Title & Bullets Body over Two-Column` | Header + body row + two columns |
| `Title & Bullets Body under Two-Column` | Header + two columns + body row |
| `Title & Bullets Body around Two-Column` | Header + body row + two columns + body row |

### Layout → Text Item Index Map
When a slide is created from a master, text items appear in a fixed order. Items after the content items are duplicates (master-inherited) or hidden zero-size placeholders — **ignore them**.

**Title** (4 text items, use first 2):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Title | 140, 181 | 1640×366 |
| 2 | Subtitle | 140, 557 | 1640×125 |

Access via: `defaultTitleItem` (title), `defaultBodyItem` (subtitle)

**Title & Bullets** (4 text items, use first 2):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654×180 |
| 2 | Body | 133, 190 | 1654×732 |

Access via: `defaultTitleItem` (header), `defaultBodyItem` (body)

**Title & Bullets Two-Column** (5 text items, use first 3):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654×180 |
| 2 | Right column | 960, 190 | 827×732 |
| 3 | Left column | 133, 190 | 827×732 |

Access via: `defaultTitleItem` (header), `text item 2` (right), `text item 3` (left)
> ⚠️ **Note**: Right column is index 2, left is index 3 (counterintuitive)

**Title & Bullets Body over Two-Column** (6 text items, use first 4):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654×180 |
| 2 | Right column | 960, 268 | 827×732 |
| 3 | Body (above cols) | 133, 190 | 1654×76 |
| 4 | Left column | 133, 268 | 827×732 |

Access via: `defaultTitleItem` (header), `text item 3` (body), `text item 4` (left), `text item 2` (right)

**Title & Bullets Body under Two-Column** (6 text items, use first 4):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654×180 |
| 2 | Right column | 960, 190 | 827×749 |
| 3 | Body (below cols) | 133, 941 | 1654×132 |
| 4 | Left column | 133, 190 | 827×749 |

Access via: `defaultTitleItem` (header), `text item 4` (left), `text item 2` (right), `text item 3` (body)

**Title & Bullets Body around Two-Column** (7 text items, use first 5):
| Index | Role | Position | Size |
|-------|------|----------|------|
| 1 | Header | 133, 28 | 1654×180 |
| 2 | Right column | 960, 268 | 827×671 |
| 3 | Body (above cols) | 133, 190 | 1654×76 |
| 4 | Left column | 133, 268 | 827×671 |
| 5 | Body (below cols) | 133, 941 | 1654×132 |

### Template Content Slides (1–7)
Slides 1–7 in the template are example/reference slides. They will be deleted from output files — the CLI creates fresh slides from the master layouts.

---

## CLI Design

### Commands

```bash
# Build a full presentation from a JSON instruction file
keynote-cli build <instructions.json>

# Inspect a .key file — dump structure as JSON (slide count, text items, images)
keynote-cli inspect <file.key>

# Export a .key file to PDF
keynote-cli export <file.key> [--output path.pdf]
```

### `build` — Instruction JSON Format

```json
{
  "template": "/absolute/path/to/template.key",
  "output": "/absolute/path/to/output.key",
  "slides": [
    {
      "layout": "Title",
      "content": {
        "title": "Presentation Title",
        "subtitle": "Author, Affiliation"
      }
    },
    {
      "layout": "Header-Body",
      "content": {
        "header": "Section Header",
        "body": "Body text here\n• Bullet point"
      }
    },
    {
      "layout": "Header-TwoCol",
      "content": {
        "header": "Two Column Slide",
        "left": "Left column text",
        "right": "Right column text"
      }
    },
    {
      "layout": "Header-Body-TwoCol",
      "content": {
        "header": "Body Over Columns",
        "body": "Full-width body text",
        "left": "Left column",
        "right": "Right column"
      }
    },
    {
      "layout": "Header-TwoCol-Body",
      "content": {
        "header": "Columns Over Body",
        "left": "Left column",
        "right": "Right column",
        "body": "Full-width body text"
      }
    },
    {
      "layout": "Header-Body-TwoCol-Body",
      "content": {
        "header": "Body Around Columns",
        "body_top": "Top body text",
        "left": "Left column",
        "right": "Right column",
        "body_bottom": "Bottom body text"
      }
    }
  ]
}
```

Slides may also include images:
```json
{
  "layout": "Header-Body",
  "content": { "header": "Results", "body": "See figure:" },
  "images": [
    {
      "file": "/absolute/path/to/figure.png",
      "position": [960, 300],
      "size": [800, 500]
    }
  ]
}
```

### Layout Name → Master Slide Mapping
| JSON `layout` value | Keynote master slide name |
|---|---|
| `Title` | `Title` |
| `Header-Body` | `Title & Bullets` |
| `Header-TwoCol` | `Title & Bullets Two-Column` |
| `Header-Body-TwoCol` | `Title & Bullets Body over Two-Column` |
| `Header-TwoCol-Body` | `Title & Bullets Body under Two-Column` |
| `Header-Body-TwoCol-Body` | `Title & Bullets Body around Two-Column` |

### Layout Content Key → Text Item Mapping
| Layout | Content key | Access method |
|---|---|---|
| **Title** | `title` | `default title item` |
| | `subtitle` | `default body item` |
| **Header-Body** | `header` | `default title item` |
| | `body` | `default body item` |
| **Header-TwoCol** | `header` | `default title item` |
| | `left` | `text item 3` |
| | `right` | `text item 2` |
| **Header-Body-TwoCol** | `header` | `default title item` |
| | `body` | `text item 3` |
| | `left` | `text item 4` |
| | `right` | `text item 2` |
| **Header-TwoCol-Body** | `header` | `default title item` |
| | `left` | `text item 4` |
| | `right` | `text item 2` |
| | `body` | `text item 3` |
| **Header-Body-TwoCol-Body** | `header` | `default title item` |
| | `body_top` | `text item 3` |
| | `left` | `text item 4` |
| | `right` | `text item 2` |
| | `body_bottom` | `text item 5` |

### `inspect` — Output Format

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
        {"index": 1, "text": "Presentation Title", "position": [140, 181], "size": [1640, 366]},
        {"index": 2, "text": "Author Name", "position": [140, 557], "size": [1640, 125]}
      ],
      "images": []
    }
  ]
}
```

---

## Implementation Plan

### Phase 1: `keynote-cli build`
- [ ] Shell script wrapper that reads JSON with `jq` or Python, generates AppleScript
- [ ] Single `osascript` call that:
  1. Opens template
  2. Counts existing template slides (to delete later)
  3. For each instruction: creates slide from master, populates text items per mapping
  4. Handles `Header-Body-TwoCol-Body` bottom body workaround
  5. Adds images where specified
  6. Deletes original template slides
  7. Saves to output path
- [ ] Error handling: validate layout names, required content keys

### Phase 2: `keynote-cli inspect`
- [ ] Reads a `.key` file, outputs JSON describing all slides, text items, images
- [ ] Filters out duplicate/hidden text items (zero-size, or position matching a previous item)

### Phase 3: `keynote-cli export`
- [ ] Simple wrapper around Keynote's export command

### Phase 4: Robustness
- [ ] Handle edge cases: missing content keys (leave placeholder empty), missing image files
- [ ] Validate JSON input before generating AppleScript
- [ ] Meaningful error messages on failure

---

## Known Issues & Workarounds

1. **Two-Column item order is counterintuitive**: Right column is `text item 2`, left is `text item 3` (or `text item 4` in body+col layouts). This is baked into the master slide and consistent.

2. **`defaultBodyItem` is broken for custom masters**: For all three custom two-column masters and the base two-column master, `defaultBodyItem` points to a hidden 0×0 placeholder. Must use `text item N` by index instead.

3. **Duplicate text items**: Every slide shows master-inherited duplicates. The CLI must only target the first N content items per layout (items after that are duplicates at the same positions or hidden 0×0 placeholders).

4. **Text formatting on master placeholders**: Setting font/size/color works on freshly created slides. The template's master slides already define the correct formatting, so the CLI only needs to set `object text` — formatting is inherited automatically.

5. **Placeholder labels not exposed**: The instructional text shown in Keynote's Edit Slide Layout view (e.g. "Title Text", custom labels) is not accessible via AppleScript. Text items are identified by index only.
