# keynote-cli — Plan

## Architecture

**AppleScript via `osascript`**, called from Python. Scripts are compiled into batched AppleScript for performance.

- **JXA rejected**: `masterSlides` is completely broken (`"Can't convert types"` on every access).
- **Swift deferred**: No benefit at this stage.

## Capability matrix

### Document

| Operation | Status | Notes |
|-----------|--------|-------|
| Open/save/close | done | |
| Export to PDF/PNG/JPEG/PPTX/HTML | done | `--format` flag on `export` command |
| Export to movie | not yet | Keynote supports `export ... as QuickTime movie` |
| Set document theme | done | `set-theme` script command |
| Presentation playback (start/stop/advance) | not yet | `start`, `stop`, `show next/previous` |
| Get slide count | done | via `inspect` |
| Get/set current slide | not yet | |

### Slides

| Operation | Status | Notes |
|-----------|--------|-------|
| Add slide from master | done | |
| Delete slides by range | done | |
| Duplicate slide | done | Optionally place after a specific slide |
| Move/reorder slides | done | |
| Skip/unskip slide | done | `skip-slide` / `unskip-slide` |
| Set master (base slide) | done | Scriptable way to change backgrounds |
| Get/set presenter notes | done | |
| Inspect slide structure | done | JSON output |

### Text

| Operation | Status | Notes |
|-----------|--------|-------|
| Set text on any item | done | via `defaultTitleItem`, `defaultBodyItem`, `textItem:N` |
| Set paragraph indent levels | done | via `--indents` |
| Add free text boxes | done | Position, size, font, color |
| Set font/size/color on text | done | Must use font variant names (e.g. `"Helvetica-Bold"`) |
| Find/replace text | done | Across all slides or a single slide |
| Override element properties | done | text, position, size, font, color, opacity, rotation |
| Set text alignment | no | Must be set in master slide |
| Bold/italic/underline styling | done | `set-style` script command |

### Shapes

| Operation | Status | Notes |
|-----------|--------|-------|
| Add shape | done | Position, size, text, rotation, opacity |
| Add line | done | `add-line` with start/end points |
| Set shape text/position/size/rotation/opacity | done | Via `add-shape` or `override` |
| Duplicate shape | done | `duplicate-shape` — workaround for copying pre-styled shapes |
| Delete shape | done | `delete-shape` by index |
| Set shape fill color | no | `background fill type` is read-only in AppleScript |
| Set shape z-order | no | Requires GUI scripting |

### Images

| Operation | Status | Notes |
|-----------|--------|-------|
| Add image | done | From file path, with position/size |
| Set position/size/opacity/rotation | done | Via `override` |
| Delete image | done | `delete-image` by index |
| Swap image source | no | `file` and `file name` are read-only — must delete and re-insert |

### Tables

| Operation | Status | Notes |
|-----------|--------|-------|
| Add table | not yet | `make new table` with row/column count |
| Get/set cell value | not yet | By row/column index |
| Get/set cell formula | not yet | |
| Add/delete rows/columns | not yet | |
| Merge/split cells | not yet | |

### Equations

| Operation | Status | Notes |
|-----------|--------|-------|
| Insert LaTeX equation | done | GUI scripting via System Events — requires Accessibility permissions |

### Media

| Operation | Status | Notes |
|-----------|--------|-------|
| Add movie/audio | not yet | `make new movie/audio clip` |
| Set position/size/autoplay/loop | not yet | |

### Transitions & builds

| Operation | Status | Notes |
|-----------|--------|-------|
| Get/set slide transition | not yet | Transition style and duration are partially scriptable |
| Get build order | not yet | Read-heavy; write is mostly GUI-only |

## Automation mechanisms

All script commands use standard `tell application "Keynote"` AppleScript, which works headlessly (Keynote can be in the background or not running — it launches automatically). The one exception is `insert-equations`, which uses GUI scripting via System Events.

| Command | Mechanism | Keynote must be open? | Needs frontmost window? | Accessibility permissions? |
|---------|-----------|----------------------|------------------------|---------------------------|
| `run` (all script commands) | Keynote scripting | No (opens automatically) | No | No |
| `inspect` | Keynote scripting | No (opens automatically) | No | No |
| `export` | Keynote scripting | No (opens automatically) | No | No |
| `insert-equations` | Keynote + System Events | Yes (document must be open) | Yes (`activate` is called) | Yes |

`insert-equations` drives the GUI: it focuses Keynote, uses Cmd+F to find placeholder text, clicks Insert > Equation, types LaTeX into the editor, and waits for the renderer. This requires the calling process (Terminal, etc.) to be listed under System Settings > Privacy & Security > Accessibility.

### Operations requiring GUI scripting

These cannot be done headlessly and need System Events + Accessibility permissions:
- Equation insertion (implemented)
- Setting shape/text fill color
- Setting slide background to an arbitrary color
- Build/animation write operations
- Z-order changes (send to back/front)

## Known issues

1. **Duplicate text items on every slide**: Master-inherited duplicates appear. When targeting by index, know that items after the content items may be duplicates or hidden 0x0 placeholders.

2. **`defaultBodyItem` broken on some masters**: For many custom masters, `defaultBodyItem` points to a hidden 0x0 placeholder. Use `textItem:N` instead.

3. **Placeholder labels not exposed**: The instructional text in Keynote's Edit Slide Layout view is not accessible via AppleScript. Items are identified by index only.

4. **Text formatting inherited from master**: The template's masters define formatting. Setting `object text` inherits formatting automatically.

5. **Image source is read-only**: `file` and `file name` on images are read-only. To change an image, delete it and re-insert.

## Future direction

Standalone mutation commands (e.g. `keynote-cli set-text --document front ...` or `keynote-cli set-text --file out.key ...`) could allow individual operations outside a script. This would require specifying the document source — either an already-open window (`--document front`) or a file to open (`--file path.key`). Currently, mutation commands only work inside script files run via `keynote-cli run`.
