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
| Export to movie | done | `--format movie` |
| Set document theme | done | `set-theme` script command |
| Start presentation | done | `present` standalone command with `--from N` |
| Playback control (stop/advance/rewind) | not yet | `stop`, `show next/previous` |
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
| Add table | done | `add-table` with row/column count, optional position/size |
| Set cell value | done | `set-cell` by row/column index |
| Get cell value | not yet | |
| Get/set cell formula | not yet | |
| Add row/column | done | `add-row`, `add-col` |
| Delete row/column | done | `delete-row`, `delete-col` |
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

### Master slides & themes

| Operation | Status | Notes |
|-----------|--------|-------|
| List master slide names | done | via `inspect` |
| Apply master to a slide | done | `set-master` |
| Apply theme to document | done | `set-theme` |
| List available themes | not yet | `get name of every theme` |
| Create document from theme | not yet | `make new document with properties {document theme: theme "X"}` |
| Edit text style on a master | done | `override` / `set-style` targeting master slide items |
| Edit shapes on a master | done | `add-shape` etc. work inside `tell master slide` |
| Set master background fill | no | `background fill type` is read-only |
| Add/delete a master slide | no | Not in scripting dictionary |
| Rename a master slide | no | `name` is read-only |
| Save as `.kth` theme file | no | GUI-only (`File > Save Theme…` via System Events) |
| Modify placeholder tags | no | Not exposed to AppleScript |

### Hyperlinks

Keynote's scripting dictionary has no native hyperlink class. All link operations require GUI scripting.

| Operation | Status | Notes |
|-----------|--------|-------|
| Add URL hyperlink to text | no | GUI-only: select text, Cmd+K, type URL, Return |
| Add slide navigation link to object | no | GUI-only: right-click > Add Link > Slide (locale-dependent menu item) |
| Read existing URL from text | no | `get URL of word N` exists but is inconsistent across Keynote versions |
| Remove a link | no | GUI-only: select text, Cmd+K, clear URL |

If implemented, the approach would mirror `insert-equations`: activate Keynote, drive the UI via System Events, require Accessibility permissions. Slide navigation links are especially fragile because they rely on context menu item names that change with locale. Best done as a batch pass at the end of a build, not interleaved with other operations.

### Transitions & builds

| Operation | Status | Notes |
|-----------|--------|-------|
| Set slide transition | done | `set-transition` with style and optional duration |
| Get slide transition | not yet | |
| Get build order | not yet | Read-heavy; write is mostly GUI-only |

## Automation mechanisms

All script commands use standard `tell application "Keynote"` AppleScript, which works headlessly (Keynote can be in the background or not running — it launches automatically). The one exception is `insert-equations`, which uses GUI scripting via System Events.

| Command | Mechanism | Keynote must be open? | Needs frontmost window? | Accessibility permissions? |
|---------|-----------|----------------------|------------------------|---------------------------|
| `run` (all script commands) | Keynote scripting | No (opens automatically) | No | No |
| `inspect` | Keynote scripting | No (opens automatically) | No | No |
| `export` | Keynote scripting | No (opens automatically) | No | No |
| `present` | Keynote scripting | No (opens automatically) | Yes (`activate` is called) | No |
| `insert-equations` | Keynote + System Events | Yes (document must be open) | Yes (`activate` is called) | Yes |

`insert-equations` drives the GUI: it focuses Keynote, uses Cmd+F to find placeholder text, clicks Insert > Equation, types LaTeX into the editor, and waits for the renderer. This requires the calling process (Terminal, etc.) to be listed under System Settings > Privacy & Security > Accessibility.

### Operations requiring GUI scripting

These cannot be done headlessly and need System Events + Accessibility permissions:
- Equation insertion (implemented)
- Hyperlinks — URL links (Cmd+K) and slide navigation links (context menu)
- Setting shape/text fill color
- Setting slide/master background to an arbitrary color
- Saving as `.kth` theme file (`File > Save Theme…`)
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
