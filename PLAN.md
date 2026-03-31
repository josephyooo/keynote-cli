# keynote-cli — Plan

## Architecture

**AppleScript via `osascript`**, called from Python.

- **JXA rejected**: `masterSlides` is completely broken (`"Can't convert types"` on every access). Cannot list, select, or create slides from specific layouts.
- **Swift deferred**: No benefit at this stage. AppleScript handles everything needed. No compilation step.

## Capability status

| Feature | Status | Notes |
|---------|--------|-------|
| Open/save `.key` files | done | |
| List master slide names | done | |
| Create slide from master layout | done | `make new slide with properties {base slide: master slide "X"}` |
| Set text on placeholders | done | via `defaultTitleItem`, `defaultBodyItem`, `text item N` |
| Set paragraph indent levels | done | via `set indent level of paragraph N` |
| Add free text items | done | Position, size, font, color |
| Set font/size/color on text | done | Must use font variant names (e.g. `"Helvetica-Bold"`) |
| Add images | done | `make new image with properties {file:..., position:..., width:..., height:...}` |
| Position/resize elements | done | |
| Override element properties | done | Supports opacity, rotation, text, position, size, font, color |
| Duplicate/delete slides | done | |
| Export to PDF | done | |
| Inspect slide structure | done | JSON output with text items, images, shapes, notes |
| Insert LaTeX equations | done | Via GUI scripting (System Events) |
| Batch slide creation | done | 20 slides per AppleScript call |
| Per-slide error handling | done | Error messages include slide number and layout |
| Set text alignment | no | Must be set in master slide |
| Rename/delete master slides | no | Must be preset in template |

## Known issues

1. **Two-column item order is counterintuitive**: Right column is `text item 2`, left is `text item 3` (or higher in body+col layouts). Baked into the master slide. The CLI hides this behind `left`/`right` content keys.

2. **`defaultBodyItem` broken for custom masters**: For all two-column masters, `defaultBodyItem` points to a hidden 0x0 placeholder. Must use `text item N` by index instead. Only works for `Title` and `Header-Body`.

3. **Duplicate text items on every slide**: Master-inherited duplicates appear after the content items. The CLI targets only the first N content items per layout.

4. **Placeholder labels not exposed via AppleScript**: The instructional text in Keynote's Edit Slide Layout view is not accessible. Text items are identified by index only.

5. **Formatting inherited from master**: Setting font/size/color works, but the template's masters already define correct formatting. The CLI only needs to set `object text` — formatting is inherited automatically.
