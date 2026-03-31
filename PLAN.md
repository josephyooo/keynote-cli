# keynote-cli — Plan

## Architecture

**AppleScript via `osascript`**, called from Python. Scripts are compiled into batched AppleScript for performance.

- **JXA rejected**: `masterSlides` is completely broken (`"Can't convert types"` on every access).
- **Swift deferred**: No benefit at this stage.

## Capability status

| Feature | Status | Notes |
|---------|--------|-------|
| Open/save `.key` files | done | |
| Create slide from any named master | done | |
| Set text on any item | done | via `defaultTitleItem`, `defaultBodyItem`, `textItem:N` |
| Set paragraph indent levels | done | via `--indents` |
| Add free text items | done | Position, size, font, color |
| Set font/size/color on text | done | Must use font variant names (e.g. `"Helvetica-Bold"`) |
| Add images | done | With position/size |
| Override element properties | done | text, position, size, font, color, opacity, rotation |
| Set presenter notes | done | |
| Delete slides by range | done | |
| Export to PDF | done | |
| Inspect slide structure | done | JSON output |
| Insert LaTeX equations | done | Via GUI scripting |
| Batch execution (20 slides/call) | done | |
| Per-slide error handling | done | |
| Set text alignment | no | Must be set in master slide |
| Rename/delete master slides | no | Must be preset in template |

## Known issues

1. **Duplicate text items on every slide**: Master-inherited duplicates appear. When targeting by index, know that items after the content items may be duplicates or hidden 0x0 placeholders.

2. **`defaultBodyItem` broken on some masters**: For many custom masters, `defaultBodyItem` points to a hidden 0x0 placeholder. Use `textItem:N` instead.

3. **Placeholder labels not exposed**: The instructional text in Keynote's Edit Slide Layout view is not accessible via AppleScript. Items are identified by index only.

4. **Text formatting inherited from master**: The template's masters define formatting. Setting `object text` inherits formatting automatically.

## Future direction

Standalone mutation commands (e.g. `keynote-cli set-text --document front ...` or `keynote-cli set-text --file out.key ...`) could allow individual operations outside a script. This would require specifying the document source — either an already-open window (`--document front`) or a file to open (`--file path.key`). Currently, mutation commands only work inside script files run via `keynote-cli run`.
