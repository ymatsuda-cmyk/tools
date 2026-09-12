# Change Log

## 1.3.0 - 2026-09-12

### Release Title

Mindmap tab for the detail pane

### Release Changes

- Added a "マインドマップ" tab next to サマリ in the detail pane, using the same mechanism as clipstock: markmap renders Markdown loaded from CDN.
- Added Notion「マインドマップ」column (rich_text, markmap Markdown) with `saveMindmap` action; the tab shows nothing when the column is empty.
- Generate from the transcript with the active LLM connection as structured JSON (`{title, branches}`) and build the Markdown here, so a model cannot break the markmap parse. Falls back to the summary structure when generation fails.
- Added a 手で直す mode that edits the Markdown directly, with the unsaved-changes warning when leaving the tab.
- Put a cursor on the map: it starts on the centre node, moves with the arrow keys, folds with ←/→, edits in place with Space, and adds a child with Tab or a sibling with Enter. Edits are written back into the Markdown as an unsaved draft.
- Fed the edited Markdown back through `setData` instead of rebuilding the map, so only the new branch appears and the zoom, pan and folds stay as they were.
- Moved the mindmap engine into `api/mindmap2` (shared with the video knowledge app) and left only the summary-to-Markdown generation here.

### Release Affected Files

- `api/mindmap2/mindmap2.js`
- `api/mindmap2/mindmap2.css`
- `api/mindmap2/index.html`
- `minutes/index.html`
- `minutes/src/main.js`
- `minutes/src/ui/render.js`
- `minutes/src/lib/mindmap-view.js`
- `minutes/src/lib/gas.js`
- `minutes/gas/Code.gs`
- `minutes/css/styles.css`

## 1.2.0 - 2026-09-04

### Release Title

Beta parity update: audio upload pipeline and GAS upload proxy support

### Release Changes

- Added upload entry points to the minutes UI: Google script loader, upload button, and upload module bootstrap.
- Added new upload client module with chunked transfer, pending-item rendering, and sidecar metadata registration.
- Extended GAS action router with `initUpload`, `putChunk`, and `writeSidecar` endpoints.
- Added resumable Drive upload session support and chunk relay implementation in GAS.
- Added sidecar JSON writer for Mac mini inbox workflow handoff.
- Improved agenda JSON parse diagnostics with explicit error logging.
- Updated rich text writer to chunk long strings instead of truncating at 2000 chars.
- Added upload-related styles for modal fields, drop zone, progress, status messages, and pending cards.

### Release Affected Files

- `minutes/index.html`
- `minutes/src/upload.js`
- `minutes/gas/Code.gs`
- `minutes/css/styles.css`

## 1.1.0 - 2026-09-02

### Title

Beta parity release for minutes: LLM connection model, list controls, and raw-context handling

### Changes

- Migrated LLM settings from single-model profiles to connection-based settings with multi-model support.
- Added migration logic so existing local settings continue to work after the schema change.
- Added quick model switching from the active-model label in the top bar.
- Updated settings UI to edit connections and per-connection model lists.
- Added tag visibility toggle integration and synchronized toolbar/topbar state.
- Improved assign mode behavior: month navigation can narrow period and reset can return to full-period view.
- Restored assign-mode filter consistency by clearing selected tags when period scope changes.
- Added raw transcript context count persistence via `rawContextCount` on summary save.
- Added lightweight `updateRawContextCount` updates to cache transcript length without full regeneration.
- Enhanced raw chat context flow with agenda/raw context mode handling and count display.

### Affected Files

- `minutes/index.html`
- `minutes/css/styles.css`
- `minutes/src/main.js`
- `minutes/src/lib/gas.js`
- `minutes/src/lib/llm-settings.js`
- `minutes/src/lib/summarize.js`
- `minutes/src/ui/render.js`

### Notes

- This release reflects promoted changes from `LLM/beta/minutes` into `minutes`.
