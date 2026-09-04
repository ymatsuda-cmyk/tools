# Change Log

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
