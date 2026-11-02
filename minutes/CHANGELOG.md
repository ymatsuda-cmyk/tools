# Change Log

## 1.8.0 - 2026-11-02

### Release Title

Text registration takes raw text (required) and a summary text (optional); register a summary text from the agenda tab

### Release Changes

- 「テキストから議事録を登録」を、原文(必須)と要約テキスト(任意)の2欄にした(1.7.0 の切替ボタンは廃止)。要約テキストを入れると、タイプ1・2を自動判別して議事・決定事項・ToDo・論点に分け、状態「要約」で登録する。空なら従来どおり状態「文字起こし」。原文はどちらの場合もページ本文に入る。
- 議事タブの見出しに「要約テキストから登録」ボタンを追加。貼り付けた要約で議事を置き換え、決定事項・ToDo・論点は要約に含まれるときだけ置き換える(ToDoは文言の近さで突き合わせ、チェック状態と経過ログを引き継ぐ)。サマリは空のときだけ補う。
- GAS の変更なし(1.7.0 の `createFromText` の `summary` をそのまま使用)。

### Release Affected Files

- `minutes/src/text-register.js`
- `minutes/src/main.js`
- `minutes/src/ui/render.js`
- `minutes/css/styles.css`
- `minutes/CHANGELOG.md`

## 1.7.0 - 2026-11-02

### Release Title

Register minutes from an existing summary text (type 1 / type 2)

### Release Changes

- 「テキストから議事録を登録」に「原文(文字起こし)／要約テキスト」の切替を追加。要約テキストでは、すでにある要約を貼り付けてそのまま登録できる(要約の再生成は不要)。
- タイプ1(`## 会議情報 / ## 議事メモ / ## 次の段取り / ## AIからの提案`)とタイプ2(`会議のメモ:` ～ `フォローアップ タスク:`)を自動判別する(`src/lib/summary-text.js`)。貼り付けた時点で読み取り件数を表示し、読み取れない場合は理由を出す。
- 変換: 議題 → 議事(トピック/要点/結論)、タイプ1の「結論」→ 決定事項、「次の段取り」/「フォローアップ」→ ToDo、「AIからの提案」→ 論点。一覧カードの要約は議題名の並びから作る。タイプ1は日付を開始日時に自動入力する(タイプ2は日付が無いので手入力)。
- 状態は「要約」、要約モデルは「テキスト取り込み」で登録し、貼り付けた全文はページ本文にも残す。
- `createFromText` in `gas/Code.gs` accepts `summary`. **Redeploy the GAS web app.**

### Release Affected Files

- `minutes/src/lib/summary-text.js` (new)
- `minutes/src/text-register.js`
- `minutes/gas/Code.gs`
- `minutes/css/styles.css`
- `minutes/CHANGELOG.md`

## 1.6.0 - 2026-10-15

### Release Title

Remove status filter, permissions on text registration, header button tooltips

### Release Changes

- Removed the 文字起こし / 再取得 / 要約 status filter chips from the top bar (and the unused `filterByStatus`).
- 「テキストから議事録を登録」now has a 権限 field. Admins (xYz) pick any permissions with suggestions from existing items (empty = admin only); other users see their own role as a locked chip.
- `createFromText` in `gas/Code.gs` accepts `permissions` only when the role is admin; for other roles it still forces the verified role. **Redeploy the GAS web app** for admin-selected permissions to take effect.
- Top bar icon buttons now highlight on hover and show their label in a tooltip (`src/header-tooltip.js`). The tooltip is rendered as a fixed element on body so the overflow-hidden top bar doesn't clip it; not shown for touch input.

### Release Affected Files

- `minutes/index.html`
- `minutes/src/main.js`
- `minutes/src/lib/filters.js`
- `minutes/src/text-register.js`
- `minutes/src/header-tooltip.js` (new)
- `minutes/gas/Code.gs`
- `minutes/css/styles.css`

## 1.5.0 - 2026-10-15

### Release Title

Register minutes from text (title, tags, start time, duration, transcript)

### Release Changes

- Added a「テキストから登録」button to the top bar that opens a modal for title, tags, start date/time, duration (hours/minutes/seconds) and the raw transcript.
- Writes to existing Notion columns: ミーティング名, カテゴリー, 日時, 会議時間 (as「1時間 12分 36秒」, hours omitted when zero) and the page body; also sets 原文文字数 and 状態 = 文字起こし. No summary is generated.
- Added `createFromText` action to `gas/Code.gs`. The transcript is packed line by line into paragraph blocks of up to 2000 characters and sent 40 blocks per request to stay under Notion's payload limit.
- Tag suggestions come from tags used in index.json items the viewer is allowed to see; the verified role is attached as 権限 for non-admin users, same as the audio upload.
- Reused the upload pending card with a new `registered` status (「一覧への反映待ち」) until index.json picks up the page.

### Release Affected Files

- `minutes/index.html`
- `minutes/src/text-register.js` (new)
- `minutes/src/upload.js`
- `minutes/gas/Code.gs`
- `minutes/css/styles.css`

## 1.4.0 - 2026-10-01

### Release Title

ToDo progress log and learning rules from manual edits (restored on top of 1.3.0)

### Release Changes

- Restored the 1.3.0 base after commit 3467906c overwrote `src/upload.js`, `src/lib/mindmap-view.js` and the upload/mindmap actions in `gas/Code.gs`; only the features below are layered on top.
- Added a per-ToDo progress log (日付付きで追記のみ). Logs are stored in the new Notion「ToDo経過」column as JSON; the「ToDo」column stays human-readable.
- Moved the log input to the fixed footer of the ToDo tab; clicking a card selects the target (defaults to the first open ToDo).
- Merged ToDos before/after an update by text similarity so logs and checks survive re-summarisation; unmatched logs are kept as orphans on regenerate and confirmed before deletion on manual edit.
- Added learning rules: saving a manual edit extracts a one-line rule in the background into the「要約ルール」DB, and active rules are appended to the summary prompt.
- Added a 学習一覧 section in settings to toggle or delete rules.

### Release Affected Files

- `minutes/gas/Code.gs`
- `minutes/src/main.js`
- `minutes/src/ui/render.js`
- `minutes/src/lib/todos.js`
- `minutes/src/lib/gas.js`
- `minutes/src/lib/summarize.js`
- `minutes/css/styles.css`

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
