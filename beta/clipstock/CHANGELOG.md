# Change Log

## 1.3.1 - 2026-09-03

### Title

Fix markmap failing to initialise: wrong CDN filename for markmap-lib

### Changes

- Stopped hardcoding browser bundle filenames. `markmap-view` ships `dist/browser/index.js` but `markmap-lib` ships `dist/browser/index.iife.js`; the previous code used `index.js` for both, so `Transformer` never appeared on `window.markmap` and initialisation failed. Requesting the bare package path lets the CDN resolve the entry from the package's own `jsdelivr` field.
- Verified each dependency immediately after loading it, so the error names the specific package and URL that failed instead of reporting a generic initialisation failure.

### Affected Files

- `videos/src/lib/mindmap.js`

### Notes

- The raw-markdown fallback behaved as intended during the failure — the content stayed readable.
- Versions remain pinned to `@0.18`; only the filename was at fault.

## 1.3.0 - 2026-09-03

### Title

Jump to the moment in the video: timecodes resolved from verbatim quotes

### Changes

- Added `src/lib/timecode.js`: parsing and formatting of `[mm:ss]` / `[h:mm:ss]`, transcript segmentation, `youtubeUrlAt`, and `resolveQuote`.
- Changed the fields stage to ask the model for the verbatim quote behind each point rather than for a time. Times are then resolved by matching that quote against the timestamped transcript, so a fabricated time cannot reach the UI — an unmatched quote simply yields no link.
- Quote matching normalises width, case and punctuation, and retries with progressively shorter prefixes, since models tend to paraphrase the tail of a quote.
- Rendered resolved times as playback links on section headings and points, in the transcript tab, in mindmap branches (via markdown links markmap makes clickable), and in chat answers.
- Asked the per-video chat to cite `[mm:ss]` from the transcript, but only when the transcript actually carries timestamps.
- Skipped the quote request entirely for transcripts without timestamps, so older material costs nothing extra and still renders.
- Added `docs/process_videos_patch.md` with the Python change: group segments into roughly 30-second lines prefixed with `[mm:ss]`.

### Affected Files

- `videos/src/lib/timecode.js` (new)
- `videos/docs/process_videos_patch.md` (new)
- `videos/src/lib/generate.js`
- `videos/src/lib/mindmap.js`
- `videos/src/lib/chat.js`
- `videos/src/ui/render.js`
- `videos/src/main.js`
- `videos/css/styles.css`
- `videos/preview.html`

### Notes

- Requires the Python change to take effect; existing videos need 状態 set back to 新規 to be re-transcribed.
- Times are only attached to 分野別要約. 応用 and 活用アイデア are generated from the summary rather than the transcript, so there is no quote to anchor them to.
- The absence of a link is meaningful: it means the claim could not be located in the transcript.

## 1.2.0 - 2026-09-03

### Title

Add a vocabulary panel: tag frequency, long tail, and merge candidates

### Changes

- Added `src/lib/vocab.js` with `tagStats`, `vocabSummary` and `mergeCandidates`, computed entirely from the already-loaded list — no extra API call and no LLM, so the same data always produces the same suggestions.
- Added a vocabulary panel (topbar, next to settings) showing tag usage as bars, with tags used twice or fewer separated below a dashed line as the cleanup queue.
- Surfaced merge candidates from co-occurrence: a less-used tag that almost always appears alongside a more-used one is likely a rephrasing rather than a separate angle. Each candidate carries a "merge" and a "keep separate" action.
- Kept only the strongest target per source tag, since a tag used once co-occurs 100% with every other tag on that video and would otherwise flood the list.
- Labelled single-use candidates as weak evidence rather than hiding them, since spelling and translation variants surface exactly there.
- Added a `mergeTag` GAS action that rewrites the tag across every matching page server-side, re-reading current tags so a stale client list cannot clobber them. It re-queries from the start each round rather than paging with a cursor, because rewritten pages drop out of the filter and would cause a cursor to skip rows.
- Stored "keep separate" decisions in localStorage rather than adding a Notion column, with a control to clear them.

### Affected Files

- `videos/src/lib/vocab.js` (new)
- `videos/src/ui/vocab.js` (new)
- `videos/src/lib/gas.js`
- `videos/src/main.js`
- `videos/index.html`
- `videos/css/styles.css`
- `videos/gas/Code.gs`

### Notes

- Merging is not reversible from the app; the confirmation says so.
- New-tag rate over time is deliberately not included — it needs a running log that does not exist yet.

## 1.1.0 - 2026-09-03

### Title

Constrain tag generation to the existing vocabulary

### Changes

- Passed the existing tag vocabulary (top 60 by usage) into the first generation stage, instructing the model to pick from it and to coin at most one new tag per video.
- Added `src/lib/tags.js` with `reconcileTags`, which maps returned tags back onto existing spellings using an NFKC + case + separator-insensitive key, so full-width/half-width and casing variants collapse deterministically rather than relying on the prompt.
- Dropped tags that are sentences (containing punctuation) or longer than 20 characters.
- Kept the vocabulary growing within a bulk run, so later videos in the same run see tags coined by earlier ones.
- Left the vocabulary unconstrained when it is empty, so the first videos can establish a base set.
- Added a dismissible notice when a new tag is coined, reusing the bulk progress bar slot.

### Affected Files

- `videos/src/lib/tags.js` (new)
- `videos/src/lib/generate.js`
- `videos/src/main.js`

### Notes

- Reconciliation only merges spelling variants of the same word. Words that are merely close in meaning (生成AI and LLM, Notion and ノーション) are left alone — merging those is a judgement call for a person, not a string comparison.

## 1.0.0 - 2026-09-02

### Title

Initial release of `videos`: video knowledge library built on the `minutes` architecture

### Changes

- Added a new app `videos` targeting the Notion 動画 database, reusing the browser → GAS → Notion structure of `minutes`.
- Replaced the static `index.json` dependency with a `listVideos` GAS action that queries the Notion database directly, so the list no longer needs a batch-generated intermediate file.
- Added a `listIdeas` GAS action returning only 応用/活用アイデア, keeping the main list payload small.
- Lifted the 2000-character property cap by chunking `rich_text` into multiple objects (up to 100 × 2000) in `richTextProp_`.
- Stored 分野別要約 / 応用 / 活用アイデア in a human-readable `## heading` + `- bullet` format instead of JSON, so the values remain useful when read directly in Notion.
- Stored マインドマップ as markmap Markdown rather than embedded HTML, with detection and iframe fallback for pages written by the previous skill.
- Split AI generation into three stages (core / fields / apply), persisting after each stage so a failure does not discard earlier results, and enabling per-section regeneration.
- Reused the shared `gemma-chat.settings` localStorage key for AI connections, so connections added in `minutes` or `gemma-chat` are available here.
- Added a thumbnail grid library view with per-video generation progress dots and an unread marker.
- Added an idea feed view that flattens 応用/活用アイデア across the whole library, with a shuffle mode for resurfacing older material.
- Added per-video chat (raw transcript or summary as context) and cross-video chat spaces limited to 20 targets.
- Added bulk generation over items whose 状態 is 完了, with progress and cancel.
- Added manual editing, tag editing, title editing, re-transcribe (状態 → 新規) and logical exclusion (状態 → 除外).

### Affected Files

- `videos/index.html`
- `videos/css/styles.css`
- `videos/gas/Code.gs`
- `videos/src/main.js`
- `videos/src/ui/render.js`
- `videos/src/ui/settings.js`
- `videos/src/lib/gas.js`
- `videos/src/lib/generate.js`
- `videos/src/lib/sections.js`
- `videos/src/lib/mindmap.js`
- `videos/src/lib/chat.js`
- `videos/src/lib/cache.js`
- `videos/src/lib/filters.js`
- `videos/src/lib/videos-config.js`
- `videos/src/lib/llm-client.js` (copied from `minutes`)
- `videos/src/lib/llm-settings.js` (copied from `minutes`)
- `videos/src/lib/markdown.js` (copied from `minutes`)
- `videos/docs/SETUP.md`

### Notes

- Requires new Notion columns: 分野別要約 / 応用 / 活用アイデア / メモ / 要約モデル / 要約日時 / 原文文字数. See `docs/SETUP.md`.
- The Python cron script (`process_videos.py`) is unchanged; it still owns everything up to 状態 = 完了.
- `llm-client.js`, `llm-settings.js` and `markdown.js` are byte-identical copies of the `minutes` versions. Keep them in sync when either side changes.
