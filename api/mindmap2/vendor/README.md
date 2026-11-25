# vendor — mindmap2 が読む markmap 一式（版を固定）

| ファイル | 元 | 版 |
|---|---|---|
| d3-7.9.0.min.js | npm `d3` dist/d3.min.js | 7.9.0（ISC） |
| markmap-lib-0.18.12.js | npm `markmap-lib` dist/browser/index.iife.js | 0.18.12（MIT） |
| markmap-view-0.18.12.js | npm `markmap-view` dist/browser/index.js | 0.18.12（MIT） |

外部スクリプトを許さないアプリ（MinuteStock）でも描けるよう、リポジトリに置いています。
mindmap2.js はここを先に読み、無いときだけ CDN を使います。
版を上げるときは、公開から2週間以上たった版にし、ファイル名の版と mindmap2.js の `CDN` を合わせて変えます。
