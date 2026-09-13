# コンテンツナレッジ(contentsstock) セットアップ

動画を Google Drive に置くと、Mac 側で文字起こしされて Notion に並び、
画面からAIで サマリ / マインドマップ / 分野別 / 応用 / 活用 を作れるようにする。

```
ブラウザ(beta/contentsstock)
   │ アップロード
   ▼
GAS(contents) ──► Google Drive : inbox フォルダ
                        │ 15分ごと / 常駐
                        ▼
                  Mac: video_inbox.py（mlx-whisper で文字起こし）
                        │ Notion へページ作成 + 動画を contents フォルダへ移動
                        ▼
                     Notion コンテンツDB
                        ▲
                        │ 一覧・詳細・AI生成の保存
                  GAS(contents) ◄── ブラウザ
```

## 1. Notion のDB

対象DB: `d600e7a535dc83caadf381afe7abea03`

必要なカラム(名前は GAS の `PROP_*` と一致させること):

| カラム | 型 | 用途 |
| --- | --- | --- |
| タイトル | title | 動画名 |
| ファイル名 | rich_text | 取り込み元のファイル名 |
| Driveリンク | url | 元動画(再生用) |
| 種別 | select | mp4 / mov / pdf など |
| タグ | multi_select | 絞り込み |
| 状態 | select | 完了 / 要約済み / 除外 |
| 要約 | rich_text | カード用サマリ |
| マインドマップ | rich_text | markmap用Markdown |
| 分野別要約 / 応用 / 活用アイデア | rich_text | セクション形式 |
| メモ | rich_text | 自由記述 |
| 要約モデル | rich_text | 生成に使ったモデル名 |
| 要約日時 | date | 生成日時 |
| 原文文字数 | number | 文字起こしの長さ |
| 公開 | checkbox | 一覧公開の予約(任意) |

Notion Integration をこのDBに「接続」しておくこと。

## 2. GAS

`beta/contentsstock/gas/Code.gs` をプロジェクトへ貼り付ける。

- サービス → **Drive API** を追加する(動画アップロードで使う)
- スクリプトプロパティ
  - `NOTION_TOKEN` … Notion Integration のシークレット
  - `ACCESS_TOKEN` … 画面の「共有トークン」と同じ文字列(自分で決める)
  - `INBOX_FOLDER_ID` … `10dNn2zgtWCL4FpyYzam_EayKNtD7mGkz`
  - `CONTENTS_DB_ID` … 省略可(既定で上のDB)
  - `code` … 権限コードの対応表。例 `{"dfkjnga":"xYz"}`
- デプロイ: ウェブアプリ / 実行するユーザー: 自分 / アクセス: **全員**

デプロイ後、初回だけ実行を承認する(Drive と外部リクエストの権限)。

## 3. 画面

`beta/contentsstock/index.html` を開き、右上の歯車から設定する。

- GAS ウェブアプリURL(`/exec`)
- 共有トークン(GASの `ACCESS_TOKEN` と同じ値)
- 権限コード(使う場合)
- AI接続: baseURL + APIキー + モデル名
  - 議事録アプリ・動画ナレッジと同じ設定を使う(localStorage を共有している)

## 4. Mac 側(文字起こし)

Google Drive for desktop を**ミラーリング**で同期し、次の2つのフォルダがローカルに見えるようにする。

- inbox: `10dNn2zgtWCL4FpyYzam_EayKNtD7mGkz`
- contents: `16SN7XBWosS7WfbpEPUby4gWPDyAAY_px`

`~/.contentsstock.env` に設定を書く。

```
NOTION_TOKEN=secret_xxx
CONTENTS_DB_ID=d600e7a535dc83caadf381afe7abea03
CONTENTS_INBOX=/Users/you/Google Drive/マイドライブ/contents-inbox
CONTENTS_STORE=/Users/you/Google Drive/マイドライブ/contents
WHISPER_MODEL=mlx-community/whisper-large-v2-mlx
WHISPER_LANGUAGE=ja
```

導入:

```bash
pip install mlx-whisper requests
brew install ffmpeg
python3 mac/scripts/contentsstock/video_inbox.py --once --dry-run   # 対象の確認
python3 mac/scripts/contentsstock/video_inbox.py --once             # 1回だけ処理
python3 mac/scripts/contentsstock/video_inbox.py --interval 300     # 常駐
```

launchd に登録する場合は `mac/scripts/plaud/SETUP.md` の plist を参考にする
(cron ではなく launchd を使うのは、Drive のミラーを読むのにユーザーコンテキストが要るため)。

## 5. 一覧JSON(表示を速くする)

一覧を開くたびに GAS 経由で Notion を全件クエリすると、件数が増えるほど最初の描画が遅くなる。
`mac/scripts/contents/build_contentsstock_json.py` を cron で回して `data/contentsstock/` に
JSON を置いておくと、画面はそれを読むだけになる。JSON が無い・壊れているときは
これまでどおり Notion から直接読むので、動かなくても表示は止まらない。

出力:

- `index-doc.json` … コンテンツ一覧(カード表示・検索・絞り込み)
- `idea-doc.json` … 応用・活用アイデア(アイデア一覧画面)

```cron
20 * * * * /usr/bin/python3 ~/tools/mac/scripts/contents/build_contentsstock_json.py >> ~/contentsstock_json.log 2>&1
```

書き出しは一時ファイルを作ってから差し替えるので、途中で読んでも壊れた JSON を掴まない。
置き場所を変えたいときは設定の「一覧JSONの置き場所」に URL を入れる。

詳細(タブごとの本文)は鮮度が要るため、従来どおり Notion から取る。

## 6. 使い方

1. 画面右上「アップロード」で動画を選ぶ → Drive の inbox に入る
2. Mac 側が文字起こしして Notion にページを作る(動画は contents フォルダへ移動)
3. 画面を再読み込みすると一覧に並ぶ
4. 詳細を開いて「すべて生成」、または各タブの「この項目を作り直す」
5. マインドマップはキーボードで編集できる(↑↓移動 / スペース編集 / Tab子追加 / Enter同階層 / Delete削除)

### 一覧の切り替え

上部のタブで3つの見え方を行き来する。

- **コンテンツ** … 取り込んだもの全部
- **アイデア** … 応用・活用アイデアを1件ずつ並べる。要らないものは目のアイコンで外す
  (戻すときは詳細の「応用」「活用」タブの目のアイコン)
- **マインドマップ** … 詳細のマインドマップタブで「公開」したものだけ並ぶ

タグの整理は上部右端のタグアイコンから。いつも一緒に付いているタグの組を統合候補として出す。
「まとめる」を押すと Notion 側の全ページを書き換えるので、元に戻すには手作業が要る。

## 補足

- 文書(PDF/Word など)の取り込みは従来どおり `mac/scripts/contentsstock/contents_watch.py` が担当する。
  同じDB・同じ画面で扱えるよう、種別だけが違う形にしてある。
- 文字起こしは `[mm:ss]` 付きの行で保存される。原文タブでそのまま読める。
