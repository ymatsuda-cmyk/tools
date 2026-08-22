# gemma chat（ビルド不要版）

Kaggle 上の Ollama（gemma4:12b）を OpenAI 互換エンドポイント経由で使う、
長文・ファイル対応のチャットアプリ。素の JS / HTML / CSS のみ。
ビルドもパッケージマネージャも不要。

## ローカルで動かす

ES モジュールを使うので `file://` では動かない。ローカルサーバーを立てる。

```bash
cd gemma-chat-js
python3 -m http.server 8000
```

`http://localhost:8000/` を開く。Vite 版と違ってサブパスは不要。

## GitHub Pages へ

```bash
git init && git add -A && git commit -m "initial"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

Settings → Pages → Source を「Deploy from a branch」にして main / (root) を選ぶだけ。
ビルドがないので GitHub Actions は不要。パスはすべて相対なので
リポジトリ名を気にする必要もない。

## ライブラリ

`index.html` の importmap で CDN から読む。バージョンを上げたいときはここを書き換える。

| ライブラリ | 用途 | 読み込み |
|---|---|---|
| dexie | IndexedDB | importmap |
| marked | Markdown | importmap |
| dompurify | HTML サニタイズ | importmap |
| highlight.js | シンタックスハイライト | 遅延 import |
| pdfjs-dist | PDF 抽出 | 遅延 import |
| mammoth | docx 抽出 | 遅延 script タグ |
| xlsx | Excel 抽出 | 遅延 script タグ |

パーサ3種は実際にそのファイルを添付したときだけ落ちてくるので、初回表示は軽い。

CDN 依存をなくしたい場合は各ファイルを `vendor/` に置き、importmap の URL を
`./vendor/dexie.js` のような相対パスに書き換える。コードの変更は不要。

## ファイル構成

```
index.html          importmap と骨組み
css/styles.css
src/main.js         状態管理と送信処理
src/lib/dom.js      h() ヘルパーと script ローダ
src/lib/tokens.js   トークン推定（実測値ベース）
src/lib/settings.js localStorage とキー検証
src/lib/db.js       Dexie スキーマ
src/lib/client.js   OpenAI 互換 SSE クライアント
src/lib/parsers.js  PDF / docx / xlsx / テキスト
src/lib/markdown.js サニタイズとスロットル描画
src/ui/*.js         各画面部品
```

## 実装済み

- 設定（ベースURL・APIキー・モデル・num_ctx）と接続テスト
- SSE ストリーミング表示、途中停止
- IndexedDB による会話履歴の永続化
- ファイル添付とパース、2,000 字超のペーストを自動カード化
- コンテキスト使用量メーターと待ち時間予測
- 上限超過時の送信ブロックと、原因になっている添付の名指し
- 応答後に `usage.prompt_tokens` を推定と突き合わせて切り捨てを検出

## 実測値（P100 16GB / gemma4:12b）

- prefill 速度: 約 270 tok/s
- num_ctx 32768 が確保できる
- 入力の実質上限: 約 30,000 トークン
- 日本語の散文で約 2.3 字/トークン、英数字混在で約 1.2 字/トークン

`src/lib/tokens.js` の定数はこの実測に基づく。GPU が変われば測り直すこと。

## 必要なプロキシ側の変更

ストリーミング時の切り捨て検出には、Kaggle 側プロキシが最終チャンクに
`usage` を載せる必要がある。`to_chunk` を差し替える。

```python
def to_chunk(o, model, first):
    delta = {'role': 'assistant'} if first else {}
    delta['content'] = o.get('message', {}).get('content', '')
    chunk = {'id': 'chatcmpl-1', 'object': 'chat.completion.chunk',
             'created': int(time.time()), 'model': model,
             'choices': [{'index': 0, 'delta': delta,
                          'finish_reason': 'stop' if o.get('done') else None}]}
    if o.get('done'):
        pt = o.get('prompt_eval_count', 0)
        ct = o.get('eval_count', 0)
        chunk['usage'] = {'prompt_tokens': pt, 'completion_tokens': ct,
                          'total_tokens': pt + ct}
    return chunk
```

## 未実装（Phase 2 以降）

- GAS 経由の Kaggle 起動・停止と週次クォータ表示
- 画像入力、埋め込みによる RAG、PWA 化
