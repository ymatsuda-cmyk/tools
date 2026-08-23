# Kaggle × Ollama マルチLLM基盤 — システム仕様書

## 1. 全体構成

3つの独立した層で構成される。**制御系（GAS）とデータ系（ngrok経由のプロキシ）は完全に分離**しており、チャットのやり取りはGASを一切通らない。

```
[チャットアプリ (GitHub Pages / 静的SPA)]
  │
  ├─ 制御 ──→ [GAS Web App]（1エンドポイント=1GASプロジェクト）
  │              ├─ Kaggle API（起動/停止/状態取得）
  │              └─ PropertiesService（設定・稼働ログの保存）
  │
  └─ データ ──→ [ngrok] ──→ [Kaggle Notebook]
                              ├─ ollama serve（モデル本体）
                              └─ 自作プロキシ（OpenAI互換変換、:11435）
```

### 設計上の理由
- GASのHTTPリクエストは60秒制限・ストリーミング不可のため、チャットには使えない
- Kaggle APIキーはGASの外に一切出さない（ブラウザに渡すのはproxyKeyとcontrolTokenのみ）
- **1エンドポイント = 1つのKaggleアカウント = 1つのGASプロジェクト = 1つのngrokアカウント**
  他アプリから個別接続する要件があるため、複数エンドポイントを1つのGASに集約する方式は採らない

---

## 2. GAS（Google Apps Script）

### 2.1 ファイル構成
- `コード.gs` — ルーティングとロジック本体
- `proxy_py.html` — Kaggle Notebookに push するPythonソース（HTMLファイルとして保存。`<`は`&lt;`にエスケープされるため、ソース中で`<`比較演算子を使わない）

### 2.2 設定はCONFIGプロパティに一本化

スクリプトプロパティ `CONFIG` にJSON文字列で保存する。

```json
{
  "kaggleUsername": "アカウント名",
  "kaggleApiKey": "Kaggle APIキー",
  "kaggleKernel": "username/kernel-slug",
  "kernelTitle": "push時のタイトル",
  "model": "Ollamaのモデル名（例: gemma4:12b）",
  "ngrokDomain": "予約済み固定ドメイン（https://なし）",
  "proxyApiKey": "このエンドポイント専用キー",
  "controlToken": "このエンドポイント専用キー（別値）",
  "datasets": ["Kaggleデータセットref", "..."],
  "gasWebappUrl": "このプロジェクトの/exec URL"
}
```

旧形式（個別プロパティ `KAGGLE_API_KEY` 等）へのフォールバックも実装済みだが、新規は必ずCONFIGを使う。

### 2.3 GASが自動生成する内部プロパティ（手で触らない）

| キー | 内容 |
|---|---|
| `USAGE_LOG` | 稼働ログ。JSON配列 `[{"s":開始ISO,"e":終了ISO,"m":分},...]`。上限120件で古い順に切り捨て |
| `SESSION_ID` | notebookが起動時に自己申告する内部セッションID（強制停止用） |
| `LAST_HEARTBEAT` | 30秒ごとのポーリングで更新される生存確認 |
| `MODEL_READY` | Ollamaが実際にチャット応答可能になったかのフラグ |
| `STOP_REQUESTED` | 停止フラグ |

### 2.4 doGetのアクション一覧

| action | 認証要否 | 説明 |
|---|---|---|
| `ping` | 不要 | 疎通確認のみ |
| `status` | 要 | 現在の状態一式を返す |
| `start` | 要 | Kaggle Notebookをpushして起動 |
| `stop` | 要 | 停止（下記3段階） |
| `started` | 要 | notebookからの起動通知（sessionId, modelReady受け取り） |
| `shouldStop` | 要 | notebookからの30秒ポーリング（ハートビート兼停止確認） |
| `record` | 要 | notebook終了時の稼働時間確定 |
| `forceStop` | 要 | 強制停止のみ単体実行 |

認証は `token` パラメータと `CONFIG.controlToken` の一致で行う。JSONP形式（`callback`パラメータ）にも対応。

### 2.5 停止の3段階エスカレーション

```
1. 通常停止: STOP_REQUESTEDフラグを立てる
   → notebookが30秒ポーリングで検知し自ら終了（recordが正常に走る）

2. 強制停止（ハートビート途絶時に自動選択）: 
   Kaggle API "cancel-session/{sessionId}" を呼ぶ
   → record は走らないため GAS 側が closeOpenRun() で肩代わり

3. 最終手段（cancel-session失敗時）:
   GPU無効の空ノートブックを同じスラグにpush
   → 実行中セッションが強制的に置き換えられ終了
```

`handleStop()`がハートビートの生死を見て自動的に1→2に切り替える。3は2失敗時のフォールバック。

### 2.6 週次クォータ集計

- Kaggle公式APIにクォータ取得エンドポイントは存在しない（調査済み）
- `USAGE_LOG` から自前集計。週の起点は**土曜0時UTC**
- 未クローズ行（`e:null`）は「稼働中」とみなし経過時間を暫定加算
- **注意**: 停止処理が正常に完了しないと未クローズ行が溜まり、クォータ表示が実態と乖離する。`closeStaleRuns()`で手動修復可能

---

## 3. Kaggle Notebook（proxy_py.html）

### 3.1 起動シーケンス

```
1. zstd/ollamaインストール
2. データセットから blobs/manifests を探索してコピー
   （find_models_dir()で階層構造に依存せず探索）
3. ollama serve 起動
4. ollama show でVISION対応を確認・ログ出力
5. wait_model_ready(): 実際にダミーチャットを送り応答確認（最大5分、5秒間隔）
6. プロキシ(:11435)起動
7. ngrok起動
8. セッションID抽出（KAGGLE_CONTAINER_NAMEから数値部分を正規表現抽出）
9. gas('started', sessionId, modelReady) 送信
10. 30秒間隔でgas('shouldStop')ポーリング
11. 停止指示 or MAX_MINUTES(540分)到達で終了
12. gas('record') で稼働時間確定
```

### 3.2 プレースホルダ（GAS側が push 直前に置換）

`__PROXY_KEY__` `__GAS_URL__` `__CONTROL_TOKEN__` `__NGROK_DOMAIN__` `__ENDPOINT_ID__` `__MODEL__`

### 3.3 gas()関数の実装上の注意点（すべて実際に踏んだ罠）

```python
def gas(action, **params):
    if GAS_URL.startswith('__'):
        return {}
    params.update({'action': action, 'token': CONTROL_TOKEN, 'id': ENDPOINT_ID})
    url = GAS_URL + '?' + urllib.parse.urlencode(params)
    last_err = None
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as r:
                body = r.read().decode('utf-8', 'replace')
            return json.loads(body)
        except Exception as e:
            last_err = e
            is_last = (attempt == max_attempts - 1)
            if not is_last:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f'GAS 呼び出しに失敗しました（{max_attempts}回試行）: {last_err}')
```

- **User-Agent必須**: urllibの既定UAはGoogleに弾かれログインページが返ることがある
- **リトライ必須**: GAS Web Appは稀に数秒〜数十秒、一時的にデプロイエラーを返す（Google側の既知の性質）
- **JSON応答を検証**: HTMLが返った場合に気づけるよう例外を投げる
- **`<`を使わない**: GASのHTMLファイルエディタが`&lt;`に自動エスケープし構文エラーになる。ループはwhile+比較ではなくfor+range方式にする
- `gas('started', ...)`は必ずtryで囲む。囲まないとここで例外が起きた場合ポーリングループに到達せず、チャットは動くのにGAS連携だけ死ぬ

### 3.4 OpenAI互換プロキシ（:11435）の変換ロジック

- `/v1/models`, `/v1/chat/completions` を実装
- Bearer認証（`PROXY_API_KEY`一致）
- CORS: `do_OPTIONS`で`Access-Control-Max-Age: 86400`を返しpreflightをキャッシュ
- 画像: OpenAI形式`image_url`（base64データURL）→ Ollama形式`images`配列に変換
- ストリーミング: OllamaのNDJSON → SSE形式に変換、最終チャンクに`usage`を付与（切り捨て検出用）

---

## 4. データセット要件

- `ollama pull <model>` の結果一式（`blobs/` `manifests/`）を含むKaggle Dataset
- 画像対応が必要な場合、モデル自体がマルチモーダル対応である必要がある（`ollama show`の`Capabilities`に`vision`が出るか事前確認必須）
- データセット作成はCPU(Accelerator: None)で行い、GPU週間枠を消費しない
- 階層は問わない（`find_models_dir()`が`blobs`と`manifests`を含むディレクトリを再帰探索する）

---

## 5. チャットアプリ（gemma-chat-js）

### 5.1 構成
ビルド不要の素のJS/HTML/CSS。importmapでCDNからESM読込。

```
index.html
src/main.js              状態管理・送信処理
src/lib/
  settings.js             プロファイル配列の管理
  gas.js                  GAS呼び出し（JSONP）・状態判定
  client.js               OpenAI互換SSEクライアント
  parsers.js              PDF/docx/xlsx/画像パーサ
  tokens.js               トークン推定・自動較正
  markdown.js             サニタイズ・スロットル描画
src/ui/
  endpoint-panel.js        サイドバーのエンドポイント切替
  settings-dialog.js       設定（プロファイルのタブ切替＋JSON一括編集）
  composer.js / message-list.js / dropzone.js / modal.js
```

### 5.2 設定スキーマ（localStorage: `gemma-chat.settings`）

```json
{
  "profiles": [
    {
      "id": "任意の短い識別子",
      "label": "表示名",
      "baseUrl": "https://xxx.ngrok-free.dev/v1",
      "apiKey": "proxyApiKeyの値",
      "model": "モデル名",
      "numCtx": 32768,
      "gasUrl": "このプロファイル専用のGAS /exec URL",
      "controlToken": "このプロファイル専用のCONTROL_TOKEN"
    }
  ],
  "activeId": "現在選択中のid",
  "temperature": 0.7,
  "systemPrompt": "共通システムプロンプト"
}
```

**重要**: `baseUrl`/`apiKey`/`gasUrl`/`controlToken`は**プロファイルごと**に独立している。共通設定ではない。

### 5.3 状態判定ロジック（gas.js: deriveState）

```
stopRequested === true        → 停止処理中
status が running/queued
  かつ proxyAlive かつ modelReady → 稼働中（ready）
  かつ proxyAlive のみ            → モデル読込中（loading）
  かつ zombie（ハートビート途絶）  → 応答なし（強制停止対象）
  それ以外                        → 起動中（booting）
それ以外                          → 停止中（stopped）
```

### 5.4 接続情報の取得は必ず`connectionOf(settings)`経由

```javascript
const conn = connectionOf(settings)  // activeProfileからbaseUrl等を抽出
// streamChat(conn, ...) のように渡す。settings全体を渡すと baseUrl が undefined になる
```

過去に`streamChat(settings, ...)`と誤って書いた箇所が長時間の障害原因になった。新規実装時は必ず`conn`を経由すること。

---

## 6. 実測値（P100 16GB / gemma4:12b、モデル依存で変わる）

| 項目 | 値 |
|---|---|
| prefill速度 | 約270 tok/s |
| 生成速度 | 約20 tok/s |
| num_ctx上限 | 32768（VRAM制約。モデル自体は262144まで対応） |
| トークン効率 | 日本語 約0.5 tok/字 |
| 週間クォータ | 30時間（変動あり、リセットは土曜0時UTC） |

これらはモデル・GPU種別が変わるたびに再計測が必要。

---

## 7. 既知の落とし穴一覧（再発防止用）

| 症状 | 原因 |
|---|---|
| APIキー401エラー | コピペ時の改行・全角文字混入 |
| CORS preflightが405 | 実はバックエンド未起動（ngrokのエラーページがCORSヘッダを持たない） |
| GASが`Unknown action`を返す | マルチ版/単体版のコード取り違え、または未デプロイ |
| GASが「ファイルを開けません」 | デプロイが無効。新バージョンとして再デプロイが必要 |
| 入力欄でフォーカスが飛ぶ | DOM全体を作り直す実装。構造構築と値更新を分離する |
| 画像が認識されない | データセットのモデルがテキスト専用ビルド。`ollama show`で`vision`確認 |
| チャットURLに`undefined`混入 | `streamChat`に`settings`全体を渡している（`conn`ではなく） |
| GAS連携が沈黙して死ぬ | `gas('started')`をtryで囲んでいない |
| 稼働時間が実態と乖離 | 強制停止・異常終了で未クローズ行が蓄積 |
| Notebookが原因不明で構文エラー | HTMLエディタが`<`を`&lt;`にエスケープ |
