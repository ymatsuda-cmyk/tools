# WBSカンバン タスク状況API

WBSカンバンアドイン（`addin/kanban`）が使っている Excel の `wbs` シートを、
ダッシュボードなど Excel の外から「遅延・本日〆・今週・来週」で見るための入口です。
取得先に置くデータは**暗号化**してあり、復号はブラウザの中だけで行います。
既定の取得先は `data/kanban/wbs-tasks.enc.json`（GitHub Pages経由）です。旧方式の
JSONBinにも引き続き対応しています。

```
api/kanban/
├── kanban.api.js            呼び出す側が使う入口（取得・復号・分類）
├── kanban.api.test.mjs      node --test api/kanban/kanban.api.test.mjs
└── office-script/
    └── export-wbs.ts        Power Automate から呼ぶ Office Script
```

## データの流れ

```
Excel（wbsシート）
  │ カンバンアドインの操作で R/S/O/H 列が変わり、ブックが保存される
  ▼
Power Automate  … OneDrive for Business「ファイルが変更されたとき」
  │ Excel Online (Business)「スクリプトの実行」export-wbs.ts
  │ OneDrive for Business「ファイルの更新」
  ▼
OneDrive  Apps/kanban-sync/wbs-tasks.json（平文。Microsoft 365 の中だけ）
  │ OneDrive アプリで Mac に同期
  ▼
Mac  mac/scripts/github/github_aes_sync.py（kanban_aes.json）
  │ AES-256-GCM で暗号化して data/kanban/wbs-tasks.enc.json に commit / push
  ▼
GitHub（リポジトリ内の暗号文だけ）→ GitHub Pages で配信
  │ GET data/kanban/wbs-tasks.enc.json
  ▼
ダッシュボード  kanban.api.js が復号・分類して描画
```

旧方式（JSONBinに直接PUTする `mac/scripts/jsonbin/jsonbin_sync.py`）も
`load({ binId, apiKey, keyType })` でそのまま読めます。

平文が存在するのは Excel・OneDrive・Mac の中だけです。JSONBin とブラウザの通信経路、
ダッシュボードの設定（`index.json`）には、タスクの中身もパスフレーズも入りません。

---

## 使い方

```javascript
const kanban = await import('/api/kanban/kanban.api.js')

// 既定（data/kanban/wbs-tasks.enc.json）から読むだけなら passphrase だけでよい
const data = await kanban.load({ passphrase: '…' })

// 旧方式のJSONBinから読みたいときは binId を指定する
const dataFromBin = await kanban.load({
  binId: '6512abcd…',
  apiKey: '',            // 公開Binなら空でよい
  keyType: 'access',     // 'access'（X-Access-Key）/ 'master'（X-Master-Key）
  passphrase: '…'        // 暗号文のときに必須
})

const rows = kanban.classify(data.tasks, { user: '' })   // 日付で変わるので描画のたびに呼ぶ
rows.late.rest     // 遅延（未完了で期限切れ）
rows.today.rest    // 本日〆の未完了 / rows.today.done 完了 / rows.today.total 総数
rows.week, rows.next
```

### 関数

| 関数 | 内容 |
|---|---|
| `load(config)` | 既定（`data/kanban/wbs-tasks.enc.json`）または `config.url`/`config.binId` から読み、復号して `{ schema, updatedAt, tasks }` を返す |
| `classify(tasks, { user, now })` | 4行ぶんに振り分ける。`user` で担当者を絞れる |
| `decrypt(envelope, passphrase)` / `encrypt(json, passphrase)` | 暗号化の形式どおりに復号・暗号化 |
| `overdueDays(task)` | 期限切れの日数 |
| `users(tasks)` / `formatMd(date)` / `toDate(v)` | 表示のための小物 |

`load()` の失敗は `err.code` で見分けます。

| code | 意味 |
|---|---|
| `fetch` | 接続できない・キーが違う |
| `empty` | Bin にまだ何も無い |
| `locked` | 暗号文なのにパスフレーズが渡されていない |
| `badpass` | パスフレーズが違う、または改ざん・破損 |
| `format` | 中身の形が違う |

### 分類のルール

判定に使う列と記号は、カンバン（`addin/kanban/kanban.js`）と同じです。

| 行 | 対象 | 総数 |
|---|---|---|
| 遅延 | 未完了で、予定完了日（Q列）が今日より前 | 持たない（件数のみ） |
| 本日〆 | 予定完了日が今日 | 完了済みも含む |
| 今週 | 予定期間（P〜Q列）が今週の月〜日に重なる | 完了済みも含む |
| 来週 | 予定期間が来週の月〜日に重なる | 完了済みも含む |
| TODO | 未着手（日付は問わない） | 持たない（件数のみ） |

状態は、実績完了日（S列）があれば完了、備考（O列）に `▲` があれば保留、
実績開始日（R列）があれば対応中、それ以外は未着手です。1つのタスクが複数の行に入ることがあります。

---

## JSONの形式（`wbs-tasks/v1`）

`export-wbs.ts` が返す平文です。日付は Excel のシリアル値のままです。

```json
{
  "schema": "wbs-tasks/v1",
  "updatedAt": "2026-09-29T08:40:00.000Z",
  "tasks": [
    { "id": "Y列", "title": "Z列", "category": "A列", "classification": "B列",
      "user": "N列", "note": "O列", "start": 46293, "end": 46295,
      "actualStart": 46293, "actualEnd": "", "row": 15 }
  ]
}
```

## 暗号化の形式（`kanban-aesgcm/v1`）

JSONBin に置かれるのは次の封筒だけです。

```json
{
  "enc": "kanban-aesgcm/v1",
  "kdf": "PBKDF2-SHA256",
  "iter": 250000,
  "salt": "base64（16バイト）",
  "iv": "base64（12バイト）",
  "ct": "base64（暗号文＋認証タグ16バイト）"
}
```

- 鍵はパスフレーズ（UTF-8）から PBKDF2-SHA256、`iter` 回、`salt` で作る 256bit 鍵です。
- 暗号は AES-256-GCM です。追加認証データ（AAD）に `"kanban-aesgcm/v1"` を使います。
- 平文は `wbs-tasks/v1` のJSON文字列（UTF-8）です。
- `salt` と `iv` は登録のたびに作り直します。同じ内容でも暗号文は毎回変わります。
- パスフレーズが違う場合や1バイトでも書き換えられた場合は、復号できずに `badpass` になります。

形式は `kanban.api.js`（`decrypt`/`encrypt`）と `jsonbin_sync.py`（`encrypt_payload`）の
2か所にあります。変えるときは必ず両方を直し、`enc` の版を上げてください。

**パスフレーズについて**
- 長くしてください（単語を4〜5個つなげる程度）。PBKDF2 は総当たりを遅くするだけなので、短いと破られます。
- 変えるときは、Mac 側の値を変えたあとにダッシュボードのカードで入れ直します。
  それまでの間、カードは「パスフレーズが違います」と表示します。

---

## セットアップ

### 1. Power Automate

1. OneDrive に `Apps/kanban-sync/` を作り、中身が `{}` の `wbs-tasks.json` を置きます。
   **wbsのブックとは別のフォルダ**にしてください。同じフォルダだと出力がトリガーを呼び、フローが無限ループします。
2. Excel で「自動化」→「新しいスクリプト」を開き、`office-script/export-wbs.ts` を貼り付けて保存します。
3. 次の3ステップでフローを作ります。
   - トリガー：OneDrive for Business「ファイルが変更されたとき」（ブックのあるフォルダ）
   - Excel Online (Business)「スクリプトの実行」：`export-wbs.ts`
   - OneDrive for Business「ファイルの更新」：`wbs-tasks.json`、内容はスクリプトの `result`
4. トリガーの設定で「コンカレンシー制御」をオンにし、並列処理の度合いを 1 にします。

### 2. Mac（暗号化して登録）

`cryptography` が要ります。

```bash
pip3 install cryptography
```

設定ファイルは**リポジトリの外**に置きます（`api_key` は書き込みのできる X-Master-Key です）。

```json
{
  "id": "kanban-wbs",
  "source_file": "/Users/yuya/Library/CloudStorage/OneDrive-…/Apps/kanban-sync/wbs-tasks.json",
  "bin_id": "6512abcd…",
  "api_key": "$2a$10$…（X-Master-Key）",
  "encrypt": true,
  "passphrase_file": "~/.config/kanban-sync/passphrase",
  "require_schema": "wbs-tasks/v1"
}
```

| キー | 内容 |
|---|---|
| `encrypt` | `true` で暗号化して登録する |
| `passphrase_env` | パスフレーズを入れた環境変数の名前（こちらが優先） |
| `passphrase_file` | パスフレーズを1行だけ書いたファイル。`chmod 600` にしておく |
| `require_schema` | この schema でない内容は送らない（書きかけのファイル対策） |
| `state_file` | 前回送った内容のハッシュの置き場。既定は `~/.cache/jsonbin_sync/{id}.json` |

```bash
mkdir -p ~/.config/kanban-sync
printf '%s' '長いパスフレーズ' > ~/.config/kanban-sync/passphrase
chmod 600 ~/.config/kanban-sync/passphrase
python3 mac/scripts/jsonbin/jsonbin_sync.py --config ~/.config/kanban-sync/jsonbin.json
```

暗号文は毎回変わるので、Bin の中身とは比べません。前回送った**平文のハッシュ**を手元に残し、
同じなら送りません。`updatedAt` が前回より古い内容も送りません（同期の順番が前後したとき用）。

OneDrive のフォルダは Finder で「常にこのデバイス上に保持」にしておいてください。
オンデマンドのままだとファイルの実体が無く、読み取りに失敗します。

### 3. JSONBin

- 中身は暗号文なので、Bin は**公開**にしても構いません。そうすればダッシュボードに読み取りキーを置かずに済みます。
- 非公開にする場合は、読み取り専用の Access Key を作ってダッシュボードに設定します。
  X-Master-Key はダッシュボードに入れないでください（設定は `index.json` に入ります）。

### 4. ダッシュボード

設定 →監視 →種類「WBSカンバン（遅延・今週・来週）」で追加します。書き方は `dashboard/README.md` を見てください。
