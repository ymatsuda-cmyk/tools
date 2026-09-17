# 受信メール取得API

受信トレイ（トップフォルダ）のメールを取ってくる共通の入口です。取得元は差し替えられ、
Outlook（Microsoft Graph）と、サインインなしで使えるJSON決め打ちの2つが入っています。
アカウントは何個でも並べられます。サブフォルダは見ません。

```
api/mail/
├── mail.api.js            呼び出す側が使う入口
├── auth-redirect.html     サインインのポップアップが戻ってくる先
└── providers/
    ├── outlook.js         Microsoft Graph（受信トレイのみ）
    └── json.js            サインイン不要。固定のメールや共有JSON
```

Outlook はブラウザから直接 Microsoft にサインインします。JSON はサインインもサーバーも要りません。

---

## 使い方

```javascript
const mail = await import('/api/mail/mail.api.js')

const accounts = mail.normalizeAccounts([
  { id: 'work', label: '仕事', provider: 'outlook', clientId: '…', tenant: 'organizations' },
  { id: 'home', label: '個人', provider: 'outlook', clientId: '…', tenant: 'consumers' },
  { id: 'share', label: '共有箱', provider: 'json', maxAgeDays: 3,
    mails: [{ subject: 'お知らせ', from: 'info@example.com', receivedAt: '2026-09-16T10:00:00+09:00' }] }
])

await mail.signIn(accounts[0])                    // ポップアップが開く（jsonは不要）
const result = await mail.recent(accounts, { maxAgeDays: 7 })
```

### 関数

| 関数 | 内容 |
|---|---|
| `normalizeAccounts(list)` | 設定の書き漏らし（provider・label・色）を埋める |
| `recent(accounts, { maxAgeDays })` | 期間内（既定7日）の受信トレイのメールをまとめて取る |
| `status(accounts)` | サインイン状態だけを見る |
| `signIn(account)` / `signOut(account)` | ポップアップでサインイン・サインアウト |
| `providers()` | 選べる取得元と、設定に必要な項目 |
| `redirectUri()` | Azure に登録するURI |
| `formatTime` | 表示のための小物 |

### `recent()` が返すもの

```javascript
{
  mails: [{
    key, id, accountId, accountLabel, color, provider,
    subject, from, receivedAt: Date, isRead, preview, url
  }],
  accounts: [{ id, label, color, signedIn, username, count, error, needsSignIn }],
  signedIn: true,
  errors: [{ id, message }]
}
```

メールは受信日時が新しい順に並びます。アカウントが1つ落ちても、残りのメールはそのまま返ります。
理由は `accounts[].error` に入ります。

期間の絞り込みは `maxAgeDays`（何日前まで有効か）で行います。アカウントごとに
`account.maxAgeDays` を指定すると、そのアカウントだけ既定値を上書きできます。

---

## Outlook の準備

### 1. アプリを登録する

[Azure Portal](https://portal.azure.com) → Microsoft Entra ID → アプリの登録 → 新規登録

| 項目 | 値 |
|---|---|
| 名前 | 何でもよい（例 `dashboard-mail`） |
| サポートされているアカウントの種類 | 職場と個人の両方を使うなら **任意の組織 + 個人の Microsoft アカウント** |
| リダイレクト URI | プラットフォームは **シングルページ アプリケーション (SPA)**、URI は下記 |

リダイレクト URI は、このAPIを置いた場所の `auth-redirect.html` です。

```
https://ユーザー名.github.io/リポジトリ名/api/mail/auth-redirect.html
```

`redirectUri()` を呼べば、いま必要なURLがそのまま返ります。
ローカルで試すときは、その時のURL（`http://localhost:5500/...`）も追加してください。

> **SPA を選んでください。** 「Web」で登録するとクライアントシークレットを求められ、
> ブラウザだけでは動きません。

### 2. アクセス許可

API のアクセス許可 → Microsoft Graph → **委任されたアクセス許可**

- `User.Read`
- `Mail.Read`

どちらも個人が同意できる範囲なので、管理者の同意は要りません
（職場のテナントで同意を制限している場合は管理者に依頼してください）。

### 3. アカウントを2つ使う

**アプリ登録は1つで足ります。** `clientId` を同じにしたまま、`id` と `label` だけ
変えて2件並べ、それぞれでサインインしてください。サインイン時にアカウントの
選択画面が出るので、別々のアカウントを選びます。

`api/schedule/README.md` の同じ項目と同様、職場のテナントが外部アプリを制限している
場合だけ、そのテナント側でもう1つアプリを登録してください。

---

## JSON で決め打ちする（`provider: "json"`）

サインインが要りません。Outlook に無い受信箱（共有メールボックスの転記など）を
足したいときに使います。書き方は2通りです。

### 1. アカウントの中に直接書く

```json
{ "id": "share", "label": "共有箱", "provider": "json", "maxAgeDays": 3,
  "mails": [
    { "subject": "お知らせ", "from": "info@example.com", "receivedAt": "2026-09-16T10:00:00+09:00" },
    { "subject": "請求書", "from": "billing@example.com", "receivedAt": "2026-09-15T09:00:00+09:00", "isRead": false }
  ]
}
```

### 2. 外部のJSONファイルから取る

```json
{ "id": "share", "label": "共有箱", "provider": "json",
  "url": "https://example.com/mails.json" }
```

URL先のJSONは、次のどの形でも構いません。

| 形 | 例 |
|---|---|
| 配列 | `[ { "subject": "…", "receivedAt": "…" }, … ]` |
| 包んだオブジェクト | `{ "mails": [ {...} ] }`（`events`/`items`/`data`/`value`/`values` でも可） |
| メール1件そのもの | `{ "subject": "…", "receivedAt": "…" }` |
| NDJSON（1行1件） | `{...}` を改行区切りで並べたもの。GASなどが1件ずつ追記していく運用向け |

`receivedAt` に**タイムゾーンの表記が無い**場合は **UTC として扱います**。
日本時間で書きたいときは `+09:00` を付けてください（`2026-09-16T13:30:00+09:00`）。

### 1件の書き方

| キー | 必須 | 内容 |
|---|---|---|
| `subject`（`title` でも可） | 任意 | 省略すると「(件名なし)」 |
| `receivedAt`（`receivedDateTime`/`date` でも可） | ほぼ必須 | ISO形式の日時 |
| `from`（`sender` でも可） | 任意 | 送信者の表示 |
| `isRead` | 任意 | `false` にすると未読（太字）表示になる。省略時は既読扱い |
| `preview`（`bodyPreview` でも可） | 任意 | 使わないが将来のために受け取る |
| `url`（`webLink` でも可） | 任意 | クリック先（未使用） |
| `id` | 任意 | 省略すると配列内の順番から自動で振る |

---

## カード側の機能（ダッシュボード）

- **期間**: カードの「何日前まで有効か」（既定7日）より古いメールは出しません。
  アカウントの `maxAgeDays` があればそちらを優先します。
- **削除アイコン（✕）**: そのメールをカードから非表示にします（ブラウザのlocalStorageに記録。
  メール自体は削除しません。次回以降も出てきません）。
- **★（優先）**: クリックで ★/☆ を切り替え、色を付けて一覧の先頭に出します。
  こちらもlocalStorageに記録するだけで、メール自体は変更しません。

---

## 取得元を足す

1. `providers/新しい名前.js` を作り、`outlook.js` か `json.js` と同じものを公開する
   - `id` `label` `FIELDS` `redirectUri()` `status(account)` `signIn(account)`
     `signOut(account)` `messages(account, {since})`
2. `messages()` はメールの配列を返す。1件の形は
   `{ id, subject, from, receivedAt: Date, isRead, preview, url }`
3. `mail.api.js` の `LOADERS` に1行足す

アカウント名・色・並べ替え・エラーのまとめは `mail.api.js` 側でやるので、
取得元はその取得元固有の事情だけを持てば済みます。
