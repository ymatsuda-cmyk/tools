# 予定取得API

今日の予定を取ってくる共通の入口です。取得元は差し替えられ、Outlook（Microsoft Graph）と、
サインインなしで使えるJSON決め打ちの2つが入っています。アカウントは何個でも並べられます。

```
api/schedule/
├── schedule.api.js        呼び出す側が使う入口
├── auth-redirect.html     サインインのポップアップが戻ってくる先
└── providers/
    ├── outlook.js         Microsoft Graph
    └── json.js            サインイン不要。固定の予定や共有JSON
```

Outlook はブラウザから直接 Microsoft にサインインします。JSON はサインインもサーバーも要りません。

---

## 使い方

```javascript
const schedule = await import('/api/schedule/schedule.api.js')

const accounts = schedule.normalizeAccounts([
  { id: 'work', label: '仕事', provider: 'outlook', clientId: '…', tenant: 'organizations' },
  { id: 'home', label: '個人', provider: 'outlook', clientId: '…', tenant: 'consumers' },
  { id: 'plan', label: '予定表', provider: 'json',
    events: [{ title: '健康診断', allDay: true, start: '2026-09-20' }] }
])

await schedule.signIn(accounts[0])        // ポップアップが開く（jsonは不要）
const result = await schedule.today(accounts)
```

### 関数

| 関数 | 内容 |
|---|---|
| `normalizeAccounts(list)` | 設定の書き漏らし（provider・label・色）を埋める |
| `today(accounts)` | 今日 0:00〜翌 0:00 の予定 |
| `range(accounts, {from, to})` | 期間を指定して取る |
| `status(accounts)` | サインイン状態だけを見る |
| `signIn(account)` / `signOut(account)` | ポップアップでサインイン・サインアウト |
| `providers()` | 選べる取得元と、設定に必要な項目 |
| `redirectUri()` | Azure に登録するURI |
| `formatTime` / `nextEvent` / `ongoing` | 表示のための小物 |

### `today()` が返すもの

```javascript
{
  from, to, dayKey: '2026-09-16',
  signedIn: true,
  events: [{
    key, id, accountId, accountLabel, color, provider,
    title, start: Date, end: Date, allDay, cancelled, free, declined,
    location, organizer, onlineUrl, url
  }],
  accounts: [{ id, label, color, signedIn, username, count, error, needsSignIn }],
  errors: [{ id, message }]
}
```

予定は開始が早い順に並びます（終日は先頭）。アカウントが1つ落ちても、
残りの予定はそのまま返ります。理由は `accounts[].error` に入ります。

---

## Outlook の準備

### 1. アプリを登録する

[Azure Portal](https://portal.azure.com) → Microsoft Entra ID → アプリの登録 → 新規登録

| 項目 | 値 |
|---|---|
| 名前 | 何でもよい（例 `dashboard-schedule`） |
| サポートされているアカウントの種類 | 職場と個人の両方を使うなら **任意の組織 + 個人の Microsoft アカウント** |
| リダイレクト URI | プラットフォームは **シングルページ アプリケーション (SPA)**、URI は下記 |

リダイレクト URI は、このAPIを置いた場所の `auth-redirect.html` です。

```
https://ユーザー名.github.io/リポジトリ名/api/schedule/auth-redirect.html
```

`redirectUri()` を呼べば、いま必要なURLがそのまま返ります。
ローカルで試すときは、その時のURL（`http://localhost:5500/...`）も追加してください。

> **SPA を選んでください。** 「Web」で登録するとクライアントシークレットを求められ、
> ブラウザだけでは動きません。

### 2. アクセス許可

API のアクセス許可 → Microsoft Graph → **委任されたアクセス許可**

- `User.Read`
- `Calendars.Read`

どちらも個人が同意できる範囲なので、管理者の同意は要りません
（職場のテナントで同意を制限している場合は管理者に依頼してください）。

### 3. アカウントを2つ使う

**アプリ登録は1つで足ります。** `clientId` を同じにしたまま、`id` と `label` だけ
変えて2件並べ、それぞれでサインインしてください。サインイン時にアカウントの
選択画面が出るので、別々のアカウントを選びます。

職場のテナントが外部アプリを制限している場合だけ、そのテナント側でもう1つ
アプリを登録し、`clientId` と `tenant` を分けてください。

| `tenant` | 対象 |
|---|---|
| `common` | 職場/学校と個人の両方（既定） |
| `organizations` | 職場/学校のみ |
| `consumers` | 個人（outlook.com / hotmail.com）のみ |
| テナントID | そのテナントのみ |

サインインの状態はブラウザに保存されます（MSAL の localStorage）。
端末やブラウザを変えたときは、もう一度サインインしてください。

---

## JSON で決め打ちする（`provider: "json"`）

サインインが要りません。承認待ちの間や、Outlook に無い予定（旅行の日程など）を
足したいときに使います。書き方は2通りです。

### 1. アカウントの中に直接書く

```json
{ "id": "plan", "label": "予定表", "provider": "json",
  "events": [
    { "title": "健康診断", "allDay": true, "start": "2026-09-20" },
    { "title": "歯医者", "start": "2026-09-16T19:00:00+09:00", "end": "2026-09-16T20:00:00+09:00", "location": "駅前" }
  ]
}
```

### 2. 外部のJSONファイルから取る

```json
{ "id": "plan", "label": "予定表", "provider": "json",
  "url": "https://example.com/events.json" }
```

URL先のJSONは、次のどの形でも構いません。

| 形 | 例 |
|---|---|
| 配列 | `[ { "title": "…", "start": "…" }, … ]` |
| 包んだオブジェクト | `{ "events": [ {...} ] }`（`items`/`data`/`value`/`values` でも可） |
| 予定1件そのもの | `{ "title": "…", "start": "…" }` |
| NDJSON（1行1件） | `{...}` を改行区切りで並べたもの。GASなどが1件ずつ追記していく運用向け |

それ以外の形（配列を持たないオブジェクトなど）だと
「予定のJSONを読み取れません」というエラーになります。

`start` / `end` に**タイムゾーンの表記が無い**場合（`2026-09-16T04:30:00.0000000` のような
Outlookの生の応答そのもの）は **UTC として扱います**。日本時間で書きたいときは
`+09:00` を付けてください（`2026-09-16T13:30:00+09:00`）。

`title` の代わりに `subject`、`allDay` の代わりに `isAllDay`、`start`/`end` が
`{ "dateTime": "…" }` の形（Graph APIの生の応答そのもの）でも読めます。

### 1件の書き方

| キー | 必須 | 内容 |
|---|---|---|
| `title`（`subject` でも可） | 任意 | 省略すると「(件名なし)」 |
| `start` | ほぼ必須 | 終日なら `"2026-09-20"`、時刻ありなら `"2026-09-16T19:00:00+09:00"` |
| `end` | 任意 | 省略すると `start` と同じ（終日は無し） |
| `allDay`（`isAllDay` でも可） | 任意 | `true` で終日扱い |
| `location` | 任意 | 表示に出る |
| `id` | 任意 | 省略すると配列内の順番から自動で振る |

---

## 取得元を足す

1. `providers/新しい名前.js` を作り、`outlook.js` か `json.js` と同じものを公開する
   - `id` `label` `FIELDS` `redirectUri()` `status(account)` `signIn(account)`
     `signOut(account)` `events(account, {from, to})`
2. `events()` は予定の配列を返す。1件の形は
   `{ id, title, start: Date, end: Date, allDay, cancelled, location, organizer, onlineUrl, url }`
3. `schedule.api.js` の `LOADERS` に1行足す

アカウント名・色・並べ替え・エラーのまとめは `schedule.api.js` 側でやるので、
取得元はその取得元固有の事情だけを持てば済みます。
