# 予定取得API

今日の予定を取ってくる共通の入口です。取得元は差し替えでき、いまは Outlook
（Microsoft Graph）だけが入っています。アカウントは何個でも並べられます。

```
api/schedule/
├── schedule.api.js        呼び出す側が使う入口
├── auth-redirect.html     サインインのポップアップが戻ってくる先
└── providers/
    └── outlook.js         Microsoft Graph
```

ブラウザから直接 Microsoft にサインインします。GAS もサーバーも要りません。

---

## 使い方

```javascript
const schedule = await import('/api/schedule/schedule.api.js')

const accounts = schedule.normalizeAccounts([
  { id: 'work', label: '仕事', provider: 'outlook', clientId: '…', tenant: 'organizations' },
  { id: 'home', label: '個人', provider: 'outlook', clientId: '…', tenant: 'consumers' }
])

await schedule.signIn(accounts[0])        // ポップアップが開く
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

## 取得元を足す

1. `providers/新しい名前.js` を作り、`outlook.js` と同じものを公開する
   - `id` `label` `FIELDS` `redirectUri()` `status(account)` `signIn(account)`
     `signOut(account)` `events(account, {from, to})`
2. `events()` は予定の配列を返す。1件の形は
   `{ id, title, start: Date, end: Date, allDay, cancelled, location, organizer, onlineUrl, url }`
3. `schedule.api.js` の `LOADERS` に1行足す

アカウント名・色・並べ替え・エラーのまとめは `schedule.api.js` 側でやるので、
取得元はその取得元固有の事情だけを持てば済みます。
