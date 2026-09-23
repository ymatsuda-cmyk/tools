# 顧客一覧API（customer9）

`/customer9` の顧客一覧ページが使う、顧客データ取得用のAPIです。取得元は差し替えられ、
GAS（Google Apps Script Web App）と、GAS未設定でも確認できるJSON（組み込みサンプル）の
2つが入っています。

```
api/customer9/
├── customer9.api.js       呼び出す側が使う入口
└── providers/
    ├── gas.js              GAS Web App（スプレッドシート連携）
    └── json.js             サインイン・デプロイ不要。固定の顧客データや共有JSON
```

---

## 使い方

```javascript
const customer9 = await import('/api/customer9/customer9.api.js')

// GAS未設定でも動く（組み込みサンプルの顧客一覧が返る）
const demo = await customer9.list(customer9.normalizeConfig({}))

// GAS Web Appを使う場合
const config = customer9.normalizeConfig({
  provider: 'gas',
  url: 'https://script.google.com/macros/s/xxx/exec'
})
const { customers } = await customer9.list(config)
```

### 関数

| 関数 | 内容 |
|---|---|
| `normalizeConfig(raw)` | 設定の書き漏らし（provider・url）を埋める |
| `list(config)` | 顧客一覧を取得する。`{ provider, customers }` を返す |
| `providers()` | 選べる取得元と、設定に必要な項目 |

### `list()` が返す顧客の形

```javascript
{ code: 'C001', name: '株式会社サンプル商事', email: '…', phone: '…', status: '取引中' }
```

---

## GAS Web App の準備（`gas` プロバイダを使う場合）

1. 顧客データを持つスプレッドシートを用意する（1行目はヘッダー。列例: `code, name, email, phone, status`）
2. [Google Apps Script](https://script.google.com) で新規プロジェクトを作成し、スプレッドシートに紐づける
3. `Code.gs` に以下を貼り付ける（シート名は必要に応じて書き換える）

```javascript
// 顧客一覧を返すWeb App
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('customers');
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();

  const customers = rows.map(row => {
    const item = {};
    headers.forEach((key, i) => { item[key] = row[i]; });
    return item;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ customers: customers }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選び、アクセスできるユーザーを設定して発行する
5. 発行された `https://script.google.com/macros/s/.../exec` を `/customer9` ページの設定に貼り付ける

CORSの制約でうまく取得できない場合は、代わりに `json` プロバイダの「JSONのURL」に
同じデータを公開したファイルのURLを指定する方法でも確認できます。
