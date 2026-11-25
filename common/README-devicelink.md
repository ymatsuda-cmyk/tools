# DeviceLink — 設定を別の端末へ渡す共通部品

アプリの接続情報（リポジトリ・トークンなど）を、パスフレーズで暗号化したまま次の3つの方法で渡す。

| 方法 | 向いている場面 | 中継 |
|---|---|---|
| QR コード | iPad・iPhone を追加する（カメラで読む） | 不要 |
| 6桁コード | PC を追加する、ホーム画面に追加したアプリで受け取る | GAS（5分・1回きり） |
| リンク | 自分の端末どうしで AirDrop などで送る | 不要 |

## ファイル

- `devicelink.js` … ブラウザ側の部品（QR 生成ライブラリ込み）。`window.DeviceLink`
- `devicelink-relay.gs` … 全アプリ共通の中継（Google Apps Script）

## 準備（コード方式を使うとき・最初の1回だけ）

1. `devicelink-relay.gs` を Apps Script に貼り、ウェブアプリとしてデプロイ（実行：自分／アクセス：全員）
2. 表示された `…/exec` の URL を `devicelink.js` の `RELAY_URL` に入れる

## アプリへの組み込み

```html
<script src="../shared/devicelink.js"></script>
```

```js
const APP = 'myapp';   // 英小文字・数字・ハイフン。アプリごとに変える

// 渡す側（設定画面など）：ボタン一式を作る
DeviceLink.panel(document.querySelector('#devLink'), {
  app: APP,
  getPackage: () => DeviceLink.seal({ owner, repo, branch, token }, pass, { app: APP, label: `${owner}/${repo}` })
});

// 受け取る側：リンク／QR で開かれたとき
const pkg = DeviceLink.fromURL(APP);          // #dl= が無ければ null
// 受け取る側：コード
const pkg2 = await DeviceLink.receiveCode(APP, '123456');

// パスフレーズで開く（違えば Error('BADPASS')）
const data = await DeviceLink.open(pkg, pass);
```

受け取った包み（pkg）はパスフレーズ無しでは読めないので、そのまま端末に保存してからパスフレーズを聞いてよい（コードは1回きりのため、MemoStock はそうしている）。

## 安全性

- 暗号化：PBKDF2-SHA256（60万回）＋ AES-GCM。アプリ名とラベルも改ざん検出の対象
- QR・リンク・中継・コードのどこから漏れても、パスフレーズが無ければ中身は読めない
- 中継は、間違ったコードが10分間に10回でそのアプリの受け取りを一時停止
- パスフレーズを持たないアプリでは使わない（短い数字だけでは総当たりで破られる）
