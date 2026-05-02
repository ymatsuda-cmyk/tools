# Google Sheets 連携セットアップ手順

## 概要
貸借トラッカーのデータを Google Sheets に保存するための設定手順です。
所要時間：約15分

---

## Step 1 — Google スプレッドシートを作成

1. https://sheets.google.com を開く
2. 「空白のスプレッドシート」を新規作成
3. 名前を「貸借トラッカー」などに変更
4. **シートタブを2枚** 用意する
   - シート1のタブ名を `loans` に変更
   - 「＋」でシートを追加し、タブ名を `repayments` に変更

### loans シートのヘッダー（1行目）
A1: id / B1: name / C1: amount / D1: date / E1: monthly / F1: memo

### repayments シートのヘッダー（1行目）
A1: id / B1: loanId / C1: date / D1: amount / E1: memo

---

## Step 2 — Apps Script を設定

1. メニューの「拡張機能」→「Apps Script」を開く
2. エディタの中身を全削除し、下記コードを貼り付けて保存（Ctrl+S）

```javascript
const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  return handleReq(e);
}
function doPost(e) {
  return handleReq(e);
}

function handleReq(e) {
  const p = e.parameter || {};
  const b = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  const params = Object.assign({}, p, b);
  const action = params.action;
  let result;
  try {
    if (action === 'getLoans')          result = getLoans();
    else if (action === 'addLoan')      result = addLoan(params);
    else if (action === 'getRepayments') result = getRepayments(params.loanId);
    else if (action === 'addRepayment') result = addRepayment(params);
    else result = { error: 'unknown action' };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLoans() {
  const sh = SS.getSheetByName('loans');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    id: r[0], name: r[1], amount: r[2], date: r[3],
    monthly: r[4], memo: r[5]
  }));
}

function addLoan(p) {
  const sh = SS.getSheetByName('loans');
  const id = 'L' + Date.now();
  sh.appendRow([id, p.name, Number(p.amount), p.date, Number(p.monthly) || 0, p.memo || '']);
  return { id };
}

function getRepayments(loanId) {
  const sh = SS.getSheetByName('repayments');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1)
    .filter(r => r[1] === loanId)
    .map(r => ({ id: r[0], loanId: r[1], date: r[2], amount: r[3], memo: r[4] }));
}

function addRepayment(p) {
  const sh = SS.getSheetByName('repayments');
  const id = 'R' + Date.now();
  sh.appendRow([id, p.loanId, p.date, Number(p.amount), p.memo || '']);
  return { id };
}
```

---

## Step 3 — ウェブアプリとしてデプロイ

1. 「デプロイ」→「新しいデプロイ」をクリック
2. 種類の選択：**ウェブアプリ** を選択
3. 以下のように設定する：
   - 説明：貸借トラッカー（任意）
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**全員**
4. 「デプロイ」をクリック
5. Googleアカウントの権限確認が出たら「許可」する
6. 表示された **ウェブアプリURL** をコピーする

---

## Step 4 — アプリに URL を設定

1. 貸借トラッカーアプリを開く
2. 黄色い設定バナーの「設定する」をクリック（または「同期」ボタン）
3. コピーした URL を「GAS Web App URL」欄に貼り付ける
https://script.google.com/macros/s/AKfycbzUwIxPzPt1I5bnzD6s_e-tTd0SiNc6V7-qnOkkUItZeDtC2MXQS91gY5WjX_gZQjyi/exec
4. 「設定完了」をクリック
5. 「同期」ボタンを押して動作確認

---

## GitHub Pages への公開手順

1. GitHub にリポジトリを作成（例：`loan-tracker`）
2. `loan-tracker.html` を `index.html` にリネームしてアップロード
3. Settings → Pages → Source を「main ブランチ」に設定
4. `https://yourusername.github.io/loan-tracker/` で公開完了

---

## 注意事項

- GAS Web App URL はブラウザの localStorageに保存されます
- 別のブラウザや端末で使う場合は再度 URL の設定が必要です
- URL を知っている人なら誰でもアクセスできるため、機密性の高い情報は入力しないでください
- GAS未設定の場合でも、ブラウザのメモリ上でアプリは動作します（ページを閉じるとデータは消えます）
