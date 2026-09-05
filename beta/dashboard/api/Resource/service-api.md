# サービス側API仕様

稼働状況タブから監視・操作するために、各サービスがこの仕様のエンドポイントを1つ用意します。
仕様さえ合っていれば、実装言語やホスティング先は問いません。

ダッシュボードは直接このAPIを呼ばず、必ずGASを経由します。
そのため **CORS設定は不要** で、APIを外部に広く公開する必要もありません。

扱うのは次の2つだけです。

- **稼働時間の残量** — リングゲージで表示
- **起動 / 停止** — トグルスイッチで操作

---

## エンドポイント

1つのURLで、GETとPOSTを受け付けます。

| メソッド | 用途 |
|---|---|
| `GET  {endpoint}?action=status` | 現在の状態と残量を返す |
| `POST {endpoint}` | 起動 / 停止を実行する |

認証は `Authorization: Bearer {トークン}` ヘッダで行います。
トークンはGASのスクリプトプロパティに登録され、リクエスト時に自動で付きます。

---

## GET — 状態と残量を返す

### レスポンス（200）

```json
{
  "name": "GPU サーバー",
  "state": "running",
  "remaining": { "value": 4.5, "max": 30, "unit": "h" },
  "updatedAt": "2026-09-03T10:00:00Z"
}
```

| フィールド | 必須 | 内容 |
|---|---|---|
| `name` | 任意 | 表示名。省略時は登録時の名前を使う |
| `state` | **必須** | `running` `starting` `stopping` `stopped` `error` のいずれか |
| `remaining` | 推奨 | 稼働時間の残量。省略するとゲージの代わりに状態名だけ出る |
| `updatedAt` | 任意 | 計測時刻（ISO 8601） |

### remaining

| フィールド | 必須 | 内容 |
|---|---|---|
| `value` | **必須** | 残り時間。**使用済みではなく残量** |
| `max` | **必須** | 上限。`value / max` で割合を計算する |
| `unit` | 任意 | 単位。省略時は `h` |

ゲージの中央には残量そのもの（`4.5h` など）が出ます。
色は残量の割合で自動的に変わります。

| 残量 | 色 |
|---|---|
| 30%超 | 緑 |
| 10〜30% | 黄 |
| 10%以下 | 赤 |

残量という概念がないサービスは `remaining` を省略してください。
状態名とトグルだけのカードになります。

---

## POST — 起動 / 停止

### リクエスト

```json
{ "action": "start" }
```

`action` は `start` か `stop` のみです。

### レスポンス（200 または 202）

```json
{
  "ok": true,
  "state": "starting"
}
```

失敗時は、HTTPステータスをエラーにするか、次のように返します。

```json
{
  "ok": false,
  "error": "インスタンスが割り当てられていません"
}
```

`error` の文字列はそのまま画面に表示されるので、原因が分かる日本語にしてください。

### 重要な注意

**起動・停止は非同期で構いません。** すぐ完了しない場合は `starting` / `stopping` を返してください。
ダッシュボードは8秒後に自動で再取得し、その後も設定した間隔で状態を追いかけます。

処理中はトグルが自動で無効になるため、二重実行は起きません。
ただしサービス側でも、**同じ操作が重複して届いた場合に安全に無視できる作り**にしておくことを推奨します。

---

## 実装例（Node.js / Express）

```javascript
const express = require('express');
const app = express();
app.use(express.json());

const TOKEN = process.env.MONITOR_TOKEN;

function auth(req, res, next) {
  if (req.get('Authorization') !== 'Bearer ' + TOKEN) {
    return res.status(401).json({ ok: false, error: '認証エラー' });
  }
  next();
}

app.get('/api/resource', auth, async (req, res) => {
  const s = await getServerStats();
  res.json({
    name: 'GPU サーバー',
    state: s.isRunning ? 'running' : 'stopped',
    remaining: { value: s.remainingHours, max: 30, unit: 'h' },
    updatedAt: new Date().toISOString()
  });
});

app.post('/api/resource', auth, async (req, res) => {
  const { action } = req.body;
  if (action !== 'start' && action !== 'stop') {
    return res.status(400).json({ ok: false, error: '不明なaction' });
  }
  try {
    if (action === 'start') await startServer();
    else await stopServer();
    res.status(202).json({ ok: true, state: action === 'start' ? 'starting' : 'stopping' });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.listen(3000);
```

---

## ダッシュボードへの登録

設定 → 監視タブに、以下のJSONで登録します。

```json
[
  {
    "id": "gpu-server",
    "name": "GPU サーバー",
    "group": "本番",
    "endpoint": "https://your-service.example.com/api/resource"
  },
  {
    "id": "dev-server",
    "name": "検証サーバー",
    "group": "検証",
    "endpoint": "https://your-service.example.com/api/dev"
  }
]
```

| フィールド | 内容 |
|---|---|
| `id` | 一意の識別子。トークンのプロパティ名とURLパラメータに使う |
| `name` | 画面に出す名前 |
| `group` | 任意。同じ値を持つもの同士でまとめて絞り込める |
| `endpoint` | GETとPOSTを受けるURL |

---

## URLパラメータで表示を切り替える

特定のサーバーだけを表示した状態でページを開けます。
ブックマークやショートカットから直接その画面を出したいときに使います。

| パラメータ | 例 | 動作 |
|---|---|---|
| `server` | `?server=gpu-server` | そのIDのサーバーだけ表示 |
| `group` | `?group=本番` | そのグループのサーバーだけ表示 |
| `view` | `?view=monitor` | 稼働状況タブを開いた状態で起動 |

組み合わせられます。

```
https://ユーザー名.github.io/リポジトリ名/?view=monitor&server=gpu-server
```

画面上のチップで切り替えると、URLも自動で書き換わります。
そのURLをコピーすれば、同じ表示を再現できます。

`monitor` は `server` の別名として使えます。`?monitor=gpu-server` でも同じ動作です。

## トークンの登録

GASのスクリプトプロパティに、監視IDから作った名前で登録します。

| 監視ID | プロパティ名 |
|---|---|
| `gpu-server` | `MONITOR_TOKEN_GPU_SERVER` |
| `batch-worker` | `MONITOR_TOKEN_BATCH_WORKER` |

英数字以外はアンダースコアに変換し、すべて大文字にします。
トークンが未登録の場合、`Authorization` ヘッダは付きません。

---

## 動作の流れ

1. 稼働状況タブを開くと即座に1回取得し、以降は設定した間隔（既定3分）で更新
2. 他のタブに移ると自動停止し、戻ると再開
3. ブラウザのタブが非表示になっても停止
4. トグルを切り替えると確認ダイアログが出て、実行後8秒で再取得
5. 確認をキャンセルするとトグルは元の位置に戻る

APIが応答しない場合、そのカードだけ「応答がありません」と表示され、
トグルは操作できなくなります。他の監視対象には影響しません。
