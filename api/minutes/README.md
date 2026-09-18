# 議事録（Notion）取得API

議事録アプリ（`minutes/`）が Notion に貯めている一覧を、ダッシュボードなど他の画面からも
使えるようにする薄い入口です。一覧は GitHub Pages に公開された `data/minutes/index.json` を
読むだけで、サインインは要りません。権限（Notionの「権限」multi_select）の更新だけ、
議事録用GAS（`minutes/gas/Code.gs`）の `savePermissions` アクションを呼びます。

```
api/minutes/
└── minutes.api.js   呼び出す側が使う入口
```

---

## 使い方

```javascript
const minutes = await import('/api/minutes/minutes.api.js')

// 直近10件（date降順）
const { items } = await minutes.recent({ limit: 10 })

// 権限の割り当て・更新（GASのURLと共有トークンが必要）
await minutes.savePermissions(
  { gasUrl: 'https://script.google.com/macros/s/.../exec', token: '...' },
  [items[0].notionPageId],
  ['jba'],
  'add' // 'add' | 'remove' | 'replace'
)
```

### 関数

| 関数 | 内容 |
|---|---|
| `recent({ limit, indexUrl })` | 新着順に並べた一覧を取る（既定 直近10件）。認証不要 |
| `savePermissions(config, pageIds, permissions, mode)` | 権限をまとめて更新する。`config` に `gasUrl` と `token` が必要 |
| `saveTitle(config, pageId, title)` | ミーティング名（タイトル）を更新する。`config` に `gasUrl` と `token` が必要 |

### `recent()` が返すもの

`data/minutes/index.json` の1件がそのまま入ります。

```javascript
{
  items: [{
    key, title, date, duration, status, tags, permissions, notionPageId, updatedAt
  }]
}
```

### GASの準備

`minutes/gas/Code.gs` をウェブアプリとしてデプロイし、スクリプトプロパティ `ACCESS_TOKEN` に
共有トークンを設定してください。`config.token` はこの値と一致させます。詳しくは
`minutes/gas/Code.gs` 冒頭のコメントを参照してください。
