# inkboard 手書きボードAPI

iPad の Apple Pencil・指・マウスで描ける手書きボードの共通APIです。
用紙固定（A3／A4／A5／B5／16:9／4:3、縦横）と無限キャンバスに対応し、
参照表示（埋め込み先の小さな表示）をタップすると全画面の編集画面が開きます。

```
api/inkboard/
├── inkboard.js    呼び出す側が使う入口（ESモジュール。依存ライブラリなし）
├── inkboard.css   参照表示と全画面の見た目
└── index.html     確認用のデモ
```

## 使い方

```js
import { createBoard, renderBoard } from 'https://cdn.jsdelivr.net/gh/ymatsuda-cmyk/tools@main/api/inkboard/inkboard.js'
// CSS も読み込む: https://cdn.jsdelivr.net/gh/ymatsuda-cmyk/tools@main/api/inkboard/inkboard.css

const value = createBoard({ paper: 'A4', orientation: 'portrait' })  // 無限キャンバスは { paper: 'infinite' }
const board = renderBoard(el, value, { onChange: (v) => save(v) })
```

`onChange` を渡さなければ閲覧のみ（全画面でも描けない）になります。

## API

| 関数 | 内容 |
|---|---|
| `createBoard({ paper, orientation, background })` | 新しい値を作る。`paper` は `infinite` `A3` `A4` `A5` `B5` `W169` `W43`、`background` は `plain` `grid` `lines` `dots` |
| `renderBoard(container, value, options)` | 参照表示を描く。戻り値は `{ value, update(v), open(), destroy() }` |
| `openBoard(value, options)` | 全画面の編集画面だけを開く。閉じたときの値で解決する Promise を返す |
| `previewArea(value)` | 参照表示で見せる範囲 `{x, y, w, h}`。埋め込み先の枠の縦横比を合わせるのに使う |
| `exportImage(value, { area, scale, type })` | 画像（既定は PNG）の Blob を作る |
| `parseBoard(value)` ／ `serializeBoard(data)` | 値とデータの相互変換 |

`renderBoard` の options：`onChange` `onOpen` `onClose` `clickToOpen`（既定 true）`emptyText` `title` `changeDelay`（既定 800ms）。
`onChange` は線を書くたびではなく、書き終えてから `changeDelay` 待って呼びます。閉じたとき・タブを隠したときは待たずに呼びます。

## 値の形式

改行を含まない1行のJSONです。Markdown のコードブロックや表のセルにそのまま置けます。

```json
{"v":1,"mode":"paper","paper":"A4","orient":"portrait","bg":"plain","view":[40,80,600,340],"s":[["pen","#1f1f1f",3,"AAg…"]]}
```

- `view` は参照表示の範囲（世界座標）。`null` なら用紙全体、無限キャンバスは描いた線の全体を表示します
- `s` は線の一覧 `[道具, 色, 太さ, 点列]`。点列は「x, y, 筆圧」を前の点との差分にして、1文字5ビットの可変長で詰めた文字列です（座標は 1/4px 単位、筆圧は 0〜63）

## 編集画面の操作

| 操作 | 方法 |
|---|---|
| 描く | Apple Pencil（筆圧で太さが変わる）、マウス。「指で描く」がオンなら指1本でも |
| 移動 | 指1本（「指で描く」がオフのとき）、2本指、Space＋ドラッグ、ホイール、「移動」ツール |
| 拡大縮小 | 2本指のピンチ、⌘／Ctrl＋ホイール、＋／−、0 で全体表示 |
| 道具 | P ペン、M マーカー、E 消しゴム（線ごと消す）、R 参照表示の範囲、H 移動 |
| 元に戻す | ⌘Z、やり直す ⇧⌘Z |
| 閉じる | 完了、Esc |

Apple Pencil を一度使うと、指は自動で移動・拡大縮小に切り替わります（手のひらで線が引かれないように）。
Pencil で描いている最中と直後の指の入力も無視します。
