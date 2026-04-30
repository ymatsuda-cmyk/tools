# Scenario-tracker

シナリオ一覧の作成状況・バグ影響を管理する Excel Office アドインです。

## ディレクトリ構成

```
Scenario-tracker/
├── manifest.xml            # Office アドイン マニフェスト
├── index.html              # GitHub Pages ランディングページ
├── README.md
├── assets/
│   ├── icon-16.png         # ★ 要追加
│   ├── icon-32.png         # ★ 要追加
│   └── icon-80.png         # ★ 要追加
├── src/
│   └── parser.js           # Excel読み取り・（済）書き戻しロジック
└── taskpane/
    ├── taskpane.html       # メインUI（デモモード内蔵）
    └── commands.html
```

## デプロイ手順

### 1. このフォルダを GitHub にプッシュ

```bash
cd "C:\Users\yuuya\OneDrive - 株式会社ライターム\src\tools2\tools\addin\Scenario-tracker"

git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/Scenario-tracker.git
git push -u origin main
```

### 2. GitHub Pages 有効化

GitHub リポジトリ → **Settings** → **Pages**
- Source: `Deploy from a branch`
- Branch: `main` / `/ (root)`
- **Save**

公開URL: `https://YOUR_USERNAME.github.io/Scenario-tracker/`

### 3. manifest.xml のURL更新

`manifest.xml` 内の `YOUR_GITHUB_USERNAME` を実際のユーザー名に一括置換してプッシュ。

```powershell
# PowerShell で一括置換
(Get-Content manifest.xml) -replace 'YOUR_GITHUB_USERNAME', 'あなたのユーザー名' | Set-Content manifest.xml
```

### 4. アイコン画像を追加

`assets/` フォルダに PNG 画像を配置：
- `icon-16.png` (16×16px)
- `icon-32.png` (32×32px)
- `icon-80.png` (80×80px)

### 5. Excel へアドイン登録

**Excel** → 挿入 → アドイン → **アドインのアップロード** → `manifest.xml` を選択

---

## デモモードについて

`taskpane.html` をブラウザで直接開いた場合（Excel外）は自動的にデモモードで起動します。

- 黄色バナー「⚠️ デモモードでプレビュー中」が表示される
- 各ブランド・各フェーズごとのサンプルデータ（27件）で表示

| ブランド | フェーズ | 件数 |
|---------|---------|------|
| 楽天Edy  | PH1     | 7件  |
| iD       | PH1     | 4件  |
| 交通系IC  | PH1     | 6件  |
| QUICPay  | PH2     | 2件  |
| WAON     | PH2     | 2件  |
| nanaco   | PH2     | 2件  |
| クレジット | PH2    | 2件  |
| 銀聯      | PH2    | 2件  |

---

## 列定義（parser.js）

| シート | フェーズ列 | 完了阻害列 | 軽微列 | 開始日列 | 完了日列 |
|--------|----------|----------|-------|---------|---------|
| 異常（通常） | 96 | 87 | 88 | 90 | 91 |
| 異常（電源断） | 70 | 61 | 62 | 64 | 65 |
| 異常（通信断） | 89 | 80 | 81 | 83 | 84 |
| 正常（クレ・銀聯） | 20 | — | — | 14 | 15 |

※ 列番号はすべて 0-indexed
