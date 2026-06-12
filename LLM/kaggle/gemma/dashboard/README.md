# Kaggle GPU Dashboard

GitHub Pages + GitHub Actions で Kaggle GPU の起動/停止を管理するダッシュボード。

## 構成

```
.
├── docs/
│   ├── index.html      # ダッシュボード（GitHub Pages）
│   └── status.json     # 状態ファイル（Actionsが自動更新）
└── .github/workflows/
    ├── kaggle-start.yml
    └── kaggle-stop.yml
```

## セットアップ手順

### 1. リポジトリの設定

既存リポジトリにこのファイル群を追加:

```bash
cp -r kaggle-dashboard/. your-repo/
cd your-repo
git add .
git commit -m "Add Kaggle GPU dashboard"
git push
```

### 2. GitHub Pages を有効化

`Settings → Pages → Source: Deploy from branch → Branch: main → /docs`

### 3. GitHub Secrets を設定

`Settings → Secrets and variables → Actions → New repository secret`

| Secret名 | 値 |
|---|---|
| `KAGGLE_USERNAME` | あなたのKaggleユーザー名（例: `ymatsuda2025`）|
| `KAGGLE_KEY` | Kaggle API Key（kaggle.jsonのkey）|
| `KAGGLE_KERNEL_SLUG` | ノートブックのスラッグ（例: `kaggle-ollama`）|
| `TUNNEL_URL` | Cloudflare TunnelのURL（例: `https://xxxx.trycloudflare.com`）|

### 4. ダッシュボードを開いてトークンを設定

1. `https://[username].github.io/[repo]/` にアクセス
2. ⚙ Settings を開く
3. GitHub Personal Access Token を入力（スコープ: `repo` + `workflow`）
4. Repository を入力（例: `ymatsuda2025/kaggle-gpu`）

### 5. 動作確認

- **Start Session** → kaggle-start.yml が起動 → status.json が `running` に更新 → ページに反映
- **Stop Session**  → kaggle-stop.yml が起動 → 稼働時間を累計して `stopped` に更新

## 注意事項

- GitHub Actionsのワークフロー起動には **30秒〜2分**のラグがある
- Kaggle APIのstop（途中停止）はカーネルのAPI次第で動作しない場合がある → ノートブック内タイムアウトと併用推奨
- Cloudflare Tunnelの固定URLがある場合は `TUNNEL_URL` Secretに設定すると起動時に自動反映される
- `weekly_hours` は手動リセットが必要（月曜朝にGitHubから直接 status.json を編集）

## Cron 自動起動（オプション）

Mac側から自動起動したい場合は `crontab -e` に追加:

```cron
# 毎朝9時に起動 (GitHub Actions経由)
0 9 * * 1-5 curl -s -X POST \
  -H "Authorization: Bearer $(cat ~/.gh_token)" \
  -H "Accept: application/vnd.github+json" \
  https://api.github.com/repos/YOUR_REPO/actions/workflows/kaggle-start.yml/dispatches \
  -d '{"ref":"main"}' >> /tmp/kaggle-cron.log 2>&1
```

`~/.gh_token` にトークンを保存しておくこと。
