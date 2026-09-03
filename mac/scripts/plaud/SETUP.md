# audio-inbox セットアップ

入り口を3つ用意し、すべて Drive の同じフォルダに集約する。Mac mini はローカル
フォルダを見るだけなので Drive API のコードは不要。

```
minutes-viewer (upload.js) ─┐
iOS 共有シート → ドライブ ─┼→ Drive /audio-inbox/ ─ ミラー ─→ ~/Google Drive/.../audio-inbox
Finder にドラッグ ─────────┘                                          │
                                                    drive_inbox.py（15分毎）
                                                    文字起こし → Notion → 削除
```

段階的に進める。**1 と 2 だけで入り口が2つ増える。** 3 は後回しでよい。

---

## 1. Drive for desktop をミラーモードに

1. Google Drive for desktop をインストール
2. 設定 → Google ドライブ → **「マイドライブをミラーリング」** を選択
   （ストリーミングだと実体がなくプレースホルダになるので必ずミラー）
3. Drive の Web で `マイドライブ/audio-inbox` フォルダを作成
4. `~/Google Drive/マイドライブ/audio-inbox` が現れることを確認

パスが違う場合（英語表記など）は `~/.plaud_notion_sync.env` に書く。

```
AUDIO_INBOX=/Users/ymatsuda/Google Drive/My Drive/audio-inbox
```

## 2. drive_inbox.py を設置

```bash
cp drive_inbox.py ~/scripts/
mkdir -p ~/logs

# .env に Notion のプロパティ名を実際の DB に合わせて追記
cat >> ~/.plaud_notion_sync.env <<'EOF'
AUDIO_INBOX=/Users/ymatsuda/Google Drive/マイドライブ/audio-inbox
PROP_TITLE=名前
PROP_DATE=日付
PROP_CATEGORY=カテゴリー
PROP_PERMISSION=権限
PROP_STATUS=状態
STATUS_DONE=文字起こし
EOF
```

まず空振りで確認する。

```bash
DRY_RUN=1 python ~/scripts/drive_inbox.py
```

`処理対象なし` が出れば inbox のパスは正しい。次に音声を1本置いて、
検出されるか（`>> ファイル名` が出るか）を見る。

初回は削除せず残す設定で通す。

```bash
KEEP_DONE=1 python ~/scripts/drive_inbox.py
```

Notion に正しく入ったら `KEEP_DONE` を外す。**削除はゴミ箱を経由しない**ので、
動作が安定するまでは `KEEP_DONE=1` のままにしておくのが安全。

### launchd 登録

```bash
cp com.rtarm.drive-inbox.plist ~/Library/LaunchAgents/
# plist 内の python パスを実環境に合わせて編集してから
launchctl unload ~/Library/LaunchAgents/com.rtarm.drive-inbox.plist 2>/dev/null
launchctl load  ~/Library/LaunchAgents/com.rtarm.drive-inbox.plist
launchctl list | grep drive-inbox
tail -f ~/logs/drive-inbox.log
```

15分毎に起動する。実行が重なった場合は `/tmp/drive_inbox.lock` で二重処理を防ぐ。

### index.json の再生成

`refresh_index()` は `plaud_transcribe_notion.py` の関数を名前で探している。
ログに「index 再生成の関数が見つかりません」と出たら、実際の関数名を
`refresh_index()` 内のタプルに追記する。

## 3. minutes-viewer にアップロード UI を足す（任意）

ここだけ Google Cloud の設定が必要。無料。

### OAuth クライアント作成

1. console.cloud.google.com → 新しいプロジェクト `minutes-uploader`
2. 「APIとサービス」→ ライブラリ → **Google Drive API** を有効化
3. 「OAuth同意画面」
   - User Type: **外部**
   - スコープに `.../auth/drive.file` を追加
   - テストユーザーに自分の Gmail を追加（公開申請は不要）
4. 「認証情報」→ OAuth クライアント ID
   - 種類: **ウェブアプリケーション**
   - 承認済みの JavaScript 生成元: `https://ymatsuda-cmyk.github.io`
     （パスは付けない。オリジンのみ）
   - リダイレクト URI: 空のまま
5. 発行された `xxxx.apps.googleusercontent.com` をコピー

### フォルダ ID の取得

Drive で `audio-inbox` を開き、URL の `/folders/` 以降をコピー。

> `drive.file` スコープは自アプリが作ったファイルしか見えないため、
> フォルダの検索はできない。ID は手で設定に入れる必要がある。

### 組み込み

`index.html` の `<head>` に1行。

```html
<script src="https://accounts.google.com/gsi/client" async defer></script>
```

アップローダを置きたい場所に。

```html
<div id="uploader"></div>
<script type="module">
  import { initUploader } from './upload.js';
  initUploader({ mount: document.getElementById('uploader') });
</script>
```

初回はフォームの「接続設定」を開いてクライアント ID とフォルダ ID を入力。
`minutes:upload` として localStorage に保存され、カテゴリーと権限の前回値も
記憶される。

---

## 入り口2（iOS 共有シート）の運用

JSON を作れないので、日時はファイル名か更新日時から推定する。
名前を付けるなら次の形にすると日時とタイトルが正確に入る。

```
20260903-1400_GMO定例.m4a
```

付けなくてもファイルの更新日時が使われるので、事故にはならない。

## トラブル時

| 症状 | 見るところ |
|---|---|
| 検出されない | `AUDIO_INBOX` のパス、ミラーモードか、更新から60秒経ったか |
| 途中のファイルを拾う | `STABLE_SECONDS` を増やす（既定60） |
| Notion 400 エラー | プロパティ名と型（status / select / multi_select）の一致 |
| 同じ内容が2件登録される | ロックファイルの残骸。`rm /tmp/drive_inbox.lock` |
| 文字起こしが空 | `failed/*.error.txt` にトレースバックが残る |
| ブラウザで Location エラー | 拡張機能を無効化。5MB 以下なら自動で multipart に切替 |
