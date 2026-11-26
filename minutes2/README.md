# 議事録 v2（第1段階）

ClipStock と同じ作りで、議事録を暗号化して GitHub の private リポジトリに保管します。
権限コードでログインし、閲覧できるグループの議事録だけが開けます。

## 中身

| フォルダ | 置き場所 | 役割 |
|---|---|---|
| `index.html`・`src/`・`css/` | `ymatsuda-cmyk/tools` の `minutes2/`（GitHub Pages） | アプリ本体（一覧・詳細・変換管理・登録・設定） |
| `gas/Code.gs`・`appsscript.json` | 新しい GAS プロジェクト（今の minutes の GAS とは別） | ログイン、音声の中継、暗号化済みデータの書き込み、鍵束の保管（スクリプトプロパティ） |
| `mac/` | Mac mini | 鍵の準備（`setup_keys.py`）、5分おきの監視（`minutes_worker.py`）、Notion 議事録DBからの取り込み（`notion_import.py`） |
| `skill/minutes2-claude-summarizer/` | Claude.ai のスキル | 要約待ちを Claude.ai で要約して Notion に書く |

データ用リポジトリ（例：`ymatsuda-cmyk/minutes-data`、private）に置かれるもの：

| ファイル | 中身 |
|---|---|
| `index.json` | 一覧：ID・日時・長さ・状態・登録元、会議の鍵（グループ公開鍵で包んだもの）、暗号化したタイトルとタグ |
| `pub.json` | グループの公開鍵（秘密ではない） |
| `m/<ID>/t.json` `s.json` `n.json` | 文字起こし・要約・メモ（すべて会議の鍵で暗号化） |
| `worker.json` | Mac の最終確認時刻 |

## 準備（初回だけ）

### 1. GitHub
1. private リポジトリ `minutes-data` を作る（空のままでよい）。
2. fine-grained トークンを2つ作る。どちらも対象は `minutes-data` だけ。
   - 書き込み用：Contents = Read and write（Mac と GAS が使う）
   - 読み取り用：Contents = Read-only（鍵束に入れて各端末へ配る）

### 2. GAS を作る（まだ値は入れない）
1. 新しいプロジェクトを作り、`gas/Code.gs` を貼る。プロジェクトの設定で「appsscript.json をエディタで表示」をオンにし、`gas/appsscript.json` の中身に置き換える（Drive の書き込み権限を確実に付けるため）。
2. デプロイ → ウェブアプリ（実行ユーザー：自分、アクセス：全員）。初回に Drive の許可を求められます。
3. 表示された URL（…/exec）を控える。

### 3. Mac：鍵を作って GAS に登録する
```sh
mkdir -p ~/.minutes2 && cp mac/.env.example ~/.minutes2/.env && chmod 600 ~/.minutes2/.env
python3 -m venv ~/.minutes2/venv && ~/.minutes2/venv/bin/pip install -r mac/requirements.txt
cd mac && ~/.minutes2/venv/bin/python setup_keys.py token
```
1. 表示された値を、GAS のスクリプトプロパティ **`SETUP_TOKEN`** と、`.env` の `GAS_SETUP_TOKEN` の両方に入れる。
   **GAS に手で入れるのはこれだけです。**
2. `.env` の残り（`GH_*`・`GAS_URL`・`AUDIO_FOLDER_ID`・`DRIVE_DIR` など）を書く。
3. 鍵を作って登録する：
   ```sh
   ~/.minutes2/venv/bin/python setup_keys.py init --roles jba,kijun,xYz
   ~/.minutes2/venv/bin/python setup_keys.py check     # 全部「ok」なら完了
   ```
   - `--roles` には今の権限（role）をそのまま並べます。管理者 `xYz` は必須です。
   - 鍵束・コードのハッシュ・秘密値・GitHub の書き込みトークンなどを、GAS の `setup` に送って
     スクリプトプロパティへ登録します（値はどこにも表示しません）。
   - 最後に「新しい権限コード」が表示されます。**この画面でしか表示されません。** 各グループの人に安全な方法で渡してください。
   - `~/.minutes2/keys.json` に全グループの秘密鍵が入ります。この Mac の外に出さないでください（暗号化したバックアップは可）。

### 4. アプリ

1. `src/config.js` の `GAS_URL` に、手順2の URL を入れる。
2. この `minutes2/` フォルダを `tools` リポジトリに push すると、GitHub Pages で公開されます。`https://ymatsuda-cmyk.github.io/tools/minutes2/` で開きます。

### 5. Mac の監視
1. Google ドライブ for desktop で、音声の受付フォルダを Mac に同期する（`DRIVE_DIR` に設定した場所）。
2. `mac/com.minutes2.worker.plist` のパスを書き換えて `~/Library/LaunchAgents/` に置き、
   `launchctl load ~/Library/LaunchAgents/com.minutes2.worker.plist`
3. ログは `~/.minutes2/worker.log`。アプリの変換管理に「Mac mini 監視中」と出れば動いています。

### 6. Notion 議事録DBからの取り込み（PLAUD Note の受け口）
Notion 議事録DBは、PLAUD Note からの登録口として今までどおり使います（plaud-notion-sync・export・文字起こしはそのまま）。
Mac が5分おきに取り込み、暗号化して GitHub に入れます。

| Notion の状態 | 取り込み後のアプリ | 取り込むもの |
|---|---|---|
| 文字起こし | 要約待ち | 文字起こし（本文） |
| 要約 | 要約済み | 文字起こし＋要約・決定事項・ToDo（完了も）・論点・議事・メモ |

1. Notion のインテグレーションを議事録DBに接続し、`.env` に `NOTION_TOKEN` と `NOTION_MINUTES_DB_ID` を書く。
2. まず確認だけ：`python3 notion_import.py --dry-run`（何件・どの権限で入るかを表示。何も書きません）
3. 過去分を少しずつ移す：`python3 notion_import.py --limit 20` → アプリで中身を確かめてから、残りを `python3 notion_import.py`
4. 以降は `minutes_worker.py` が毎回、新しく「文字起こし」「要約」になったページを取り込みます。

- 閲覧グループは「権限」の値をそのまま role として使います。対応する role が無い値は外して、ログに残します。権限が空なら管理者だけが閲覧できます。
- 取り込んだページには「取込ID」が入り、状態は「取込済み」になります（「取込ID」の列は最初の実行で自動で追加されます）。
- `NOTION_CLEAR_AFTER_IMPORT=1`（初期値）なら、取り込み後に Notion の本文と要約・メモの欄を消します。
  ミーティング名・日時・会議時間・権限・カテゴリーなどの管理用の項目は残ります。消したくないときは `0` にするか、`--keep` を付けて実行します。
- 本文を消すので、今の minutes アプリと今の要約スキル（notion-transcript-summarizer）では、取り込み済みの会議は読めなくなります。要約は新しいアプリ（URL指定AI・Claude.ai）で行います。

### 7. Claude.ai で要約（使うときだけ）
1. Notion に作業DB「議事録 要約作業」を作る。プロパティ：名前（タイトル）／ID（テキスト）／状態（セレクト：要約待ち・要約済み）／メモ（テキスト）
2. 同じインテグレーションを作業DBにも接続し、`NOTION_WORK_DB_ID` を `.env` に書く。
3. `skill/minutes2-claude-summarizer/SKILL.md` の DB ID を書き換え、フォルダを zip にして Claude.ai のスキルに追加する。
4. アプリの設定 →「Claude.ai で要約するページ」に、使うプロジェクトのリンクを入れる。

Notion には要約待ちの文字起こしが一時的に（平文で）置かれ、要約を回収したら Mac が消します。

## 運用

| やりたいこと | 方法 |
|---|---|
| 権限コードを変える | `setup_keys.py rotate --role jba`（GAS への登録まで自動。古いコードはその時点で使えなくなる） |
| グループを増やす | `setup_keys.py add-role --role eigyo`（管理者の権限コードを聞かれます） |
| 読み取りトークンを替える | `.env` を書き換えて `setup_keys.py rebuild`（全グループのコードを聞かれます） |
| 書き込みトークン・受付フォルダを替える | `.env` を書き換えて `setup_keys.py push` |
| GAS の登録を確かめる | `setup_keys.py check`（キーの有無だけ表示） |
| 閲覧グループを変える | アプリで管理者が詳細 →「閲覧グループ」 |

## 安全のための注意
- 権限コード・ペッパー・トークンをソースやコメントに書かないでください（今の `minutes/gas/Code.gs` の `verifyCode_` のコメントは削除し、今のコードは入れ替えを）。
- 設定の保管：Claude.ai のページ・文字の大きさ・一覧の表示・ToDo の期間・お気に入りは、データ用リポジトリの `settings/<グループID>.json` に暗号化して置きます（会議と同じく、使い捨ての鍵をグループの公開鍵で包む）。書き込みは GAS の `putSettings` 経由。同じグループの端末には、開いたときに新しいほうが反映されます。AI の接続先（URL・API キー）は個人の設定なので含めず、その端末の localStorage にだけ置きます。
- 他端末へ設定共有：設定 →「他端末へ設定共有」で、QR・6桁コード（`common/devicelink.js`・中継は `common/gas/devicelink-relay.gs`）・リンクの3通りで権限コードと、その端末の AI の接続先を渡します（受け取る端末にすでに接続先があれば上書きしません）。中身はその場で作る10文字のパスフレーズで暗号化し、受け取る端末でそのパスフレーズを入れます。ホーム画面に追加したアプリで受け取るときはコードを使います。
- 権限コードは GAS 側で鍵付きハッシュ（HMAC）にして照合します。10分に30回失敗すると、しばらくログインを止めます。
- GAS の `setup` は `SETUP_TOKEN` が一致したときだけ値を登録します（5回失敗すると10分止まります）。値を読み出す手段は用意していません。
- `SETUP_TOKEN` を替えるときは、`setup_keys.py token` で作り直し、GAS と `.env` の両方を書き換えます。
- 閲覧グループから外しても、外す前にその端末が読み込んだ内容は残ることがあります。人が抜けたときは、そのグループのコードを `rotate` してください。
- 音声ファイルは Drive の受付フォルダに暗号化せずに置かれます（ファイル名はランダムな ID）。

## 第1段階に入っていないもの（第2段階の候補）
- マインドマップの作成（今は Notion から取り込んだ分を、本文に表示するだけ）、要約の手直し編集、学習ルール、横断チャット
- 時刻を押したときの音声再生（今は文字起こしの該当行へ移動）
- 管理画面からの権限コード変更（今は Mac の `setup_keys.py`）

## 閲覧範囲・タグ・マインドマップ

- 閲覧範囲（管理者だけ）：一覧の題名の前に、管理者のみ（鍵）か共有（人のアイコンとグループ数）を出します。一覧の上の「範囲：すべて／管理者のみ／共有」で絞り込めます。詳細の見出しに「閲覧できる：…」と「変更」を出し、閲覧グループを選び直せます。
- タグ：詳細の見出しの「＋ タグ」で、付いているタグ・この会議に合いそうなタグ（ほかの会議のタグのうち、題名や要約に出てくるもの）・よく使うタグを押して付け外しします。押すとすぐ保存します。新しいタグもその場で追加できます。
- マインドマップ：共通部品 `api/mindmap2` で描きます（markmap は `api/mindmap2/vendor/` に版を固定して置いたものを読むので、CSP の script-src は 'self' のまま）。markmap が style を差し込むため、style-src に 'unsafe-inline' を足しています。管理者と登録した人は、項目を選んでボタン（またはキー）で直せ、直すと自動で `s.mindmap` に保存します。「要約から作り直す」で保存した地図を消せます。

## ログイン（初回：権限コード＋パスコード／2回目以降：パスコード）

- 初回（この端末に保管庫が無い）：権限コード（グループを決める）と、この端末のパスコード（8文字以上）を設定します。権限コードは GAS で確かめ、合っていれば保管庫を作って開きます。
- 2回目以降：パスコードだけで開きます（Face ID・Touch ID を登録していればそれでも開けます）。開いたあと、権限コードで GAS にログインします。権限コードが変わっていたときは、新しい権限コードだけを聞き直します（パスコードはそのまま）。
- 保管庫（`src/vault.js`）：権限コード・AI の接続先（URL・API キー）を、パスコードから作る鍵（PBKDF2-SHA256 600,000 回 → AES-GCM 256）で暗号化して localStorage の `m2.vault` に置きます。パスコードと鍵はメモリにだけ置き、GAS には送りません。文字の大きさなど表示の設定は暗号化しません。
- 誤入力が5回続くと保管庫を消して初回の画面へ戻します（会議のデータは GitHub にあるので、権限コードを入れ直せば戻れます）。パスコードを忘れたときも、「パスコードを忘れた／権限コードを入れ直す」から同じ手順です。
- 放置したら自動で鍵をかけます（既定15分・設定で 5／15／60分／かけない）。編集中・要約中はかけません。鍵をかけると、画面に描いた議事録も消します。
- AI の接続先は MinuteStock 専用（保管庫の中）です。gemma-chat など他のアプリと共有している `gemma-chat.settings` には書きません。以前の版からの引っ越しでは、共有の保存に残っていた接続情報を保管庫へコピーします（共有側は残るので、設定の AI タブの「ここから消す」で消せます）。
- 以前の版（権限コードを端末の鍵で覚えて自動で開く方式）の控えがあれば、初回の画面で権限コードは聞かず、パスコードだけ決めてもらい、控えは消します。
- 他端末へ設定共有：受け取った端末では、その端末のパスコードを決めます（権限コードと、選んだときは AI の接続先が入ります）。
- 注意：パスコードが短いと、保存データをコピーされたときに破られます（8文字以上にしている理由です）。

