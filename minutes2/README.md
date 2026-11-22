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
- 端末に覚えておく「合言葉」：権限コードで一度開くと、合言葉を決められます（6文字以上。設定からも変更・解除できます）。権限コードは、合言葉から PBKDF2 で作った鍵で暗号化して、その端末の localStorage にだけ置きます。合言葉も権限コードも、GitHub や GAS には保存しません（権限コードはログインのときだけ GAS に送ります）。端末を人に使われる心配があるときは、推測されにくい合言葉にしてください。権限コードを `rotate` で替えたら、覚えていた分は使えなくなるので、新しいコードで開き直して合言葉を決め直します。
- 権限コードは GAS 側で鍵付きハッシュ（HMAC）にして照合します。10分に30回失敗すると、しばらくログインを止めます。
- GAS の `setup` は `SETUP_TOKEN` が一致したときだけ値を登録します（5回失敗すると10分止まります）。値を読み出す手段は用意していません。
- `SETUP_TOKEN` を替えるときは、`setup_keys.py token` で作り直し、GAS と `.env` の両方を書き換えます。
- 閲覧グループから外しても、外す前にその端末が読み込んだ内容は残ることがあります。人が抜けたときは、そのグループのコードを `rotate` してください。
- 音声ファイルは Drive の受付フォルダに暗号化せずに置かれます（ファイル名はランダムな ID）。

## 第1段階に入っていないもの（第2段階の候補）
- マインドマップの作成（今は Notion から取り込んだ分を、本文に表示するだけ）、要約の手直し編集、学習ルール、横断チャット
- 時刻を押したときの音声再生（今は文字起こしの該当行へ移動）
- 管理画面からの権限コード変更（今は Mac の `setup_keys.py`）
