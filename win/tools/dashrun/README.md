# dashrun — ダッシュボードのカードから Windows の exe を起動する

ダッシュボード（`dashboard/index.html`）でリンクの実行方法を「ローカル実行」にして exe のパスを入力すると、
カードのクリックで `dashrun://launch?path=...` が開き、このフォルダのハンドラが exe を起動します。

## ファイル

| ファイル | 役割 |
| --- | --- |
| `dashrun_handler.pyw` | `dashrun://` を受け取るハンドラ。URL を検証して `launch_core` に渡す |
| `launch_core.py` | パス検証・許可済みリスト・許可ダイアログ・起動の共通処理（将来 `local_bridge.py` からも使う） |
| `install_dashrun.ps1` | 登録・解除スクリプト（任意。スマートアプリコントロールで止められる場合は下の貼り付け手順を使う） |
| `test_launch_core.py` | 単体テスト（`python -m unittest test_launch_core.py`） |

## セットアップ（Windows）

Windows 11 のスマートアプリコントロールは、ダウンロードした署名なしの `.bat` / `.ps1` をブロックします。
そのため、登録はスクリプトファイルを使わず PowerShell にコマンドを直接貼り付けて行います（管理者権限は不要）。

1. このフォルダを置き場所を決めてコピーする（例: `C:\Tools\dashrun`）。移動したら再登録が必要です。
2. PowerShell を開き、1 行目の `$dir` を置き場所に合わせてから、まとめて貼り付ける。

   ```powershell
   $dir = 'C:\Tools\dashrun'
   $py  = (Get-Command pythonw.exe).Source
   $key = 'HKCU:\Software\Classes\dashrun'
   New-Item $key -Force | Out-Null
   Set-ItemProperty $key '(default)' 'URL:dashrun Protocol'
   New-ItemProperty $key 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null
   New-Item "$key\shell\open\command" -Force | Out-Null
   Set-ItemProperty "$key\shell\open\command" '(default)' "`"$py`" `"$dir\dashrun_handler.pyw`" `"%1`""
   "登録しました: $py"
   ```

   `pythonw.exe` が見つからないと言われたら、python.org 版の Python をインストールしてください（Store 版は動かないことがあります）。

3. 動作確認。許可ダイアログが出てメモ帳が起動すれば完了です。

   ```powershell
   Start-Process 'dashrun://launch?path=C%3A%5CWindows%5Csystem32%5Cnotepad.exe'
   ```

4. ダッシュボードの「追加」→ 実行方法「ローカル実行」→「exeのパス」に貼り付けて保存。

### 登録の解除

```powershell
Remove-Item 'HKCU:\Software\Classes\dashrun' -Recurse -Force
```

許可済みリスト（`%APPDATA%\dashrun`）は残るので、不要なら手動で削除してください。

### 起動先の exe について

起動する exe そのものもスマートアプリコントロールの判定を受けます。自作の署名なし exe は、dashrun 経由でも
エクスプローラーからでもブロックされることがあります。うまく起動しないときは、まず exe を直接ダブルクリックして確認してください。

## 動き

- 初めてのパスは Windows 側で「起動を許可しますか？」と確認し、許可したものだけ `%APPDATA%\dashrun\approved.json` に記録します。
- 2 回目以降は確認なしで起動します。許可を取り消すときは `approved.json` から該当行を削除してください。
- ブラウザ側でも毎回「dashrun を開きますか？」が出ます。「常に許可」にチェックすると省略できます。
- ログ: `%APPDATA%\dashrun\dashrun.log`

## 制限（安全のため）

- ドライブ文字から始まるローカルの `.exe` のみ。ネットワークパス（`\\server\...`）・`.bat` / `.cmd` / `.ps1`・引数は不可。
- URL に `path` 以外のパラメータが付いていたら起動しません。
- 起動結果はダッシュボードには返りません（トーストは「起動を依頼しました」まで）。

## local_bridge 方式への移行

ダッシュボード側はカードに `path` だけを保存し、起動は `LAUNCHERS`（`index.html`）経由で行っています。
移行時は次の 2 つを実装し、設定の「連携」→「ローカルアプリの起動方式」を切り替えるだけです。

- `local_bridge.py` に `/launch` を追加し、中で `launch_core.handle_launch_request(path)` を呼ぶ
- `index.html` の `LAUNCHERS.bridge.launch()` で `/launch` を呼び、戻り値の `status` をそのまま返す

許可済みリスト（`approved.json`）はそのまま引き継がれます。
