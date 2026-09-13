# イヤホン切替(btswitch) セットアップ

Soundcore Liberty 2 Pro は同時に1台としかつながらないので、
「使いたい端末を選ぶ → いま使っている端末が切断する → 選んだ端末がつなぐ」
という手順を画面から起こせるようにする。

```
ブラウザ(index.html : PC / スマホどこからでも)
        │ 「この端末で使う」
        ▼
      GAS  ← いまどの端末が使うか、だけを持つ掲示板
        ▲ 10秒ごとに確認 + 生存報告
        │
  各PCの常駐エージェント(agent/)── Bluetooth 接続 / 切断
```

ブラウザから他の端末の Bluetooth は操作できないため、実際の接続 / 切断は
各PCで動く常駐スクリプトが行う。スマホにはエージェントを置けないので、
「スマホ」を選ぶと **すべてのPCが切断** し、スマホ側が自動でつながる状態にする。

## 1. GAS

`beta/btswitch/gas/Code.gs` を新しいGASプロジェクトへ貼り付ける。

- スクリプトプロパティ `ACCESS_TOKEN` … 合言葉(自分で決める。画面とエージェントで共通)
- デプロイ: ウェブアプリ / 実行するユーザー: 自分 / アクセス: **全員**
- 発行された `/exec` のURLを控える

## 2. 画面

`beta/btswitch/index.html` を開き、歯車から設定する。

- GAS ウェブアプリURL
- 合言葉(`ACCESS_TOKEN` と同じ)
- この端末のID(そのPCのエージェントと同じ `DeviceId`。スマホから見るだけなら空でよい)

スマホのホーム画面に追加しておくと、手元から切り替えられる。

## 3. Windows のPC

イヤホンのMACアドレスを調べる(設定 → Bluetooth → デバイス詳細、または `btdiscovery`)。

**管理者権限なしで使いたい場合**は Bluetooth Command Line Tools を入れる(推奨)。
<https://bluetoothinstaller.com/bluetooth-command-line-tools>

```powershell
cd beta\btswitch\agent
.\bt_agent.ps1 -GasUrl https://script.google.com/macros/s/xxx/exec `
               -Token himitsu -DeviceId pc-home -Label "自宅デスクトップ" `
               -Mac 00:11:22:33:44:55
```

- `btcom.exe` が見つかればそれを使う(管理者権限なし)
- 見つからない場合は PnP デバイスの無効化 / 有効化で代用する → **管理者として実行**が必要

自動起動はタスクスケジューラで「ログオン時」に上記コマンドを登録する
(`powershell -WindowStyle Hidden -File C:\path\bt_agent.ps1 -GasUrl ... `)。

## 4. Mac

```bash
brew install blueutil
blueutil --paired                    # イヤホンのMACを確認
BTSWITCH_GAS_URL=https://script.google.com/macros/s/xxx/exec \
BTSWITCH_TOKEN=himitsu \
BTSWITCH_DEVICE_ID=mac-mini \
BTSWITCH_LABEL="Mac mini" \
BTSWITCH_MAC=00-11-22-33-44-55 \
beta/btswitch/agent/bt_agent.sh
```

常駐させる場合は launchd に登録する(`mac/scripts/plaud/SETUP.md` の plist が参考になる)。

## 5. スマホ

エージェントは不要。画面で「スマホ」または「すべて切断」を選ぶと、
PC側が切断してイヤホンが空くので、スマホのBluetooth設定から(または自動で)つながる。

## 使い方

1. 画面を開く(PCでもスマホでも可)
2. 使いたい端末のカードを押す
3. 10秒以内に、いま使っている端末が切断して切り替わる

カードには「使用中」「接続あり / 未接続」「常駐OK / 常駐が止まっています」が出る。
常駐が止まっている端末は切り替えても反応しないので、そのPCでエージェントを起動する。

## うまくいかないとき

| 症状 | 確認 |
| --- | --- |
| 切り替わらない | そのPCのカードが「常駐OK」か。止まっていればエージェントを起動する |
| Windowsで切断できない | `btcom.exe` が無い場合は管理者として実行しているか |
| 接続はするが音が出ない | 出力先が Liberty 2 Pro になっているか(Windowsは切替後に既定デバイスが戻ることがある) |
| スマホがつながらない | イヤホンをケースに戻して開け直すと、空いている端末を探しに行く |
