#!/usr/bin/env bash
# イヤホン切替(btswitch)の macOS 常駐エージェント。
#
# GAS に置かれた「いまどの端末が使うか」を一定間隔で見に行き、
# 自分が指名されていれば接続、そうでなければ切断する。
#
# 事前に blueutil が必要:  brew install blueutil
# イヤホンのMACは `blueutil --paired` で確認する。
#
# 使い方:
#   BTSWITCH_GAS_URL=https://script.google.com/macros/s/xxx/exec \
#   BTSWITCH_TOKEN=himitsu \
#   BTSWITCH_DEVICE_ID=mac-mini \
#   BTSWITCH_LABEL="Mac mini" \
#   BTSWITCH_MAC=00-11-22-33-44-55 \
#   ./bt_agent.sh
set -uo pipefail

GAS_URL="${BTSWITCH_GAS_URL:?BTSWITCH_GAS_URL が未設定です}"
TOKEN="${BTSWITCH_TOKEN:?BTSWITCH_TOKEN が未設定です}"
DEVICE_ID="${BTSWITCH_DEVICE_ID:?BTSWITCH_DEVICE_ID が未設定です}"
LABEL="${BTSWITCH_LABEL:-$(scutil --get ComputerName 2>/dev/null || hostname)}"
MAC="${BTSWITCH_MAC:?BTSWITCH_MAC が未設定です (blueutil --paired で確認)}"
INTERVAL="${BTSWITCH_INTERVAL:-10}"

log() { printf '[%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*"; }

gas() { # $1 = JSONのbody(tokenは呼び出し側で入れる)
  curl -sS -L -X POST "$GAS_URL" \
    -H 'Content-Type: text/plain;charset=utf-8' \
    -d "$1"
}

state_owner() {
  gas "{\"action\":\"getState\",\"token\":\"$TOKEN\"}" |
    python3 -c 'import json,sys; d=json.load(sys.stdin); print(d.get("data",{}).get("owner","") if d.get("ok") else "")'
}

is_connected() {
  [ "$(blueutil --is-connected "$MAC" 2>/dev/null)" = "1" ]
}

heartbeat() { # $1 = connected(true/false)
  gas "{\"action\":\"heartbeat\",\"token\":\"$TOKEN\",\"deviceId\":\"$DEVICE_ID\",\"label\":\"$LABEL\",\"connected\":$1,\"note\":\"blueutil\"}" >/dev/null
}

command -v blueutil >/dev/null || { log '!! blueutil が見つかりません (brew install blueutil)'; exit 1; }

log "開始します: $DEVICE_ID ($LABEL) / ${INTERVAL}秒ごと"
last_owner=""

while true; do
  owner="$(state_owner)"
  if [ -n "$owner" ] && [ "$owner" != "$last_owner" ]; then
    log "現在の使用端末: $owner"
    last_owner="$owner"
  fi

  if [ "$owner" = "$DEVICE_ID" ]; then
    if ! is_connected; then
      log '自分の番になりました。接続します'
      blueutil --connect "$MAC" || log '!! 接続に失敗しました'
      sleep 3
    fi
  else
    if is_connected; then
      log "他の端末($owner)に渡します。切断します"
      blueutil --disconnect "$MAC" || log '!! 切断に失敗しました'
      sleep 3
    fi
  fi

  if is_connected; then heartbeat true; else heartbeat false; fi
  sleep "$INTERVAL"
done
