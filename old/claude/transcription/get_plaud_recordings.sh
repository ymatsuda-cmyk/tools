#!/bin/bash

OUTPUT_FILE="plaud_recordings.json"

echo "▶ Plaud 録音一覧を取得しています..."
echo "▶ 出力ファイル: ${OUTPUT_FILE}"

plaud files list \
  --all \
  --json \
  > "${OUTPUT_FILE}"

if [ $? -eq 0 ]; then
  echo "✅ 取得完了"
else
  echo "❌ エラーが発生しました"
fi