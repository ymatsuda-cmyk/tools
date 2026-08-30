#!/bin/bash
# ============================================================
# git_push_all.sh
# JSONの設定に基づいて複数フォルダを一括コミット＆プッシュする
# repo_subfolder が指定されている場合は、リポジトリをクローンして
# 該当サブフォルダにファイルをコピーしてからpushする
# ============================================================

# set -euo pipefail は使わず、各処理でエラーハンドリングする

# --- 設定 ---
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONFIG_FILE="${SCRIPT_DIR}/git_push_config.json"
LOG_FILE="${SCRIPT_DIR}/git_push_all.log"
WORK_DIR="${HOME}/.git_push_work"
TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')

# --- カラー定義 ---
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# --- ログ関数 ---
log()         { echo -e "$1" | tee -a "$LOG_FILE"; }
log_info()    { log "${BLUE}[INFO]${NC}  $1"; }
log_ok()      { log "${GREEN}[OK]${NC}    $1"; }
log_warn()    { log "${YELLOW}[WARN]${NC}  $1"; }
log_error()   { log "${RED}[ERROR]${NC} $1"; }
log_divider() { log "------------------------------------------------------------"; }

# --- 前提チェック ---
if ! command -v jq &> /dev/null; then
  echo "エラー: jq がインストールされていません。brew install jq で導入してください。"
  exit 1
fi

if [ ! -f "$CONFIG_FILE" ]; then
  log_error "設定ファイルが見つかりません: $CONFIG_FILE"
  exit 1
fi

mkdir -p "$WORK_DIR"

# --- ログ開始 ---
echo "" >> "$LOG_FILE"
log_divider
log_info "実行開始: $TIMESTAMP"
log_divider

# --- カウンタ ---
TOTAL=0; SUCCESS=0; SKIPPED=0; FAILED=0

# ============================================================
# push の競合を吸収するヘルパー
# ============================================================
push_with_retry() {
  local REPO_DIR="$1"
  local REMOTE="$2"
  local BRANCH="$3"
  local MAX_RETRY=3
  local ATTEMPT=1

  while [ "$ATTEMPT" -le "$MAX_RETRY" ]; do
    log_info "  → git push $REMOTE $BRANCH (試行 ${ATTEMPT}/${MAX_RETRY})"
    if git -C "$REPO_DIR" push "$REMOTE" "$BRANCH" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      return 0
    fi

    if [ "$ATTEMPT" -ge "$MAX_RETRY" ]; then
      log_error "  → git push リトライ上限に達しました"
      return 1
    fi

    log_warn "  → push 競合を検出。fetch + rebase して再試行します"
    if ! git -C "$REPO_DIR" fetch "$REMOTE" "$BRANCH" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      log_error "  → 再試行用の git fetch 失敗"
      return 1
    fi

    if ! git -C "$REPO_DIR" rebase "${REMOTE}/${BRANCH}" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      log_error "  → git rebase 失敗（競合の可能性）。push を中断します"
      git -C "$REPO_DIR" rebase --abort >/dev/null 2>&1 || true
      return 1
    fi

    ATTEMPT=$((ATTEMPT + 1))
  done

  return 1
}

# ============================================================
# サブフォルダ付きでpushする関数
# ============================================================
push_with_subfolder() {
  local SOURCE="$1"
  local REPO="$2"
  local SUBFOLDER="$3"
  local BRANCH="$4"
  local REMOTE="$5"
  local MSG="$6"
  local CACHE_DIR="${WORK_DIR}/$(echo "$REPO" | sed 's|[^a-zA-Z0-9]|_|g')"

  # リポジトリのローカルキャッシュを準備
  if [ ! -d "$CACHE_DIR/.git" ]; then
    log_info "  → リポジトリをクローン中..."
    if ! git clone --branch "$BRANCH" "$REPO" "$CACHE_DIR" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      log_error "  → クローン失敗"
      return 1
    fi
  else
    log_info "  → キャッシュを最新に同期 (git fetch + reset)..."
    if ! git -C "$CACHE_DIR" fetch "$REMOTE" "$BRANCH" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      log_error "  → git fetch 失敗"
      return 1
    fi

    # キャッシュは作業用のため、分岐していてもリモートに強制同期する
    if ! git -C "$CACHE_DIR" checkout -B "$BRANCH" "${REMOTE}/${BRANCH}" 2>/dev/null; then
      log_warn "  → トラッキングブランチ作成に失敗。reset で継続します"
      if ! git -C "$CACHE_DIR" checkout "$BRANCH" 2>/dev/null; then
        log_error "  → ブランチ切り替え失敗: $BRANCH"
        return 1
      fi
    fi

    if ! git -C "$CACHE_DIR" reset --hard "${REMOTE}/${BRANCH}" 2>&1 | \
        sed 's/^/    /' | tee -a "$LOG_FILE"; then
      log_error "  → git reset --hard 失敗"
      return 1
    fi
  fi

  # サブフォルダを作成してファイルをコピー
  local TARGET_DIR="${CACHE_DIR}/${SUBFOLDER}"
  mkdir -p "$TARGET_DIR"
  log_info "  → ファイルをコピー: $SOURCE → $TARGET_DIR"
  rsync -av --delete \
    --exclude="._*" \
    --exclude=".DS_Store" \
    --exclude=".Spotlight-V100" \
    --exclude=".Trashes" \
    "$SOURCE/" "$TARGET_DIR/" 2>&1 | \
    sed 's/^/    /' | tee -a "$LOG_FILE"

  # macOSメタデータファイルを念のため削除
  find "$TARGET_DIR" -name "._*" -delete 2>/dev/null || true
  find "$TARGET_DIR" -name ".DS_Store" -delete 2>/dev/null || true

  # .gitignore に除外設定を追加（なければ作成）
  local GITIGNORE="${CACHE_DIR}/.gitignore"
  if ! grep -q "^\._\*" "$GITIGNORE" 2>/dev/null; then
    cat >> "$GITIGNORE" << 'IGNORE'
._*
.DS_Store
.Spotlight-V100
.Trashes
IGNORE
    log_info "  → .gitignore に macOS除外設定を追加"
  fi

  # 変更チェック
  git -C "$CACHE_DIR" add .
  if git -C "$CACHE_DIR" diff --cached --quiet; then
    log_warn "  → 変更なし。コミットをスキップ"
    return 2
  fi

  # コミット＆プッシュ
  local FULL_MSG="${MSG} [$(date '+%Y-%m-%d %H:%M:%S')]"
  log_info "  → git commit: \"$FULL_MSG\""
  git -C "$CACHE_DIR" commit -m "$FULL_MSG"

  if ! push_with_retry "$CACHE_DIR" "$REMOTE" "$BRANCH"; then
    log_error "  → git push 失敗"
    return 1
  fi

  return 0
}

# ============================================================
# リポジトリルートに直接pushする関数
# ============================================================
push_direct() {
  local SOURCE="$1"
  local REPO="$2"
  local BRANCH="$3"
  local REMOTE="$4"
  local MSG="$5"

  if [ ! -d "${SOURCE}/.git" ]; then
    log_info "  → .git が存在しません。初期化してリモートを登録します..."
    git -C "$SOURCE" init
    git -C "$SOURCE" remote add "$REMOTE" "$REPO"
    git -C "$SOURCE" fetch "$REMOTE" "$BRANCH" 2>/dev/null || true
    git -C "$SOURCE" checkout -B "$BRANCH" \
      "${REMOTE}/${BRANCH}" 2>/dev/null || \
      git -C "$SOURCE" checkout -B "$BRANCH" 2>/dev/null || true
  else
    CURRENT_REMOTE=$(git -C "$SOURCE" remote get-url "$REMOTE" 2>/dev/null || echo "")
    if [ "$CURRENT_REMOTE" != "$REPO" ]; then
      log_warn "  → リモートURLを更新: $CURRENT_REMOTE → $REPO"
      git -C "$SOURCE" remote set-url "$REMOTE" "$REPO"
    fi
  fi

  git -C "$SOURCE" add .
  if git -C "$SOURCE" diff --cached --quiet; then
    log_warn "  → 変更なし。コミットをスキップ"
    return 2
  fi

  local FULL_MSG="${MSG} [$(date '+%Y-%m-%d %H:%M:%S')]"
  log_info "  → git commit: \"$FULL_MSG\""
  git -C "$SOURCE" commit -m "$FULL_MSG"

  if ! push_with_retry "$SOURCE" "$REMOTE" "$BRANCH"; then
    log_error "  → git push 失敗"
    return 1
  fi

  return 0
}

# ============================================================
# メイン処理
# ============================================================
ENTRY_COUNT=$(jq 'length' "$CONFIG_FILE")

for i in $(seq 0 $((ENTRY_COUNT - 1))); do
  ENTRY=$(jq ".[$i]" "$CONFIG_FILE")

  ID=$(echo "$ENTRY"          | jq -r '.id')
  DESCRIPTION=$(echo "$ENTRY" | jq -r '.description')
  ENABLED=$(echo "$ENTRY"     | jq -r '.enabled')
  SOURCE=$(echo "$ENTRY"      | jq -r '.source_folder')
  REPO=$(echo "$ENTRY"        | jq -r '.repository')
  SUBFOLDER=$(echo "$ENTRY"   | jq -r '.repo_subfolder // ""')
  BRANCH=$(echo "$ENTRY"      | jq -r '.branch')
  MSG=$(echo "$ENTRY"         | jq -r '.commit_message')
  REMOTE=$(echo "$ENTRY"      | jq -r '.remote_name')

  TOTAL=$((TOTAL + 1))
  log ""
  log_info "[$ID] $DESCRIPTION"

  # enabled チェック
  if [ "$ENABLED" != "true" ]; then
    log_warn "  → enabled=false のためスキップ"
    SKIPPED=$((SKIPPED + 1))
    continue
  fi

  # ソースフォルダ存在チェック
  if [ ! -d "$SOURCE" ]; then
    log_error "  → フォルダが見つかりません: $SOURCE"
    FAILED=$((FAILED + 1))
    continue
  fi

  # push実行（サブフォルダあり or なし）
  if [ -n "$SUBFOLDER" ]; then
    log_info "  → モード: サブフォルダ push ($SUBFOLDER)"
    push_with_subfolder "$SOURCE" "$REPO" "$SUBFOLDER" "$BRANCH" "$REMOTE" "$MSG"
    RESULT=$?
  else
    log_info "  → モード: ダイレクト push"
    push_direct "$SOURCE" "$REPO" "$BRANCH" "$REMOTE" "$MSG"
    RESULT=$?
  fi

  if [ "$RESULT" -eq 0 ]; then
    log_ok "  → プッシュ成功！"
    SUCCESS=$((SUCCESS + 1))
  elif [ "$RESULT" -eq 2 ]; then
    SKIPPED=$((SKIPPED + 1))
  else
    log_error "  → プッシュ失敗（次のエントリに続きます）"
    FAILED=$((FAILED + 1))
  fi
done

# --- サマリー ---
log ""
log_divider
log_info "完了サマリー"
log_info "  合計    : $TOTAL"
log_ok   "  成功    : $SUCCESS"
log_warn "  スキップ: $SKIPPED"
log_error "  失敗    : $FAILED"
log_divider

[ "$FAILED" -eq 0 ] || exit 1