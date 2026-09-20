"""
Phase1: Teams投稿からGitHub Issueを作成し、Teams返信用JSONと
Phase2以降で使用する状態JSONを出力する。

重複防止:
- Teams Message ID単位でprocessedマーカーを作成
- マーカー作成はO_EXCLによる原子的な予約
- on_createdとOneDrive再同期が重なってもIssueは1件だけ作成
- on_movedは処理しない
"""

import html
import json
import os
import pathlib
import shutil
import signal
import sys
import threading
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer


# ============================================================
# 環境変数
# ============================================================

REQUIRED_ENV_VARS = [
    "GITHUB_TOKEN",
    "GITHUB_OWNER",
    "GITHUB_REPO",
    "AGENT_REQUEST_DIR",
    "AGENT_DONE_DIR",
    "AGENT_REPLY_DIR",
    "AGENT_STATE_DIR",
]

ENV_SETUP_EXAMPLES = {
    "GITHUB_TOKEN": 'setx GITHUB_TOKEN "github_pat_xxxxxxxxxxxxx"',
    "GITHUB_OWNER": 'setx GITHUB_OWNER "ymatsuda-cmyk"',
    "GITHUB_REPO": 'setx GITHUB_REPO "tools"',
    "AGENT_REQUEST_DIR": (
        'setx AGENT_REQUEST_DIR "C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\request"'
    ),
    "AGENT_DONE_DIR": (
        'setx AGENT_DONE_DIR "C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\done"'
    ),
    "AGENT_REPLY_DIR": (
        'setx AGENT_REPLY_DIR "C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\reply"'
    ),
    "AGENT_STATE_DIR": (
        'setx AGENT_STATE_DIR "C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\state"'
    ),
}


def mask_secret(secret: str) -> str:
    if len(secret) <= 8:
        return "********"
    return f"{secret[:4]}{'*' * 8}{secret[-4:]}"


def validate_environment_variables() -> None:
    missing_variables = []

    print()
    print("=" * 60)
    print("環境変数チェック")
    print("=" * 60)

    for variable_name in REQUIRED_ENV_VARS:
        value = os.getenv(variable_name)

        if value is None or not value.strip():
            missing_variables.append(variable_name)
            print(f"[NG] {variable_name}: 未設定")
            continue

        if variable_name == "GITHUB_TOKEN":
            print(f"[OK] {variable_name}: 設定済み ({mask_secret(value)})")
        else:
            print(f"[OK] {variable_name}: {value}")

    if not missing_variables:
        print("=" * 60)
        print("必要な環境変数はすべて設定されています。")
        print("=" * 60)
        print()
        return

    print()
    print("不足している環境変数があります。")
    for variable_name in missing_variables:
        print(f"- {variable_name}")

    print()
    print("PowerShellで以下を実行してください。")
    print()
    for variable_name in missing_variables:
        print(ENV_SETUP_EXAMPLES[variable_name])

    print()
    print("設定後にPowerShellを閉じ、新しいPowerShellを開いてください。")
    sys.exit(1)


validate_environment_variables()

GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_OWNER = os.environ["GITHUB_OWNER"]
GITHUB_REPO = os.environ["GITHUB_REPO"]
WATCH_DIR = pathlib.Path(os.environ["AGENT_REQUEST_DIR"])
DONE_DIR = pathlib.Path(os.environ["AGENT_DONE_DIR"])
REPLY_DIR = pathlib.Path(os.environ["AGENT_REPLY_DIR"])
STATE_DIR = pathlib.Path(os.environ["AGENT_STATE_DIR"])
PROCESSED_DIR = STATE_DIR / "processed"

GITHUB_API_URL = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPO}/issues"
)


# ============================================================
# 動作設定
# ============================================================

FILE_READY_RETRIES = 15
FILE_READY_INTERVAL_SECONDS = 1
HTTP_TIMEOUT_SECONDS = 30
HTTP_RETRIES = 3
HTTP_RETRY_INTERVAL_SECONDS = 3

MESSAGE_ID_KEYS = (
    "messageId",
    "id",
    "parentMessageId",
)

processing_files: set[str] = set()
processing_lock = threading.Lock()
shutdown_event = threading.Event()


# ============================================================
# ログ
# ============================================================

def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


# ============================================================
# フォルダチェック
# ============================================================

def validate_directories() -> None:
    if not WATCH_DIR.exists():
        print()
        print(f"[NG] 監視フォルダが存在しません: {WATCH_DIR}")
        print("OneDriveの同期状態とAGENT_REQUEST_DIRを確認してください。")
        sys.exit(1)

    if not WATCH_DIR.is_dir():
        print()
        print(f"[NG] 監視先がフォルダではありません: {WATCH_DIR}")
        sys.exit(1)

    for label, directory in (
        ("完了フォルダ", DONE_DIR),
        ("返信フォルダ", REPLY_DIR),
        ("状態フォルダ", STATE_DIR),
        ("処理済みフォルダ", PROCESSED_DIR),
    ):
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            print()
            print(f"[NG] {label}を作成できません: {directory}")
            print(error)
            sys.exit(1)


# ============================================================
# OneDriveファイル書き込み完了待機
# ============================================================

def wait_until_file_ready(file_path: pathlib.Path) -> bool:
    previous_size = -1

    for attempt in range(1, FILE_READY_RETRIES + 1):
        if not file_path.exists():
            log(
                f"ファイル生成待機 ({attempt}/{FILE_READY_RETRIES}): "
                f"{file_path.name}"
            )
            time.sleep(FILE_READY_INTERVAL_SECONDS)
            continue

        try:
            current_size = file_path.stat().st_size
        except OSError as error:
            log(
                f"ファイル情報取得待機 ({attempt}/{FILE_READY_RETRIES}): "
                f"{error}"
            )
            time.sleep(FILE_READY_INTERVAL_SECONDS)
            continue

        if current_size > 0 and current_size == previous_size:
            try:
                raw_text = file_path.read_text(encoding="utf-8-sig")
                json.loads(raw_text)
                return True
            except (
                json.JSONDecodeError,
                UnicodeDecodeError,
                OSError,
            ) as error:
                log(
                    f"JSON完成待機 ({attempt}/{FILE_READY_RETRIES}): "
                    f"{file_path.name} / {error}"
                )

        previous_size = current_size
        time.sleep(FILE_READY_INTERVAL_SECONDS)

    return False


# ============================================================
# Teams本文処理
# ============================================================

def extract_plain_text(message: object) -> str:
    if isinstance(message, dict):
        message = message.get(
            "content",
            json.dumps(message, ensure_ascii=False),
        )

    if message is None:
        return ""

    message_text = html.unescape(str(message))
    soup = BeautifulSoup(message_text, "html.parser")
    text = soup.get_text(separator="\n", strip=True)

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    return "\n".join(lines)


def extract_message_id(source_data: dict) -> str:
    for key in MESSAGE_ID_KEYS:
        value = str(source_data.get(key, "")).strip()
        if value:
            return value
    return ""


# ============================================================
# 重複Issue防止
# ============================================================

def processed_marker_path(message_id: str) -> pathlib.Path:
    """Message IDに対応する処理済みマーカーのパスを返す。"""
    safe_message_id = "".join(
        character
        for character in message_id
        if character.isalnum() or character in {"-", "_", "."}
    )

    if not safe_message_id:
        safe_message_id = "unknown"

    return PROCESSED_DIR / f"{safe_message_id}.json"


def read_processed_marker(message_id: str) -> dict:
    marker_file = processed_marker_path(message_id)

    if not marker_file.exists():
        return {}

    try:
        return json.loads(marker_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {
            "messageId": message_id,
            "status": "unknown",
        }


def reserve_message(source_data: dict) -> tuple[bool, dict]:
    """
    Message ID単位で処理権を原子的に予約する。

    O_CREAT | O_EXCL により、同時に複数のworkerが動いても
    最初の1件だけが予約に成功する。
    """
    message_id = extract_message_id(source_data)

    if not message_id:
        log("Message IDが無いため重複防止予約を取得できません。")
        return False, {}

    marker_file = processed_marker_path(message_id)
    payload = {
        "messageId": message_id,
        "status": "processing",
        "reservedAt": datetime.now().isoformat(),
        "issueNumber": None,
    }

    flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY

    try:
        file_descriptor = os.open(str(marker_file), flags)
    except FileExistsError:
        return False, read_processed_marker(message_id)
    except OSError as error:
        log(f"処理予約の作成に失敗: {marker_file.name} / {error}")
        return False, {}

    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
    except OSError as error:
        log(f"処理予約の書込に失敗: {marker_file.name} / {error}")
        try:
            marker_file.unlink(missing_ok=True)
        except OSError:
            pass
        return False, {}

    log(f"処理予約を取得: {marker_file.name}")
    return True, payload


def mark_processed(source_data: dict, issue_number: int, issue_url: str) -> None:
    message_id = extract_message_id(source_data)

    if not message_id:
        return

    marker_file = processed_marker_path(message_id)
    payload = {
        "messageId": message_id,
        "status": "completed",
        "issueNumber": issue_number,
        "issueUrl": issue_url,
        "completedAt": datetime.now().isoformat(),
    }

    marker_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    log(f"処理済み記録: {marker_file.name}")


def release_reservation(source_data: dict) -> None:
    """Issue作成失敗時に予約を解除し、再処理可能にする。"""
    message_id = extract_message_id(source_data)

    if not message_id:
        return

    marker_file = processed_marker_path(message_id)

    try:
        marker = read_processed_marker(message_id)
        if marker.get("status") == "processing":
            marker_file.unlink(missing_ok=True)
            log(f"処理予約を解除: {marker_file.name}")
    except OSError as error:
        log(f"処理予約解除失敗: {marker_file.name} / {error}")


def move_duplicate_to_done(
    source_file: pathlib.Path,
    message_id: str,
    marker: dict,
) -> pathlib.Path | None:
    """重複requestをdoneへ退避する。"""
    if not source_file.exists():
        return None

    DONE_DIR.mkdir(parents=True, exist_ok=True)
    issue_number = marker.get("issueNumber") or "pending"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_file = (
        DONE_DIR
        / f"duplicate_{issue_number}_{timestamp}_{source_file.name}"
    )

    try:
        shutil.move(str(source_file), str(target_file))
        log(
            f"重複JSONを退避: {target_file.name} "
            f"(messageId={message_id})"
        )
        return target_file
    except OSError as error:
        log(f"重複JSONの退避失敗: {source_file.name} / {error}")
        return None


# ============================================================
# Issueタイトル・本文
# ============================================================

def create_issue_title(text: str) -> str:
    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    if not lines:
        return "Teams開発依頼"

    title = lines[0]
    trigger_words = {
        "issue",
        "#issue",
        "#agent",
        "agent",
    }

    if title.lower() in trigger_words and len(lines) >= 2:
        title = lines[1]

    return title[:100]


def create_issue_body(data: dict, text: str) -> str:
    sender = str(data.get("sender", "")).strip()
    teams_message_id = extract_message_id(data)
    posted_at = str(data.get("datetime", "")).strip()

    return f"""# Teams開発依頼

## 送信者

{sender or "不明"}

## Teams Message ID

{teams_message_id or "不明"}

## 投稿日時

{posted_at or "不明"}

## 依頼内容

{text or "内容なし"}

## 実装時の注意事項

- 本リポジトリはビルド不要の静的サイトである
- 素の HTML / CSS / JavaScript で実装する
- TypeScript やフレームワークを導入しない
- package.json やビルド設定を新規に追加しない
- 既存のディレクトリ構成とコーディング規約に従う
- コメントは日本語で記述する
- ファイルは UTF-8 (BOMなし) で保存する
- 変更範囲を必要最小限にする
- old/ 配下は変更しない
- Issueに関係しない変更は行わない
- 秘密情報をコードへ直接記載しない

## 完了条件

- [ ] 要件に沿った実装が完了している
- [ ] ブラウザで表示と動作を確認できる
- [ ] 既存ページの表示に影響がない
- [ ] コンソールにエラーが出ていない
- [ ] 変更内容と確認結果がPRに記載されている
"""


# ============================================================
# GitHub Issue作成
# ============================================================

def create_github_issue(title: str, body: str) -> dict | None:
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "teams-github-issue-agent",
    }

    payload = {
        "title": title,
        "body": body,
    }

    for attempt in range(1, HTTP_RETRIES + 1):
        try:
            response = requests.post(
                GITHUB_API_URL,
                headers=headers,
                json=payload,
                timeout=HTTP_TIMEOUT_SECONDS,
            )
        except requests.RequestException as error:
            log(
                f"GitHub API通信エラー ({attempt}/{HTTP_RETRIES}): "
                f"{error}"
            )
            if attempt < HTTP_RETRIES:
                time.sleep(HTTP_RETRY_INTERVAL_SECONDS)
            continue

        log(f"GitHub APIレスポンス: {response.status_code}")

        if response.status_code == 201:
            return response.json()

        if response.status_code in {400, 401, 403, 404, 422}:
            log(f"GitHub Issue作成失敗: {response.text}")
            return None

        log(f"GitHub API一時エラー: {response.text}")

        if attempt < HTTP_RETRIES:
            time.sleep(HTTP_RETRY_INTERVAL_SECONDS)

    return None


# ============================================================
# 状態・返信JSON
# ============================================================

def create_state_json(
    issue_number: int,
    issue_title: str,
    issue_url: str,
    message_id: str,
) -> None:
    state_path = STATE_DIR / f"issue-{issue_number}.json"

    payload = {
        "issueNumber": issue_number,
        "issueTitle": issue_title,
        "issueUrl": issue_url,
        "messageId": message_id,
        "status": "created",
        "createdAt": datetime.now().isoformat(),
    }

    try:
        state_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError as error:
        log(f"状態ファイル出力失敗: {state_path.name} / {error}")
        return

    log(f"状態ファイル出力: {state_path.name}")


def create_reply_json(
    issue_number: int,
    issue_title: str,
    issue_url: str,
    source_data: dict,
) -> None:
    message_id = extract_message_id(source_data)

    if not message_id:
        log(
            f"親メッセージIDが取得できないため返信JSONを出力しません: "
            f"Issue #{issue_number}"
        )
        return

    reply_payload = {
        "type": "issue_created",
        "messageId": message_id,
        "issueNumber": issue_number,
        "issueUrl": issue_url,
        "status": "created",
        "message": (
            f"🤖 Issue #{issue_number} を作成しました\n\n"
            f"タイトル:\n{issue_title}\n\n"
            f"状態:\n実装待ち\n\n"
            f"GitHub:\n{issue_url}"
        ),
    }

    output_path = REPLY_DIR / f"issue-{issue_number}.json"

    try:
        output_path.write_text(
            json.dumps(reply_payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError as error:
        log(f"Teams返信JSON出力失敗: {output_path.name} / {error}")
        return

    log(
        f"Teams返信JSON出力: {output_path.name} "
        f"(messageId={message_id})"
    )

    create_state_json(
        issue_number=issue_number,
        issue_title=issue_title,
        issue_url=issue_url,
        message_id=message_id,
    )


# ============================================================
# request JSON移動
# ============================================================

def move_to_done(
    source_file: pathlib.Path,
    issue_number: int,
) -> pathlib.Path:
    DONE_DIR.mkdir(parents=True, exist_ok=True)
    target_file = DONE_DIR / f"{issue_number}_{source_file.name}"

    if target_file.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_file = (
            DONE_DIR
            / f"{issue_number}_{timestamp}_{source_file.name}"
        )

    shutil.move(str(source_file), str(target_file))
    return target_file


# ============================================================
# JSONファイル1件処理
# ============================================================

def process_json_file(file_path: pathlib.Path) -> None:
    try:
        resolved_path = str(file_path.resolve())
    except OSError:
        resolved_path = str(file_path)

    with processing_lock:
        if resolved_path in processing_files:
            log(f"すでに処理中のためスキップ: {file_path.name}")
            return
        processing_files.add(resolved_path)

    data: dict = {}
    reservation_acquired = False

    try:
        if file_path.suffix.lower() != ".json":
            return

        if not file_path.exists():
            return

        log(f"処理開始: {file_path.name}")

        if not wait_until_file_ready(file_path):
            log(f"JSONが安定しないため処理中止: {file_path.name}")
            return

        try:
            raw_text = file_path.read_text(encoding="utf-8-sig")
            data = json.loads(raw_text)
        except json.JSONDecodeError as error:
            log(f"JSON解析エラー: {file_path.name} / {error}")
            return
        except OSError as error:
            log(f"JSON読込エラー: {file_path.name} / {error}")
            return

        message_id = extract_message_id(data)

        if not message_id:
            log(
                f"Message IDが無いためIssueを作成しません: "
                f"{file_path.name}"
            )
            return

        reservation_acquired, existing_marker = reserve_message(data)

        if not reservation_acquired:
            existing_issue = existing_marker.get("issueNumber")
            existing_status = existing_marker.get("status", "unknown")

            log(
                f"既に処理済みまたは処理中のためスキップ: "
                f"messageId={message_id}, "
                f"status={existing_status}, "
                f"issue={existing_issue}"
            )

            move_duplicate_to_done(
                source_file=file_path,
                message_id=message_id,
                marker=existing_marker,
            )
            return

        message = data.get("message", "")
        text = extract_plain_text(message)
        title = create_issue_title(text)
        body = create_issue_body(data, text)

        log(f"Issueタイトル: {title}")

        issue = create_github_issue(title, body)

        if issue is None:
            log(
                f"Issue作成に失敗したためJSONを残します: "
                f"{file_path.name}"
            )
            release_reservation(data)
            reservation_acquired = False
            return

        issue_number = issue["number"]
        issue_url = issue.get("html_url", "")

        # Issue作成直後にマーカーをcompletedへ更新する。
        # 後続の返信JSON生成が失敗してもIssue重複を防止できる。
        mark_processed(
            source_data=data,
            issue_number=issue_number,
            issue_url=issue_url,
        )
        reservation_acquired = False

        log(f"Issue作成成功: #{issue_number}")

        if issue_url:
            log(f"Issue URL: {issue_url}")

        create_reply_json(
            issue_number=issue_number,
            issue_title=title,
            issue_url=issue_url,
            source_data=data,
        )

        target_file = move_to_done(file_path, issue_number)
        log(f"処理済みJSON移動完了: {target_file}")

    except Exception as error:
        log(
            f"予期しないエラー: {file_path.name} / "
            f"{type(error).__name__}: {error}"
        )

        # Issue番号が確定する前の例外なら予約を解除する。
        if reservation_acquired and data:
            release_reservation(data)

    finally:
        with processing_lock:
            processing_files.discard(resolved_path)


# ============================================================
# Watchdogイベント処理
# ============================================================

class AgentRequestHandler(FileSystemEventHandler):
    def queue_file(self, file_name: str) -> None:
        file_path = pathlib.Path(file_name)

        if file_path.suffix.lower() != ".json":
            return

        worker = threading.Thread(
            target=process_json_file,
            args=(file_path,),
            daemon=True,
        )
        worker.start()

    def on_created(self, event) -> None:
        if event.is_directory:
            return

        log(f"新規ファイル検知: {event.src_path}")
        self.queue_file(event.src_path)

    def on_moved(self, event) -> None:
        # OneDriveでは作成後に移動イベントが続く場合がある。
        # on_createdのみを処理し、二重起動を防ぐ。
        return


# ============================================================
# 起動時に既存JSONを処理
# ============================================================

def process_existing_files() -> None:
    existing_files = sorted(WATCH_DIR.glob("*.json"))

    if not existing_files:
        log("起動時の未処理JSONはありません。")
        return

    log(f"起動時の未処理JSON: {len(existing_files)}件")

    for file_path in existing_files:
        process_json_file(file_path)


# ============================================================
# 終了処理
# ============================================================

def handle_shutdown(signum, frame) -> None:
    log("終了要求を受け付けました。")
    shutdown_event.set()


# ============================================================
# メイン処理
# ============================================================

def main() -> None:
    validate_directories()

    log("=" * 60)
    log("Teams GitHub Issue Agent 起動")
    log(f"監視フォルダ: {WATCH_DIR}")
    log(f"完了フォルダ: {DONE_DIR}")
    log(f"返信フォルダ: {REPLY_DIR}")
    log(f"状態フォルダ: {STATE_DIR}")
    log(f"処理済みフォルダ: {PROCESSED_DIR}")
    log(f"GitHubリポジトリ: {GITHUB_OWNER}/{GITHUB_REPO}")
    log("=" * 60)

    signal.signal(signal.SIGINT, handle_shutdown)

    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, handle_shutdown)

    process_existing_files()

    event_handler = AgentRequestHandler()
    observer = Observer()
    observer.schedule(
        event_handler,
        str(WATCH_DIR),
        recursive=False,
    )
    observer.start()

    log("監視を開始しました。")
    log("終了するには Ctrl+C を押してください。")

    try:
        while not shutdown_event.is_set():
            time.sleep(1)
    finally:
        log("監視を終了しています。")
        observer.stop()
        observer.join()
        log("正常終了しました。")


if __name__ == "__main__":
    main()
