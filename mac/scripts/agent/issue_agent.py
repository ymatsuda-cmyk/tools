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
]

ENV_SETUP_EXAMPLES = {
    "GITHUB_TOKEN": (
        'setx GITHUB_TOKEN "github_pat_xxxxxxxxxxxxx"'
    ),
    "GITHUB_OWNER": (
        'setx GITHUB_OWNER "ymatsuda-cmyk"'
    ),
    "GITHUB_REPO": (
        'setx GITHUB_REPO "tools"'
    ),
    "AGENT_REQUEST_DIR": (
        'setx AGENT_REQUEST_DIR '
        '"C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\request"'
    ),
    "AGENT_DONE_DIR": (
        'setx AGENT_DONE_DIR '
        '"C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\done"'
    ),
    "AGENT_REPLY_DIR": (
        'setx AGENT_REPLY_DIR '
        '"C:\\Users\\matsuda\\OneDrive - '
        '株式会社日本ビジネスアシスト\\work\\agent\\reply"'
    ),
}


def mask_secret(secret: str) -> str:
    """
    トークンをログへそのまま出さず、一部だけ表示する。
    """
    if len(secret) <= 8:
        return "********"

    return (
        f"{secret[:4]}"
        f"{'*' * 8}"
        f"{secret[-4:]}"
    )


def validate_environment_variables() -> None:
    """
    必須環境変数を起動時に確認する。
    不足している場合は設定コマンドを表示して終了する。
    """
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
            masked_value = mask_secret(value)
            print(
                f"[OK] {variable_name}: "
                f"設定済み ({masked_value})"
            )
        else:
            print(
                f"[OK] {variable_name}: {value}"
            )

    if not missing_variables:
        print("=" * 60)
        print("必要な環境変数はすべて設定されています。")
        print("=" * 60)
        print()
        return

    print()
    print("=" * 60)
    print("不足している環境変数があります。")
    print("=" * 60)

    for variable_name in missing_variables:
        print(f"- {variable_name}")

    print()
    print("PowerShellで以下を実行してください。")
    print()

    for variable_name in missing_variables:
        print(ENV_SETUP_EXAMPLES[variable_name])

    print()
    print(
        "setxで設定した値は、現在開いているPowerShellには"
        "反映されません。"
    )
    print(
        "設定後にPowerShellを閉じて、新しいPowerShellを"
        "開いてください。"
    )
    print()

    sys.exit(1)


# ============================================================
# 環境変数チェック後に設定値を読み込む
# ============================================================

validate_environment_variables()

GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_OWNER = os.environ["GITHUB_OWNER"]
GITHUB_REPO = os.environ["GITHUB_REPO"]

WATCH_DIR = pathlib.Path(
    os.environ["AGENT_REQUEST_DIR"]
)

DONE_DIR = pathlib.Path(
    os.environ["AGENT_DONE_DIR"]
)

REPLY_DIR = pathlib.Path(
    os.environ["AGENT_REPLY_DIR"]
)

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

# Teamsの親メッセージIDとして使用する候補キー
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
    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        f"[{timestamp}] {message}",
        flush=True,
    )


# ============================================================
# フォルダチェック
# ============================================================

def validate_directories() -> None:
    """
    監視フォルダ、完了フォルダ、返信フォルダをチェックする。
    """
    if not WATCH_DIR.exists():
        print()
        print(
            f"[NG] 監視フォルダが存在しません: "
            f"{WATCH_DIR}"
        )
        print(
            "OneDriveの同期状態とAGENT_REQUEST_DIRを"
            "確認してください。"
        )
        sys.exit(1)

    if not WATCH_DIR.is_dir():
        print()
        print(
            f"[NG] 監視先がフォルダではありません: "
            f"{WATCH_DIR}"
        )
        sys.exit(1)

    for label, directory in (
        ("完了フォルダ", DONE_DIR),
        ("返信フォルダ", REPLY_DIR),
    ):
        try:
            directory.mkdir(
                parents=True,
                exist_ok=True,
            )
        except OSError as error:
            print()
            print(
                f"[NG] {label}を作成できません: "
                f"{directory}"
            )
            print(error)
            sys.exit(1)


# ============================================================
# OneDriveファイル書き込み完了待機
# ============================================================

def wait_until_file_ready(
    file_path: pathlib.Path,
) -> bool:
    """
    OneDrive同期中またはPower Automate書き込み途中の
    JSONを読み込まないようにする。
    """
    previous_size = -1

    for attempt in range(
        1,
        FILE_READY_RETRIES + 1,
    ):
        if not file_path.exists():
            log(
                f"ファイル生成待機 "
                f"({attempt}/{FILE_READY_RETRIES}): "
                f"{file_path.name}"
            )

            time.sleep(
                FILE_READY_INTERVAL_SECONDS
            )
            continue

        try:
            current_size = file_path.stat().st_size
        except OSError as error:
            log(
                f"ファイル情報取得待機 "
                f"({attempt}/{FILE_READY_RETRIES}): "
                f"{error}"
            )

            time.sleep(
                FILE_READY_INTERVAL_SECONDS
            )
            continue

        if (
            current_size > 0
            and current_size == previous_size
        ):
            try:
                raw_text = file_path.read_text(
                    encoding="utf-8-sig"
                )

                json.loads(raw_text)

                return True

            except (
                json.JSONDecodeError,
                UnicodeDecodeError,
                OSError,
            ) as error:
                log(
                    f"JSON完成待機 "
                    f"({attempt}/{FILE_READY_RETRIES}): "
                    f"{file_path.name} / {error}"
                )

        previous_size = current_size

        time.sleep(
            FILE_READY_INTERVAL_SECONDS
        )

    return False


# ============================================================
# Teams HTMLからテキストを抽出
# ============================================================

def extract_plain_text(
    message: object,
) -> str:
    """
    TeamsのHTML本文をプレーンテキストへ変換する。
    """
    if isinstance(message, dict):
        message = message.get(
            "content",
            json.dumps(
                message,
                ensure_ascii=False,
            ),
        )

    if message is None:
        return ""

    message_text = html.unescape(
        str(message)
    )

    soup = BeautifulSoup(
        message_text,
        "html.parser",
    )

    text = soup.get_text(
        separator="\n",
        strip=True,
    )

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    return "\n".join(lines)


# ============================================================
# Teams親メッセージID抽出
# ============================================================

def extract_message_id(
    source_data: dict,
) -> str:
    """
    Power Automateのキー名ゆれを吸収して
    Teamsの親メッセージIDを取り出す。
    """
    for key in MESSAGE_ID_KEYS:
        value = str(
            source_data.get(key, "")
        ).strip()

        if value:
            return value

    return ""


# ============================================================
# Issueタイトル生成
# ============================================================

def create_issue_title(
    text: str,
) -> str:
    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    if not lines:
        return "Teams開発依頼"

    title = lines[0]

    # Teamsタグが先頭に入っている場合は次の行を使う
    trigger_words = {
        "issue",
        "#issue",
        "#agent",
        "agent",
    }

    if (
        title.lower() in trigger_words
        and len(lines) >= 2
    ):
        title = lines[1]

    return title[:100]


# ============================================================
# Issue本文生成
# ============================================================

def create_issue_body(
    data: dict,
    text: str,
) -> str:
    sender = str(
        data.get("sender", "")
    ).strip()

    teams_message_id = extract_message_id(data)

    posted_at = str(
        data.get("datetime", "")
    ).strip()

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

- 既存コードと既存の設計方針を確認する
- 変更範囲を必要最小限にする
- Issueに関係しない変更は行わない
- 必要なテストを追加する
- 秘密情報をコードへ直接記載しない

## 完了条件

- [ ] 要件に沿った実装が完了している
- [ ] 必要なテストが追加されている
- [ ] 既存テストが成功している
- [ ] lintが成功している
- [ ] buildが成功している
- [ ] 変更内容とテスト結果がPRに記載されている
"""


# ============================================================
# GitHub Issue作成
# ============================================================

def create_github_issue(
    title: str,
    body: str,
) -> dict | None:
    headers = {
        "Authorization": (
            f"Bearer {GITHUB_TOKEN}"
        ),
        "Accept": (
            "application/vnd.github+json"
        ),
        "X-GitHub-Api-Version": (
            "2022-11-28"
        ),
        "User-Agent": (
            "teams-github-issue-agent"
        ),
    }

    payload = {
        "title": title,
        "body": body,
    }

    for attempt in range(
        1,
        HTTP_RETRIES + 1,
    ):
        try:
            response = requests.post(
                GITHUB_API_URL,
                headers=headers,
                json=payload,
                timeout=HTTP_TIMEOUT_SECONDS,
            )

        except requests.RequestException as error:
            log(
                f"GitHub API通信エラー "
                f"({attempt}/{HTTP_RETRIES}): "
                f"{error}"
            )

            if attempt < HTTP_RETRIES:
                time.sleep(
                    HTTP_RETRY_INTERVAL_SECONDS
                )

            continue

        log(
            f"GitHub APIレスポンス: "
            f"{response.status_code}"
        )

        if response.status_code == 201:
            return response.json()

        # 入力、認証、権限、リポジトリ指定の問題は
        # 再試行しても改善しないため終了する
        if response.status_code in {
            400,
            401,
            403,
            404,
            422,
        }:
            log(
                "GitHub Issue作成失敗: "
                f"{response.text}"
            )
            return None

        log(
            "GitHub API一時エラー: "
            f"{response.text}"
        )

        if attempt < HTTP_RETRIES:
            time.sleep(
                HTTP_RETRY_INTERVAL_SECONDS
            )

    return None


# ============================================================
# Teams返信JSON作成
# ============================================================

def create_reply_json(
    issue_number: int,
    issue_title: str,
    issue_url: str,
    source_data: dict,
) -> None:
    """
    Teamsへスレッド返信するためのJSONを出力する。
    親メッセージIDはTeamsトリガーのidを使用する。
    """
    message_id = extract_message_id(source_data)

    if not message_id:
        log(
            f"親メッセージIDが取得できないため"
            f"返信JSONを出力しません: "
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
            f"タイトル:\n"
            f"{issue_title}\n\n"
            f"状態:\n"
            f"実装待ち\n\n"
            f"GitHub:\n"
            f"{issue_url}"
        ),
    }

    output_path = (
        REPLY_DIR
        / f"issue-{issue_number}.json"
    )

    try:
        output_path.write_text(
            json.dumps(
                reply_payload,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    except OSError as error:
        log(
            f"Teams返信JSON出力失敗: "
            f"{output_path.name} / {error}"
        )
        return

    log(
        f"Teams返信JSON出力: "
        f"{output_path.name} "
        f"(messageId={message_id})"
    )


# ============================================================
# 処理済みJSON移動
# ============================================================

def move_to_done(
    source_file: pathlib.Path,
    issue_number: int,
) -> pathlib.Path:
    DONE_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    target_file = (
        DONE_DIR
        / f"{issue_number}_{source_file.name}"
    )

    # 同じファイル名がある場合は日時を追加
    if target_file.exists():
        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        target_file = (
            DONE_DIR
            / (
                f"{issue_number}_"
                f"{timestamp}_"
                f"{source_file.name}"
            )
        )

    shutil.move(
        str(source_file),
        str(target_file),
    )

    return target_file


# ============================================================
# JSONファイル1件処理
# ============================================================

def process_json_file(
    file_path: pathlib.Path,
) -> None:
    try:
        resolved_path = str(
            file_path.resolve()
        )
    except OSError:
        resolved_path = str(file_path)

    with processing_lock:
        if resolved_path in processing_files:
            log(
                f"すでに処理中のためスキップ: "
                f"{file_path.name}"
            )
            return

        processing_files.add(
            resolved_path
        )

    try:
        if file_path.suffix.lower() != ".json":
            return

        if not file_path.exists():
            return

        log(
            f"処理開始: {file_path.name}"
        )

        if not wait_until_file_ready(
            file_path
        ):
            log(
                f"JSONが安定しないため処理中止: "
                f"{file_path.name}"
            )
            return

        try:
            raw_text = file_path.read_text(
                encoding="utf-8-sig"
            )

            data = json.loads(
                raw_text
            )

        except json.JSONDecodeError as error:
            log(
                f"JSON解析エラー: "
                f"{file_path.name} / {error}"
            )
            return

        except OSError as error:
            log(
                f"JSON読込エラー: "
                f"{file_path.name} / {error}"
            )
            return

        message = data.get(
            "message",
            "",
        )

        text = extract_plain_text(
            message
        )

        title = create_issue_title(
            text
        )

        body = create_issue_body(
            data,
            text,
        )

        log(
            f"Issueタイトル: {title}"
        )

        issue = create_github_issue(
            title,
            body,
        )

        if issue is None:
            log(
                f"Issue作成に失敗したため"
                f"JSONを残します: "
                f"{file_path.name}"
            )
            return

        issue_number = issue["number"]

        issue_url = issue.get(
            "html_url",
            "",
        )

        log(
            f"Issue作成成功: "
            f"#{issue_number}"
        )

        if issue_url:
            log(
                f"Issue URL: {issue_url}"
            )

        create_reply_json(
            issue_number=issue_number,
            issue_title=title,
            issue_url=issue_url,
            source_data=data,
        )

        target_file = move_to_done(
            file_path,
            issue_number,
        )

        log(
            f"処理済みJSON移動完了: "
            f"{target_file}"
        )

    except Exception as error:
        log(
            f"予期しないエラー: "
            f"{file_path.name} / "
            f"{type(error).__name__}: "
            f"{error}"
        )

    finally:
        with processing_lock:
            processing_files.discard(
                resolved_path
            )


# ============================================================
# Watchdogイベント処理
# ============================================================

class AgentRequestHandler(
    FileSystemEventHandler
):
    def queue_file(
        self,
        file_name: str,
    ) -> None:
        file_path = pathlib.Path(
            file_name
        )

        if file_path.suffix.lower() != ".json":
            return

        worker = threading.Thread(
            target=process_json_file,
            args=(file_path,),
            daemon=True,
        )

        worker.start()

    def on_created(
        self,
        event,
    ) -> None:
        if event.is_directory:
            return

        log(
            f"新規ファイル検知: "
            f"{event.src_path}"
        )

        self.queue_file(
            event.src_path
        )

    def on_moved(
        self,
        event,
    ) -> None:
        if event.is_directory:
            return

        log(
            f"移動ファイル検知: "
            f"{event.dest_path}"
        )

        self.queue_file(
            event.dest_path
        )


# ============================================================
# 起動時に既存JSONを処理
# ============================================================

def process_existing_files() -> None:
    existing_files = sorted(
        WATCH_DIR.glob("*.json")
    )

    if not existing_files:
        log(
            "起動時の未処理JSONはありません。"
        )
        return

    log(
        f"起動時の未処理JSON: "
        f"{len(existing_files)}件"
    )

    for file_path in existing_files:
        process_json_file(
            file_path
        )


# ============================================================
# 終了処理
# ============================================================

def handle_shutdown(
    signum,
    frame,
) -> None:
    log(
        "終了要求を受け付けました。"
    )

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
    log(
        f"GitHubリポジトリ: "
        f"{GITHUB_OWNER}/{GITHUB_REPO}"
    )
    log("=" * 60)

    signal.signal(
        signal.SIGINT,
        handle_shutdown,
    )

    if hasattr(signal, "SIGTERM"):
        signal.signal(
            signal.SIGTERM,
            handle_shutdown,
        )

    # 起動前から存在する未処理JSONを処理する
    process_existing_files()

    event_handler = AgentRequestHandler()

    observer = Observer()

    observer.schedule(
        event_handler,
        str(WATCH_DIR),
        recursive=False,
    )

    observer.start()

    log(
        "監視を開始しました。"
    )
    log(
        "終了するには Ctrl+C を押してください。"
    )

    try:
        while not shutdown_event.is_set():
            time.sleep(1)

    finally:
        log(
            "監視を終了しています。"
        )

        observer.stop()
        observer.join()

        log(
            "正常終了しました。"
        )


if __name__ == "__main__":
    main()
