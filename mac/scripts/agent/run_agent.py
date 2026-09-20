"""
Phase2: GitHub IssueをVS Code + Clineへ投入し、実装完了を検知して
承認待ちJSONを出力する。

安定化方針:
- Cline入力欄は固定座標 X=1500 / Y=900 を直接クリックする。
- Clineの新規チャットボタンは自動操作しない。
  新規チャット操作でClineペインが閉じる・消える問題を回避する。
- VS Codeのコマンドパレットは使用しない。
- 日本語IMEの影響を避けるため、文字入力はpyperclip + Ctrl+Vで行う。
"""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime


REQUIRED_ENV_VARS = [
    "GITHUB_OWNER",
    "GITHUB_REPO",
    "AGENT_REPLY_DIR",
    "AGENT_WAITING_DIR",
    "AGENT_STATE_DIR",
    "TARGET_REPO_PATH",
]

GITHUB_OWNER = ""
GITHUB_REPO = ""
GITHUB_REPOSITORY = ""
REPLY_DIR = pathlib.Path(".")
WAITING_DIR = pathlib.Path(".")
STATE_DIR = pathlib.Path(".")
TARGET_REPO = pathlib.Path(".")

VSCODE_BOOT_WAIT_SECONDS = 25
CLINE_FOCUS_WAIT_SECONDS = 2
PASTE_WAIT_SECONDS = 3
IMPLEMENTATION_TIMEOUT_SECONDS = 900
POLL_INTERVAL_SECONDS = 10
STABLE_DURATION_SECONDS = 60

# 実測したCline入力欄の固定座標
CLINE_INPUT_X = 1500
CLINE_INPUT_Y = 900


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def fail(message: str, code: int = 1) -> None:
    log(f"[ERROR] {message}")
    raise SystemExit(code)


def validate_environment_variables() -> None:
    global GITHUB_OWNER
    global GITHUB_REPO
    global GITHUB_REPOSITORY
    global REPLY_DIR
    global WAITING_DIR
    global STATE_DIR
    global TARGET_REPO

    missing = []
    print()
    print("=" * 60)
    print("環境変数チェック")
    print("=" * 60)

    for name in REQUIRED_ENV_VARS:
        value = os.getenv(name)
        if value and value.strip():
            print(f"[OK] {name}: {value}")
        else:
            missing.append(name)
            print(f"[NG] {name}: 未設定")

    if missing:
        print()
        print("不足している環境変数:")
        for name in missing:
            print(f"- {name}")
        raise SystemExit(1)

    GITHUB_OWNER = os.environ["GITHUB_OWNER"]
    GITHUB_REPO = os.environ["GITHUB_REPO"]
    GITHUB_REPOSITORY = f"{GITHUB_OWNER}/{GITHUB_REPO}"
    REPLY_DIR = pathlib.Path(os.environ["AGENT_REPLY_DIR"])
    WAITING_DIR = pathlib.Path(os.environ["AGENT_WAITING_DIR"])
    STATE_DIR = pathlib.Path(os.environ["AGENT_STATE_DIR"])
    TARGET_REPO = pathlib.Path(os.environ["TARGET_REPO_PATH"])

    print("=" * 60)
    print("必要な環境変数はすべて設定されています。")
    print("=" * 60)
    print()


def run_command(
    command: list[str],
    cwd: pathlib.Path | None = None,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    log(f"> {subprocess.list2cmdline(command)}")
    result = subprocess.run(
        command,
        cwd=str(cwd) if cwd else None,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    )

    if result.stdout and result.stdout.strip():
        print(result.stdout.rstrip())
    if result.stderr and result.stderr.strip():
        print(result.stderr.rstrip(), file=sys.stderr)

    if check and result.returncode != 0:
        fail(
            f"コマンド失敗 ({result.returncode}): "
            f"{subprocess.list2cmdline(command)}"
        )
    return result


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run_command(["git", *args], cwd=TARGET_REPO, check=check)


def require_command(name: str) -> None:
    if not shutil.which(name):
        fail(f"'{name}' がPATHにありません。")


def validate_directories() -> None:
    if not (TARGET_REPO / ".git").exists():
        fail(f"Gitリポジトリではありません: {TARGET_REPO}")

    for directory in (REPLY_DIR, WAITING_DIR, STATE_DIR):
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            fail(f"フォルダを作成できません: {directory} / {error}")


def state_path(issue_number: int) -> pathlib.Path:
    return STATE_DIR / f"issue-{issue_number}.json"


def load_state(issue_number: int) -> dict:
    path = state_path(issue_number)
    if not path.exists():
        log(f"状態ファイルが見つかりません: {path.name}")
        return {}

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as error:
        log(f"状態ファイル読込失敗: {error}")
        return {}


def load_message_id(issue_number: int) -> str:
    return str(load_state(issue_number).get("messageId", "")).strip()


def update_state(issue_number: int, updates: dict) -> None:
    state = load_state(issue_number) or {"issueNumber": issue_number}
    state.update(updates)
    state["updatedAt"] = datetime.now().isoformat()

    path = state_path(issue_number)
    path.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    log(f"状態ファイル更新: {path.name}")


def load_issue(issue_number: int) -> dict:
    result = run_command(
        [
            "gh",
            "issue",
            "view",
            str(issue_number),
            "--repo",
            GITHUB_REPOSITORY,
            "--json",
            "number,title,body,url,labels,state",
        ],
        cwd=TARGET_REPO,
    )

    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as error:
        fail(f"Issue JSONを解析できません: {error}")
    return {}


def slugify(text: str, max_length: int = 40) -> str:
    value = re.sub(r"[^a-z0-9]+", "-", text.lower().strip()).strip("-")
    return value[:max_length].rstrip("-") or "task"


def prepare_branch(
    issue_number: int,
    issue_title: str,
    base_branch: str,
) -> str:
    branch = f"issue-{issue_number}-{slugify(issue_title)}"
    dirty = git("status", "--porcelain").stdout.strip()
    current = git("branch", "--show-current").stdout.strip()

    if dirty and current == branch:
        log(f"同一Issueブランチの変更を継続します: {branch}")
        return branch

    if dirty:
        fail(
            f"ブランチ {current} に未コミット変更があります。"
            "commitまたはstashしてから実行してください。"
        )

    if current == branch:
        log(f"既に対象ブランチです: {branch}")
        return branch

    git("fetch", "origin")
    git("checkout", base_branch)
    git("pull", "--ff-only", "origin", base_branch)

    exists = git(
        "show-ref",
        "--verify",
        "--quiet",
        f"refs/heads/{branch}",
        check=False,
    ).returncode == 0

    if exists:
        log(f"既存ブランチへ切り替えます: {branch}")
        git("checkout", branch)
    else:
        git("checkout", "-b", branch)

    return branch


def build_prompt(issue: dict, branch: str) -> str:
    labels = ", ".join(
        item.get("name", "") for item in issue.get("labels", [])
    ) or "なし"
    body = issue.get("body") or "(本文なし)"

    return f"""GitHub Issue #{issue['number']} を実装してください。

タイトル: {issue['title']}
URL: {issue['url']}
ラベル: {labels}

本文:
{body}

前提:
- 作業ブランチ {branch} は作成済みで、既にチェックアウトされています。
- 本リポジトリはビルド不要の静的サイトです。
- ターミナルはPowerShellです。

作業ルール:
1. 素のHTML / CSS / JavaScriptで実装してください。
2. TypeScriptやフレームワークを追加しないでください。
3. package.jsonやビルド設定を追加しないでください。
4. 新規ページはリポジトリ直下のディレクトリにindex.htmlを配置してください。
5. コメントは日本語で記述してください。
6. ファイルはUTF-8 BOMなしで保存してください。
7. old/配下とIssueに関係しないファイルは変更しないでください。
8. git操作、commit、push、PR作成は行わないでください。
9. 判断に迷う場合は勝手に決めず、質問してください。
10. 最後に変更ファイルと実装内容を簡潔に報告してください。
"""


def write_json(directory: pathlib.Path, file_name: str, payload: dict) -> bool:
    path = directory / file_name
    try:
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        log(f"JSON出力: {path}")
        return True
    except OSError as error:
        log(f"JSON出力失敗: {path} / {error}")
        return False


def notify_teams(
    message_type: str,
    issue_number: int,
    message: str,
    extra: dict | None = None,
) -> None:
    message_id = load_message_id(issue_number)
    if not message_id:
        log("messageIdが無いためTeams通知をスキップします。")
        return

    payload = {
        "type": message_type,
        "messageId": message_id,
        "issueNumber": issue_number,
        "status": message_type,
        "message": message,
    }
    if extra:
        payload.update(extra)

    write_json(
        REPLY_DIR,
        f"issue-{issue_number}-{message_type}-{datetime.now():%H%M%S}.json",
        payload,
    )


def copy_to_clipboard(text: str) -> bool:
    try:
        import pyperclip

        pyperclip.copy(text)
        log("クリップボードへコピーしました。")
        return True
    except Exception as error:
        log(f"クリップボードコピー失敗: {error}")
        return False


def launch_vscode() -> None:
    require_command("code")
    subprocess.Popen(["code", str(TARGET_REPO)], shell=True)
    log(f"VS Code起動後、{VSCODE_BOOT_WAIT_SECONDS}秒待機します。")
    time.sleep(VSCODE_BOOT_WAIT_SECONDS)


def focus_cline_input(pyautogui) -> None:
    """
    実測した固定座標でCline入力欄をクリックする。
    新規チャットボタンは操作しないため、Clineペインが消える問題を回避する。
    """
    log(f"Cline入力欄をクリック: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    pyautogui.click(CLINE_INPUT_X, CLINE_INPUT_Y)
    time.sleep(CLINE_FOCUS_WAIT_SECONDS)

    # フォーカスを安定させるため同じ位置をもう一度クリックする。
    pyautogui.click(CLINE_INPUT_X, CLINE_INPUT_Y)
    time.sleep(0.5)


def inject_prompt_gui(prompt: str) -> bool:
    try:
        import pyautogui
        import pyperclip
    except ImportError:
        log("pip install pyautogui pyperclip を実行してください。")
        return False

    pyautogui.FAILSAFE = True

    try:
        focus_cline_input(pyautogui)

        # 入力欄に以前の未送信文字が残っている場合だけ消去する。
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)
        pyautogui.press("backspace")
        time.sleep(0.3)

        # コマンド名などでクリップボードが上書きされないよう、直前に再コピーする。
        pyperclip.copy(prompt)
        pyautogui.hotkey("ctrl", "v")
        log("Cline入力欄へプロンプトを貼り付けました。")

        time.sleep(PASTE_WAIT_SECONDS)
        pyautogui.press("enter")
        log("Clineへプロンプトを送信しました。")
        return True

    except Exception as error:
        log(f"Cline自動投入に失敗しました: {error}")
        return False


def test_cline_click() -> bool:
    """Issue処理をせず、Cline入力欄のクリックだけ確認する。"""
    try:
        import pyautogui
    except ImportError:
        fail("pyautoguiが未インストールです。")

    pyautogui.FAILSAFE = True
    print()
    print("3秒後にCline入力欄をクリックします。")
    print(f"座標: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    print("VS Codeを前面に表示してください。")
    time.sleep(3)
    focus_cline_input(pyautogui)
    print("クリックしました。Cline入力欄にカーソルがあるか確認してください。")
    return True


def inject_prompt_manual() -> None:
    print()
    print("=" * 60)
    print("Cline入力欄を手動で選択し、Ctrl+V、Enterを実行してください。")
    print("=" * 60)
    input("投入したらEnterを押してください...")


def snapshot_changes() -> str:
    return git("status", "--porcelain", check=False).stdout.strip()


def wait_for_implementation() -> str:
    log("実装完了を待機します。")
    started_at = time.time()
    previous = snapshot_changes()
    stable_since: float | None = None

    while True:
        if time.time() - started_at > IMPLEMENTATION_TIMEOUT_SECONDS:
            log("実装待機がタイムアウトしました。")
            break

        time.sleep(POLL_INTERVAL_SECONDS)
        current = snapshot_changes()

        if current != previous:
            log("変更を検知しました。")
            previous = current
            stable_since = None
            continue

        if not current:
            continue

        if stable_since is None:
            stable_since = time.time()
            log(
                f"変更停止。{STABLE_DURATION_SECONDS}秒後に完了判定します。"
            )
            continue

        if time.time() - stable_since >= STABLE_DURATION_SECONDS:
            log("実装完了を検知しました。")
            break

    return previous


def list_changed_files() -> list[str]:
    files = []
    for line in snapshot_changes().splitlines():
        path = line[3:].strip().strip('"')
        if path:
            files.append(path)
    return files


def expand_changed_files(changed_files: list[str]) -> list[str]:
    expanded: set[str] = set()
    for relative_path in changed_files:
        target = TARGET_REPO / relative_path.rstrip("/")
        if target.is_dir():
            for child in target.rglob("*"):
                if child.is_file():
                    expanded.add(child.relative_to(TARGET_REPO).as_posix())
        else:
            expanded.add(relative_path.replace("\\", "/"))
    return sorted(expanded)


def write_waiting_json(
    issue: dict,
    branch: str,
    changed_files: list[str],
) -> None:
    issue_number = issue["number"]
    expanded_files = expand_changed_files(changed_files)
    changed_files_text = "\n".join(
        f"・{path}" for path in expanded_files
    ) or "・変更なし"

    summary_text = "\n".join(
        f"・ファイルを追加または更新: {path}"
        for path in expanded_files
    ) or "・変更内容を確認してください。"

    message = (
        f"⏳ Issue #{issue_number} の実装が完了しました\n\n"
        f"タイトル:\n{issue['title']}\n\n"
        f"ブランチ:\n{branch}\n\n"
        f"変更ファイル:\n{changed_files_text}\n\n"
        f"実装内容:\n{summary_text}\n\n"
        "内容を確認して承認してください。"
    )

    payload = {
        "type": "waiting_approval",
        "messageId": load_message_id(issue_number),
        "issueNumber": issue_number,
        "issueTitle": issue["title"],
        "issueUrl": issue["url"],
        "branch": branch,
        "changedFiles": expanded_files,
        "changedFilesText": changed_files_text,
        "summaryText": summary_text,
        "warningText": "",
        "diffStat": git("diff", "--stat", check=False).stdout.strip(),
        "status": "waiting_approval",
        "message": message,
    }

    write_json(WAITING_DIR, f"issue-{issue_number}.json", payload)
    update_state(
        issue_number,
        {
            "status": "waiting_approval",
            "branch": branch,
            "changedFiles": expanded_files,
        },
    )
    notify_teams(
        "waiting_approval",
        issue_number,
        message,
        {
            "branch": branch,
            "changedFiles": expanded_files,
            "summaryText": summary_text,
        },
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="GitHub IssueをClineへ投入し、承認待ちJSONを出力します。"
    )
    parser.add_argument("issue", type=int, nargs="?", help="GitHub Issue番号")
    parser.add_argument("--base", default=os.getenv("BASE_BRANCH", "main"))
    parser.add_argument("--manual", action="store_true")
    parser.add_argument("--no-vscode", action="store_true")
    parser.add_argument("--message-id", default="")
    parser.add_argument(
        "--test-click",
        action="store_true",
        help="Issue処理をせずCline入力欄のクリックだけテストする",
    )
    args = parser.parse_args()

    if args.test_click:
        test_cline_click()
        return

    if args.issue is None:
        parser.error("Issue番号を指定してください。")

    validate_environment_variables()
    validate_directories()

    for command in ("git", "gh"):
        require_command(command)

    run_command(["gh", "auth", "status"], cwd=TARGET_REPO)

    log("=" * 60)
    log(f"Cline 実装エージェント: Issue #{args.issue}")
    log(f"Cline入力座標: X={CLINE_INPUT_X}, Y={CLINE_INPUT_Y}")
    log("新規チャット操作: 無効")
    log("=" * 60)

    issue = load_issue(args.issue)
    if issue.get("state") != "OPEN":
        fail(f"Issue #{args.issue} はOPENではありません。")

    if args.message_id:
        update_state(
            args.issue,
            {
                "issueTitle": issue["title"],
                "issueUrl": issue["url"],
                "messageId": args.message_id.strip(),
            },
        )

    branch = prepare_branch(args.issue, issue["title"], args.base)

    update_state(
        args.issue,
        {"status": "implementation_started", "branch": branch},
    )
    notify_teams(
        "implementation_started",
        args.issue,
        (
            f"🚀 Issue #{args.issue} の実装を開始しました\n\n"
            f"タイトル:\n{issue['title']}\n\n"
            f"ブランチ:\n{branch}"
        ),
        {"branch": branch},
    )

    prompt = build_prompt(issue, branch)
    copy_to_clipboard(prompt)

    if not args.no_vscode:
        launch_vscode()

    injected = False
    if not args.manual:
        injected = inject_prompt_gui(prompt)

    if not injected:
        inject_prompt_manual()

    wait_for_implementation()
    changed_files = list_changed_files()

    if not changed_files:
        log("変更が検出されませんでした。")
        update_state(args.issue, {"status": "implementation_failed"})
        notify_teams(
            "implementation_failed",
            args.issue,
            f"⚠️ Issue #{args.issue} の実装で変更が検出されませんでした。",
            {"branch": branch},
        )
        return

    write_waiting_json(issue, branch, changed_files)
    log("Phase2 完了。承認待ちへ移行します。")


if __name__ == "__main__":
    main()
