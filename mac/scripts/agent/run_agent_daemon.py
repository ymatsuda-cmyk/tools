"""
GitHub Issue実装エージェントの自動起動デーモン。

AGENT_STATE_DIR の issue-*.json を監視し、status == "created" のIssueを
1件ずつ run_agent.py へ渡す。

安全設計:
- GUI操作の競合を避けるため逐次実行する
- デーモンの二重起動をロックファイルで防ぐ
- Issue単位の起動予約をstateへ記録する
- run_agent.pyが異常終了した場合はdaemon_errorへ更新する
- 古いテストIssueの誤起動を避けるため、既存stateは既定で処理しない
- 既存のcreatedを処理する場合は --include-existing と --min-issue を使用する
"""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import re
import shutil
import signal
import subprocess
import sys
import time
from datetime import datetime


REQUIRED_ENV_VARS = [
    "GITHUB_OWNER",
    "GITHUB_REPO",
    "AGENT_STATE_DIR",
    "TARGET_REPO_PATH",
]

POLL_INTERVAL_SECONDS = 5
FILE_READY_RETRIES = 10
FILE_READY_INTERVAL_SECONDS = 0.5
PROCESSABLE_STATUS = "created"
TERMINAL_STATUSES = {
    "implementation_started",
    "implementation_resumed",
    "waiting_decision",
    "waiting_approval",
    "approved",
    "completed",
    "rejected",
    "encoding_error",
    "daemon_error",
}

shutdown_requested = False


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def fail(message: str, code: int = 1) -> None:
    log(f"[ERROR] {message}")
    raise SystemExit(code)


def validate_environment_variables() -> None:
    missing: list[str] = []

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

    print("=" * 60)
    print("必要な環境変数はすべて設定されています。")
    print("=" * 60)
    print()


def require_command(name: str) -> None:
    if not shutil.which(name):
        fail(f"'{name}' がPATHにありません。")


def handle_shutdown(signum, frame) -> None:
    del signum, frame
    global shutdown_requested
    shutdown_requested = True
    log("終了要求を受け付けました。現在の処理完了後に終了します。")


def load_json_when_ready(path: pathlib.Path) -> dict | None:
    previous_size = -1

    for _ in range(FILE_READY_RETRIES):
        if not path.exists():
            return None

        try:
            current_size = path.stat().st_size
        except OSError:
            time.sleep(FILE_READY_INTERVAL_SECONDS)
            continue

        if current_size > 0 and current_size == previous_size:
            try:
                data = json.loads(path.read_text(encoding="utf-8-sig"))
                return data if isinstance(data, dict) else None
            except (json.JSONDecodeError, UnicodeDecodeError, OSError):
                pass

        previous_size = current_size
        time.sleep(FILE_READY_INTERVAL_SECONDS)

    log(f"state JSONが安定しないためスキップ: {path.name}")
    return None


def atomic_write_json(path: pathlib.Path, payload: dict) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def update_state(path: pathlib.Path, updates: dict) -> dict:
    state = load_json_when_ready(path) or {}
    state.update(updates)
    state["updatedAt"] = datetime.now().isoformat()
    atomic_write_json(path, state)
    log(f"state更新: {path.name} / status={state.get('status', '')}")
    return state


def parse_issue_number(path: pathlib.Path, state: dict) -> int | None:
    value = state.get("issueNumber")
    try:
        if value is not None:
            return int(value)
    except (TypeError, ValueError):
        pass

    match = re.fullmatch(r"issue-(\d+)\.json", path.name)
    return int(match.group(1)) if match else None


def acquire_daemon_lock(lock_path: pathlib.Path) -> None:
    try:
        descriptor = os.open(
            str(lock_path),
            os.O_CREAT | os.O_EXCL | os.O_WRONLY,
        )
    except FileExistsError:
        owner = ""
        try:
            owner = lock_path.read_text(encoding="utf-8").strip()
        except OSError:
            pass
        fail(
            "run_agent_daemon.py は既に起動している可能性があります。"
            + (f" PID={owner}" if owner else "")
            + f"\nロック: {lock_path}"
        )

    with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
        stream.write(str(os.getpid()))


def release_daemon_lock(lock_path: pathlib.Path) -> None:
    try:
        lock_path.unlink(missing_ok=True)
    except OSError as error:
        log(f"デーモンロック削除失敗: {error}")


def state_file_is_new(
    path: pathlib.Path,
    daemon_started_at: float,
    include_existing: bool,
) -> bool:
    if include_existing:
        return True

    try:
        return path.stat().st_mtime >= daemon_started_at
    except OSError:
        return False


def collect_created_issues(
    state_dir: pathlib.Path,
    daemon_started_at: float,
    include_existing: bool,
    min_issue: int,
    max_issue: int | None,
) -> list[tuple[int, pathlib.Path, dict]]:
    candidates: list[tuple[int, pathlib.Path, dict]] = []

    for path in sorted(state_dir.glob("issue-*.json")):
        if not state_file_is_new(path, daemon_started_at, include_existing):
            continue

        state = load_json_when_ready(path)
        if not state:
            continue

        issue_number = parse_issue_number(path, state)
        if issue_number is None:
            log(f"Issue番号を判定できません: {path.name}")
            continue

        if issue_number < min_issue:
            continue
        if max_issue is not None and issue_number > max_issue:
            continue

        status = str(state.get("status", "")).strip().lower()
        if status != PROCESSABLE_STATUS:
            continue

        candidates.append((issue_number, path, state))

    candidates.sort(key=lambda item: item[0])
    return candidates


def run_agent_command(
    python_executable: str,
    run_agent_path: pathlib.Path,
    issue_number: int,
    no_vscode: bool,
    manual: bool,
) -> list[str]:
    command = [python_executable, str(run_agent_path), str(issue_number)]
    if no_vscode:
        command.append("--no-vscode")
    if manual:
        command.append("--manual")
    return command


def process_issue(
    issue_number: int,
    state_path: pathlib.Path,
    run_agent_path: pathlib.Path,
    target_repo: pathlib.Path,
    no_vscode: bool,
    manual: bool,
) -> int:
    log("-" * 60)
    log(f"Issue #{issue_number} の自動実装を開始します。")

    update_state(
        state_path,
        {
            "status": "daemon_starting",
            "daemonPid": os.getpid(),
            "daemonStartedAt": datetime.now().isoformat(),
        },
    )

    command = run_agent_command(
        sys.executable,
        run_agent_path,
        issue_number,
        no_vscode,
        manual,
    )

    log(f"> {subprocess.list2cmdline(command)}")

    try:
        result = subprocess.run(
            command,
            cwd=str(run_agent_path.parent),
            check=False,
        )
    
        update_state(
            issue_number,
            {
                "runAgentExitCode": result.returncode,
                "stdout": result.stdout[-5000:],
                "stderr": result.stderr[-5000:],
            },
        )            
    except OSError as error:
        update_state(
            state_path,
            {
                "status": "daemon_error",
                "daemonError": str(error),
                "daemonFinishedAt": datetime.now().isoformat(),
            },
        )
        log(f"[ERROR] run_agent.pyを起動できません: {error}")
        return 1

    latest_state = load_json_when_ready(state_path) or {}
    latest_status = str(latest_state.get("status", "")).strip().lower()

    if result.returncode == 0:
        # run_agent.pyが状態を更新していれば、その状態を維持する。
        if latest_status == "daemon_starting":
            update_state(
                state_path,
                {
                    "status": "implementation_finished",
                    "daemonFinishedAt": datetime.now().isoformat(),
                    "runAgentExitCode": 0,
                },
            )
        else:
            update_state(
                state_path,
                {
                    "daemonFinishedAt": datetime.now().isoformat(),
                    "runAgentExitCode": 0,
                },
            )
        log(f"Issue #{issue_number} のrun_agent.pyが正常終了しました。")
        return 0

    update_state(
        state_path,
        {
            "status": "daemon_error",
            "daemonError": f"run_agent.py exit code {result.returncode}",
            "daemonFinishedAt": datetime.now().isoformat(),
            "runAgentExitCode": result.returncode,
        },
    )
    log(
        f"[ERROR] Issue #{issue_number} のrun_agent.pyが"
        f"終了コード {result.returncode} で終了しました。"
    )
    return result.returncode


def show_startup_summary(
    state_dir: pathlib.Path,
    run_agent_path: pathlib.Path,
    target_repo: pathlib.Path,
    interval: int,
    include_existing: bool,
    min_issue: int,
    max_issue: int | None,
    no_vscode: bool,
) -> None:
    log("=" * 60)
    log("run_agent 自動起動デーモン 起動")
    log(f"stateフォルダ: {state_dir}")
    log(f"run_agent.py: {run_agent_path}")
    log(f"対象リポジトリ: {target_repo}")
    log(f"監視間隔: {interval}秒")
    log(f"既存createdを処理: {include_existing}")
    log(f"最小Issue番号: {min_issue}")
    log(f"最大Issue番号: {max_issue if max_issue is not None else '指定なし'}")
    log(f"VS Code起動を省略: {no_vscode}")
    log("status=created のIssueを1件ずつ処理します。")
    log("終了するにはCtrl+Cを押してください。")
    log("=" * 60)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "stateフォルダを監視し、status=createdのIssueを"
            "run_agent.pyへ1件ずつ渡します。"
        )
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=POLL_INTERVAL_SECONDS,
        help=f"監視間隔秒数 (既定: {POLL_INTERVAL_SECONDS})",
    )
    parser.add_argument(
        "--include-existing",
        action="store_true",
        help="起動前から存在するstatus=createdのstateも処理する",
    )
    parser.add_argument(
        "--min-issue",
        type=int,
        default=0,
        help="処理する最小Issue番号",
    )
    parser.add_argument(
        "--max-issue",
        type=int,
        default=None,
        help="処理する最大Issue番号",
    )
    parser.add_argument(
        "--with-vscode",
        action="store_true",
        help="run_agent.pyによるVS Code起動を有効にする",
    )
    parser.add_argument(
        "--manual",
        action="store_true",
        help="run_agent.pyを手動貼り付けモードで起動する",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="対象Issueを最大1件処理して終了する",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="対象Issueを表示するだけでrun_agent.pyを起動しない",
    )
    args = parser.parse_args()

    validate_environment_variables()
    require_command("git")
    require_command("gh")

    state_dir = pathlib.Path(os.environ["AGENT_STATE_DIR"])
    target_repo = pathlib.Path(os.environ["TARGET_REPO_PATH"])
    script_dir = pathlib.Path(__file__).resolve().parent
    run_agent_path = script_dir / "run_agent.py"
    lock_dir = state_dir / "locks"
    lock_path = lock_dir / "run_agent_daemon.lock"

    state_dir.mkdir(parents=True, exist_ok=True)
    lock_dir.mkdir(parents=True, exist_ok=True)

    if not run_agent_path.exists():
        fail(f"run_agent.pyが見つかりません: {run_agent_path}")
    if not (target_repo / ".git").exists():
        fail(f"Gitリポジトリではありません: {target_repo}")
    if args.interval < 1:
        fail("--intervalは1以上を指定してください。")

    acquire_daemon_lock(lock_path)
    daemon_started_at = time.time()

    signal.signal(signal.SIGINT, handle_shutdown)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, handle_shutdown)

    show_startup_summary(
        state_dir,
        run_agent_path,
        target_repo,
        args.interval,
        args.include_existing,
        args.min_issue,
        args.max_issue,
        not args.with_vscode,
    )

    try:
        while not shutdown_requested:
            issues = collect_created_issues(
                state_dir,
                daemon_started_at,
                args.include_existing,
                args.min_issue,
                args.max_issue,
            )

            if not issues:
                if args.once:
                    log("処理対象のcreated Issueはありません。")
                    return
                time.sleep(args.interval)
                continue

            issue_number, path, state = issues[0]
            title = str(state.get("issueTitle", "")).strip()
            log(
                f"処理対象: Issue #{issue_number}"
                + (f" / {title}" if title else "")
            )

            if args.dry_run:
                log("dry-runのためrun_agent.pyは起動しません。")
                if args.once:
                    return
                # 同じ対象の連続表示を避ける。
                time.sleep(args.interval)
                continue

            process_issue(
                issue_number,
                path,
                run_agent_path,
                target_repo,
                no_vscode=not args.with_vscode,
                manual=args.manual,
            )

            if args.once:
                return

            # OneDriveとstate更新が落ち着くまで短く待つ。
            time.sleep(1)

    finally:
        release_daemon_lock(lock_path)
        log("run_agent 自動起動デーモンを終了しました。")


if __name__ == "__main__":
    main()
