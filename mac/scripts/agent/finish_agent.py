"""Phase4: 承認済みの変更を検査し、commit、push、PR作成まで実行する。"""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import shutil
import subprocess
import sys
from datetime import datetime

REQUIRED_ENV_VARS = [
    "GITHUB_OWNER", "GITHUB_REPO", "AGENT_REPLY_DIR",
    "AGENT_WAITING_DIR", "AGENT_APPROVAL_DIR", "AGENT_STATE_DIR",
    "AGENT_DONE_DIR", "TARGET_REPO_PATH",
]

TEXT_SUFFIXES = {
    ".html", ".htm", ".css", ".js", ".mjs", ".json", ".md",
    ".txt", ".svg", ".xml", ".gs", ".py",
}
FALLBACK_ENCODINGS = ("cp932", "shift_jis", "euc_jp")
MOJIBAKE_PATTERNS = (
    "縺", "繧", "繝", "譁", "荳", "菴", "髱", "邨", "蜈", "蛯",
    "陦", "蜷", "逕", "隕", "譛", "螟", "繚", "莉", "逡", "繧・",
)
MOJIBAKE_THRESHOLD = 5
APPROVE_VALUES = {"approve", "approved", "ok", "yes", "承認"}
REJECT_VALUES = {"reject", "rejected", "ng", "no", "却下"}


def validate_environment_variables() -> None:
    missing = []
    print("\n" + "=" * 60)
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
        print("\n不足: " + ", ".join(missing))
        raise SystemExit(1)
    print("=" * 60)
    print("必要な環境変数はすべて設定されています。")
    print("=" * 60 + "\n")


validate_environment_variables()
GITHUB_OWNER = os.environ["GITHUB_OWNER"]
GITHUB_REPO = os.environ["GITHUB_REPO"]
GITHUB_REPOSITORY = f"{GITHUB_OWNER}/{GITHUB_REPO}"
REPLY_DIR = pathlib.Path(os.environ["AGENT_REPLY_DIR"])
WAITING_DIR = pathlib.Path(os.environ["AGENT_WAITING_DIR"])
APPROVAL_DIR = pathlib.Path(os.environ["AGENT_APPROVAL_DIR"])
STATE_DIR = pathlib.Path(os.environ["AGENT_STATE_DIR"])
DONE_DIR = pathlib.Path(os.environ["AGENT_DONE_DIR"])
TARGET_REPO = pathlib.Path(os.environ["TARGET_REPO_PATH"])
LOCK_DIR = STATE_DIR / "locks"


def log(message: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}", flush=True)


def fail(message: str, code: int = 1) -> None:
    log(f"[ERROR] {message}")
    raise SystemExit(code)


def run_command(command: list[str], cwd: pathlib.Path | None = None,
                check: bool = True) -> subprocess.CompletedProcess[str]:
    log(f"> {subprocess.list2cmdline(command)}")
    result = subprocess.run(
        command, cwd=str(cwd) if cwd else None, text=True,
        encoding="utf-8", errors="replace", capture_output=True, check=False,
    )
    if result.stdout and result.stdout.strip():
        print(result.stdout.rstrip())
    if result.stderr and result.stderr.strip():
        print(result.stderr.rstrip(), file=sys.stderr)
    if check and result.returncode != 0:
        fail(f"コマンド失敗 ({result.returncode}): {subprocess.list2cmdline(command)}")
    return result


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run_command(["git", *args], cwd=TARGET_REPO, check=check)


def require_command(name: str) -> None:
    if not shutil.which(name):
        fail(f"'{name}' がPATHにありません。")


def validate_directories() -> None:
    if not (TARGET_REPO / ".git").exists():
        fail(f"Gitリポジトリではありません: {TARGET_REPO}")
    for directory in (
        REPLY_DIR, WAITING_DIR, APPROVAL_DIR, STATE_DIR, DONE_DIR, LOCK_DIR,
    ):
        directory.mkdir(parents=True, exist_ok=True)


def state_path(issue_number: int) -> pathlib.Path:
    return STATE_DIR / f"issue-{issue_number}.json"


def load_state(issue_number: int) -> dict:
    path = state_path(issue_number)
    if not path.exists():
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
    state_path(issue_number).write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log(f"状態ファイル更新: issue-{issue_number}.json")


def acquire_issue_lock(issue_number: int) -> pathlib.Path | None:
    """同じapprovalが複数回処理されることを防ぐ。"""
    path = LOCK_DIR / f"issue-{issue_number}.lock"
    try:
        fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        log(f"Issue #{issue_number} は別プロセスが処理中のためスキップします。")
        return None
    with os.fdopen(fd, "w", encoding="utf-8") as stream:
        stream.write(str(os.getpid()))
    return path


def release_issue_lock(path: pathlib.Path | None) -> None:
    if path:
        try:
            path.unlink(missing_ok=True)
        except OSError as error:
            log(f"ロック削除失敗: {error}")


def load_waiting_or_state(issue_number: int) -> dict:
    """waitingがPower Automateで削除済みでもstateから復元する。"""
    path = WAITING_DIR / f"issue-{issue_number}.json"
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as error:
            fail(f"承認待ちファイル読込失敗: {error}")
    state = load_state(issue_number)
    if state.get("branch"):
        log("waitingファイルが無いためstateファイルを使用します。")
        return state
    fail(f"承認待ち情報が見つかりません: {path}")
    return {}


def load_approval(issue_number: int) -> str:
    path = APPROVAL_DIR / f"issue-{issue_number}.json"
    if not path.exists():
        log(f"承認結果ファイルが見つかりません: {path.name}")
        return ""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as error:
        log(f"承認結果ファイル読込失敗: {error}")
        return ""
    for key in ("result", "approval", "value", "action"):
        value = str(data.get(key, "")).strip().lower()
        if value in APPROVE_VALUES:
            return "approve"
        if value in REJECT_VALUES:
            return "reject"
    return ""


def collect_changed_files() -> list[str]:
    lines = git("status", "--porcelain", check=False).stdout.strip().splitlines()
    files = []
    for line in lines:
        path = line[3:].strip().strip('"')
        if path:
            files.append(path.rstrip("/"))
    return files


def expand_text_files(changed_files: list[str]) -> list[tuple[str, pathlib.Path]]:
    expanded: dict[str, pathlib.Path] = {}
    for relative_path in changed_files:
        target = TARGET_REPO / relative_path
        candidates = list(target.rglob("*")) if target.is_dir() else [target]
        for candidate in candidates:
            if candidate.is_file() and candidate.suffix.lower() in TEXT_SUFFIXES:
                relative = candidate.relative_to(TARGET_REPO).as_posix()
                expanded[relative] = candidate
    return sorted(expanded.items())


def normalize_encoding(changed_files: list[str]) -> list[str]:
    converted = []
    for relative_path, target in expand_text_files(changed_files):
        try:
            raw = target.read_bytes()
        except OSError as error:
            log(f"読込失敗: {relative_path} / {error}")
            continue
        if not raw:
            continue
        if raw.startswith(b"\xef\xbb\xbf"):
            target.write_text(raw[3:].decode("utf-8"), encoding="utf-8", newline="")
            log(f"BOM除去: {relative_path}")
            converted.append(relative_path)
            continue
        try:
            raw.decode("utf-8")
            continue
        except UnicodeDecodeError:
            pass
        decoded = None
        used = ""
        for encoding in FALLBACK_ENCODINGS:
            try:
                decoded = raw.decode(encoding)
                used = encoding
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            log(f"エンコーディング判定不可: {relative_path}")
            continue
        target.write_text(decoded, encoding="utf-8", newline="")
        log(f"UTF-8へ変換 ({used}): {relative_path}")
        converted.append(relative_path)
    return converted


def detect_mojibake(changed_files: list[str]) -> list[dict]:
    detected = []
    for relative_path, target in expand_text_files(changed_files):
        try:
            text = target.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError) as error:
            detected.append({
                "path": relative_path, "score": MOJIBAKE_THRESHOLD,
                "reason": f"UTF-8として読めません: {error}", "patterns": [],
            })
            continue
        matches = {p: text.count(p) for p in MOJIBAKE_PATTERNS if text.count(p)}
        score = sum(matches.values())
        if score >= MOJIBAKE_THRESHOLD:
            detected.append({
                "path": relative_path, "score": score,
                "reason": "典型的な文字化け文字を検出",
                "patterns": [{"text": p, "count": c} for p, c in matches.items()],
            })
    return detected


def write_json(directory: pathlib.Path, file_name: str, payload: dict) -> bool:
    try:
        path = directory / file_name
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        log(f"JSON出力: {path}")
        return True
    except OSError as error:
        log(f"JSON出力失敗: {error}")
        return False


def notify_teams(message_type: str, issue_number: int, message: str,
                 extra: dict | None = None) -> None:
    message_id = load_message_id(issue_number)
    if not message_id:
        log("親メッセージIDが無いためTeams通知をスキップします。")
        return
    payload = {
        "type": message_type, "messageId": message_id,
        "issueNumber": issue_number, "status": message_type, "message": message,
    }
    if extra:
        payload.update(extra)
    write_json(
        REPLY_DIR,
        f"issue-{issue_number}-{message_type}-{datetime.now():%H%M%S}.json",
        payload,
    )


def stop_on_mojibake(issue_number: int, detected: list[dict]) -> None:
    details = "\n".join(f"- {x['path']} (score={x['score']})" for x in detected)
    update_state(issue_number, {"status": "encoding_error", "encodingErrors": detected})
    notify_teams(
        "encoding_error", issue_number,
        f"⚠️ Issue #{issue_number} の変更で文字化けを検出しました\n\n"
        f"対象ファイル:\n{details}\n\nコミットとPR作成を停止しました。",
        {"encodingErrors": detected},
    )
    fail("文字化けが検出されたためコミットを中止しました。", 2)


def archive_file(source: pathlib.Path) -> None:
    if not source.exists():
        return
    target = DONE_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{source.name}"
    try:
        shutil.move(str(source), str(target))
        log(f"処理済みへ移動: {target.name}")
    except OSError as error:
        log(f"移動失敗: {source.name} / {error}")


def find_existing_pr(branch: str) -> dict | None:
    result = run_command(
        [
            "gh", "pr", "list", "--repo", GITHUB_REPOSITORY,
            "--head", branch, "--state", "open",
            "--json", "number,title,url,headRefName",
            "--limit", "1",
        ],
        cwd=TARGET_REPO,
    )
    try:
        items = json.loads(result.stdout or "[]")
    except json.JSONDecodeError as error:
        fail(f"PR一覧JSONを解析できません: {error}")
    return items[0] if items else None


def branch_has_changes_from_base(branch: str, base_branch: str) -> bool:
    result = git("rev-list", "--count", f"{base_branch}..{branch}", check=False)
    if result.returncode != 0:
        return False
    try:
        return int(result.stdout.strip() or "0") > 0
    except ValueError:
        return False


def push_branch(branch: str) -> None:
    git("push", "-u", "origin", branch)


def commit_and_push(issue_number: int, issue_title: str, branch: str) -> None:
    git("add", "-A")
    staged = git("diff", "--cached", "--name-only").stdout.strip()
    if not staged:
        fail("git add後もコミット対象がありません。")
    git("commit", "-m", f"{issue_title} (Closes #{issue_number})")
    push_branch(branch)


def create_pull_request(issue_number: int, issue_title: str, branch: str,
                        base_branch: str, changed_files: list[str]) -> str:
    file_list = "\n".join(f"- {path}" for path in changed_files) or "- コミット済み変更"
    body = f"""Closes #{issue_number}

## 変更内容

{file_list}

## 確認内容

- Teamsで実装内容を承認済み
- 既存ページへの影響を確認

## 備考

本PRはAIエージェントによる実装です。
"""
    result = run_command(
        [
            "gh", "pr", "create", "--repo", GITHUB_REPOSITORY,
            "--base", base_branch, "--head", branch,
            "--title", f"#{issue_number} {issue_title}", "--body", body,
        ],
        cwd=TARGET_REPO,
    )
    urls = [x.strip() for x in result.stdout.splitlines() if x.strip().startswith("http")]
    return urls[-1] if urls else ""


def comment_on_issue(issue_number: int, body: str) -> None:
    run_command(
        ["gh", "issue", "comment", str(issue_number),
         "--repo", GITHUB_REPOSITORY, "--body", body],
        cwd=TARGET_REPO, check=False,
    )


def complete_processing(issue_number: int, issue_title: str, branch: str,
                        pr_url: str, changed_files: list[str]) -> None:
    update_state(issue_number, {
        "status": "completed", "branch": branch, "pullRequestUrl": pr_url,
    })
    file_list = "\n".join(f"- {x}" for x in changed_files) or "- コミット済み変更"
    notify_teams(
        "completed", issue_number,
        f"🎉 Issue #{issue_number} が完了しました\n\n"
        f"タイトル:\n{issue_title}\n\n変更ファイル:\n{file_list}\n\n"
        f"PR:\n{pr_url or '(URL取得失敗)'}\n\nレビューをお願いします。",
        {"branch": branch, "pullRequestUrl": pr_url},
    )
    archive_file(WAITING_DIR / f"issue-{issue_number}.json")
    archive_file(APPROVAL_DIR / f"issue-{issue_number}.json")


def handle_reject(issue_number: int, branch: str) -> None:
    git("checkout", "--", ".", check=False)
    git("clean", "-fd", check=False)
    update_state(issue_number, {"status": "rejected"})
    notify_teams(
        "rejected", issue_number,
        f"❌ Issue #{issue_number} の実装が却下されました。変更を破棄しました。",
        {"branch": branch},
    )
    archive_file(WAITING_DIR / f"issue-{issue_number}.json")
    archive_file(APPROVAL_DIR / f"issue-{issue_number}.json")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("issue", type=int)
    parser.add_argument("--base", default=os.getenv("BASE_BRANCH", "main"))
    parser.add_argument("--approve", action="store_true")
    parser.add_argument("--reject", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    validate_directories()
    for command in ("git", "gh"):
        require_command(command)

    lock_path = acquire_issue_lock(args.issue)
    if lock_path is None:
        return

    try:
        log("=" * 60)
        log("PR作成エージェント 起動")
        log(f"対象リポジトリ: {TARGET_REPO}")
        log(f"Issue: #{args.issue}")
        log("=" * 60)

        # 既に完了済みならapprovalを退避して静かに終了する。
        state = load_state(args.issue)
        if state.get("status") == "completed" and state.get("pullRequestUrl"):
            log("既に完了済みです。重複処理をスキップします。")
            archive_file(APPROVAL_DIR / f"issue-{args.issue}.json")
            archive_file(WAITING_DIR / f"issue-{args.issue}.json")
            return

        waiting = load_waiting_or_state(args.issue)
        branch = str(waiting.get("branch", "")).strip()
        issue_title = str(waiting.get("issueTitle", "")).strip() or f"Issue #{args.issue}"
        if not branch:
            fail("ブランチ情報がありません。")

        decision = "reject" if args.reject else "approve" if args.approve else load_approval(args.issue)
        if not decision:
            fail("承認結果を確認できません。")
        log(f"ブランチ: {branch}")
        log(f"承認結果: {decision}")

        current = git("branch", "--show-current").stdout.strip()
        if current != branch:
            git("checkout", branch)

        if decision == "reject":
            handle_reject(args.issue, branch)
            return

        changed_files = collect_changed_files()

        if changed_files:
            log(f"変更ファイル数: {len(changed_files)}")
            normalize_encoding(changed_files)
            detected = detect_mojibake(changed_files)
            if detected:
                stop_on_mojibake(args.issue, detected)
        else:
            log("作業ツリーはクリーンです。既存コミットとPRを確認します。")

        if args.dry_run:
            existing = find_existing_pr(branch)
            if existing:
                log(f"既存PR: #{existing['number']} {existing['url']}")
            else:
                log("既存PRはありません。")
            log("dry-runのためcommit/push/PR作成は実行しません。")
            return

        notify_teams(
            "approved", args.issue,
            f"✅ Issue #{args.issue} が承認されました。PR処理を開始します。",
            {"branch": branch},
        )

        if changed_files:
            commit_and_push(args.issue, issue_title, branch)
        else:
            # コミット済みでもリモート未反映の可能性があるためpushする。
            push_branch(branch)

        existing_pr = find_existing_pr(branch)
        if existing_pr:
            pr_url = str(existing_pr.get("url", ""))
            log(f"既存PRを使用します: #{existing_pr.get('number')} {pr_url}")
        else:
            if not branch_has_changes_from_base(branch, args.base):
                fail(
                    f"{args.base} と {branch} の間にPR対象コミットがありません。"
                )
            pr_url = create_pull_request(
                args.issue, issue_title, branch, args.base, changed_files
            )
            if pr_url:
                comment_on_issue(args.issue, f"実装PRを作成しました。\n\n{pr_url}")

        complete_processing(
            args.issue, issue_title, branch, pr_url, changed_files
        )
        log("=" * 60)
        log("Phase4 完了。")
        log("=" * 60)

    finally:
        release_issue_lock(lock_path)


if __name__ == "__main__":
    main()
