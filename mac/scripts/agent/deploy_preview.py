"""
承認前プレビューを tools-beta の GitHub Pages へ公開する。

処理:
1. C:\repo\tools の静的サイトを C:\repo\tools-beta へ同期
2. tools-beta/main へ commit / push
3. 変更ファイルからプレビューURLを生成
4. 呼び出し元が利用できるようJSONを標準出力する

前提:
- tools-beta はローカルへclone済み
- tools-betaのGitHub Pagesは main / (root) で公開済み
- gh auth / git push が利用可能
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import os
import pathlib
import shutil
import subprocess
import sys
from datetime import datetime
from urllib.parse import quote


DEFAULT_SOURCE_REPO = pathlib.Path(
    os.getenv("TARGET_REPO_PATH", r"C:\repo\tools")
)
DEFAULT_PREVIEW_REPO = pathlib.Path(
    os.getenv("PREVIEW_REPO_PATH", r"C:\repo\tools-beta")
)
DEFAULT_PREVIEW_BRANCH = os.getenv("PREVIEW_BRANCH", "main")
DEFAULT_PREVIEW_BASE_URL = os.getenv(
    "PREVIEW_BASE_URL",
    "https://ymatsuda-cmyk.github.io/tools-beta/",
)

# リポジトリ直下または任意階層で除外する名前
EXCLUDED_NAMES = {
    ".git",
    ".github",
    ".venv",
    "venv",
    "node_modules",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".idea",
    ".vscode",
    "old",
    "state",
    "approval",
    "waiting",
    "reply",
    "request",
    "decision",
    "question",
}

# 公開してはいけない可能性が高いファイル
EXCLUDED_PATTERNS = (
    ".env",
    ".env.*",
    "*.pem",
    "*.key",
    "*.pfx",
    "*.p12",
    "*.crt",
    "*.cer",
    "id_rsa*",
    "*.log",
    "*.tmp",
    "*.bak",
    "*.pyc",
    ".agent-question.json",
    "settings.json",
    "launch.json",
)

# tools-beta側で同期削除の対象外にするもの
PREVIEW_PRESERVE_NAMES = {
    ".git",
}


class DeployError(RuntimeError):
    pass


def log(message: str) -> None:
    print(
        f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}",
        file=sys.stderr,
        flush=True,
    )


def run_command(
    command: list[str],
    cwd: pathlib.Path,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    log(f"> {subprocess.list2cmdline(command)}")
    result = subprocess.run(
        command,
        cwd=str(cwd),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    )

    if result.stdout.strip():
        print(result.stdout.rstrip(), file=sys.stderr)
    if result.stderr.strip():
        print(result.stderr.rstrip(), file=sys.stderr)

    if check and result.returncode != 0:
        raise DeployError(
            f"コマンド失敗 ({result.returncode}): "
            f"{subprocess.list2cmdline(command)}"
        )
    return result


def git(
    repo: pathlib.Path,
    *args: str,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    return run_command(["git", *args], repo, check=check)


def validate_repository(path: pathlib.Path, label: str) -> None:
    if not path.exists():
        raise DeployError(f"{label}が見つかりません: {path}")
    if not (path / ".git").exists():
        raise DeployError(f"{label}はGitリポジトリではありません: {path}")


def matches_excluded_pattern(name: str) -> bool:
    lower_name = name.lower()
    return any(
        fnmatch.fnmatch(lower_name, pattern.lower())
        for pattern in EXCLUDED_PATTERNS
    )


def should_exclude(relative_path: pathlib.Path) -> bool:
    parts = relative_path.parts
    if any(part in EXCLUDED_NAMES for part in parts):
        return True
    return any(matches_excluded_pattern(part) for part in parts)


def collect_source_files(source_repo: pathlib.Path) -> dict[str, pathlib.Path]:
    files: dict[str, pathlib.Path] = {}

    for source in source_repo.rglob("*"):
        if not source.is_file():
            continue

        relative = source.relative_to(source_repo)
        if should_exclude(relative):
            continue

        relative_posix = relative.as_posix()
        files[relative_posix] = source

    if not files:
        raise DeployError("プレビューへ公開できるファイルがありません。")

    return files


def remove_preview_contents(preview_repo: pathlib.Path) -> None:
    """tools-betaの作業ツリーを空にする。.gitは保持する。"""
    for child in preview_repo.iterdir():
        if child.name in PREVIEW_PRESERVE_NAMES:
            continue

        if child.is_dir() and not child.is_symlink():
            shutil.rmtree(child)
        else:
            child.unlink(missing_ok=True)


def copy_source_files(
    source_files: dict[str, pathlib.Path],
    preview_repo: pathlib.Path,
) -> None:
    for relative, source in source_files.items():
        destination = preview_repo / pathlib.PurePosixPath(relative)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)


def add_nojekyll(preview_repo: pathlib.Path) -> None:
    """Jekyll処理を不要にし、静的ファイルをそのまま公開する。"""
    (preview_repo / ".nojekyll").write_text("", encoding="utf-8")


def prepare_preview_branch(
    preview_repo: pathlib.Path,
    branch: str,
) -> None:
    dirty = git(preview_repo, "status", "--porcelain").stdout.strip()
    if dirty:
        raise DeployError(
            "tools-betaに未コミット変更があります。"
            "内容を確認してcommitまたは破棄してください。"
        )

    git(preview_repo, "fetch", "origin")

    local_exists = git(
        preview_repo,
        "show-ref",
        "--verify",
        "--quiet",
        f"refs/heads/{branch}",
        check=False,
    ).returncode == 0

    if local_exists:
        git(preview_repo, "checkout", branch)
    else:
        remote_exists = git(
            preview_repo,
            "show-ref",
            "--verify",
            "--quiet",
            f"refs/remotes/origin/{branch}",
            check=False,
        ).returncode == 0

        if remote_exists:
            git(preview_repo, "checkout", "-b", branch, f"origin/{branch}")
        else:
            git(preview_repo, "checkout", "--orphan", branch)

    # リモートブランチがある場合のみ最新化
    remote_exists = git(
        preview_repo,
        "show-ref",
        "--verify",
        "--quiet",
        f"refs/remotes/origin/{branch}",
        check=False,
    ).returncode == 0

    if remote_exists:
        git(preview_repo, "pull", "--ff-only", "origin", branch)


def normalize_changed_files(values: list[str]) -> list[str]:
    normalized: list[str] = []
    for value in values:
        for item in value.split(","):
            path = item.strip().strip('"').replace("\\", "/")
            if path and path not in normalized:
                normalized.append(path)
    return normalized


def determine_preview_path(changed_files: list[str]) -> str:
    """index.htmlの親フォルダをプレビュー先とする。"""
    for changed_file in changed_files:
        pure_path = pathlib.PurePosixPath(changed_file)
        if pure_path.name.lower() == "index.html":
            parent = pure_path.parent.as_posix()
            return "" if parent == "." else parent.strip("/")

    # ディレクトリが渡された場合
    for changed_file in changed_files:
        pure_path = pathlib.PurePosixPath(changed_file.rstrip("/"))
        if pure_path.suffix == "" and pure_path.as_posix() not in ("", "."):
            return pure_path.as_posix().strip("/")

    return ""


def build_preview_url(base_url: str, preview_path: str) -> str:
    normalized_base = base_url.rstrip("/") + "/"
    if not preview_path:
        return normalized_base

    encoded_path = "/".join(
        quote(part, safe="-._~")
        for part in preview_path.split("/")
        if part
    )
    return f"{normalized_base}{encoded_path}/"


def commit_and_push(
    preview_repo: pathlib.Path,
    branch: str,
    issue_number: int | None,
    source_branch: str,
) -> tuple[bool, str]:
    git(preview_repo, "add", "-A")
    staged = git(preview_repo, "diff", "--cached", "--name-only").stdout.strip()

    if not staged:
        log("tools-betaに差分はありません。既存プレビューを使用します。")
        return False, ""

    issue_text = f"Issue #{issue_number}" if issue_number is not None else "manual"
    commit_message = f"Preview {issue_text} from {source_branch}"

    git(preview_repo, "commit", "-m", commit_message)
    git(preview_repo, "push", "-u", "origin", branch)

    commit_sha = git(preview_repo, "rev-parse", "HEAD").stdout.strip()
    return True, commit_sha


def deploy_preview(
    source_repo: pathlib.Path,
    preview_repo: pathlib.Path,
    preview_branch: str,
    preview_base_url: str,
    issue_number: int | None,
    changed_files: list[str],
) -> dict:
    validate_repository(source_repo, "ソースリポジトリ")
    validate_repository(preview_repo, "プレビューリポジトリ")

    source_branch = git(source_repo, "branch", "--show-current").stdout.strip()
    if not source_branch:
        raise DeployError("ソースブランチを取得できません。")

    prepare_preview_branch(preview_repo, preview_branch)

    source_files = collect_source_files(source_repo)
    log(f"公開対象ファイル数: {len(source_files)}")

    remove_preview_contents(preview_repo)
    copy_source_files(source_files, preview_repo)
    add_nojekyll(preview_repo)

    pushed, commit_sha = commit_and_push(
        preview_repo,
        preview_branch,
        issue_number,
        source_branch,
    )

    preview_path = determine_preview_path(changed_files)
    preview_url = build_preview_url(preview_base_url, preview_path)

    return {
        "status": "deployed" if pushed else "unchanged",
        "issueNumber": issue_number,
        "sourceRepository": str(source_repo),
        "sourceBranch": source_branch,
        "previewRepository": str(preview_repo),
        "previewBranch": preview_branch,
        "changedFiles": changed_files,
        "previewPath": preview_path,
        "previewUrl": preview_url,
        "commitSha": commit_sha,
        "deployedAt": datetime.now().isoformat(),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="toolsの静的サイトをtools-betaへ同期し、PagesプレビューURLを返します。"
    )
    parser.add_argument(
        "--source",
        type=pathlib.Path,
        default=DEFAULT_SOURCE_REPO,
        help=f"ソースリポジトリ (既定: {DEFAULT_SOURCE_REPO})",
    )
    parser.add_argument(
        "--preview",
        type=pathlib.Path,
        default=DEFAULT_PREVIEW_REPO,
        help=f"プレビューリポジトリ (既定: {DEFAULT_PREVIEW_REPO})",
    )
    parser.add_argument(
        "--branch",
        default=DEFAULT_PREVIEW_BRANCH,
        help=f"tools-betaの公開ブランチ (既定: {DEFAULT_PREVIEW_BRANCH})",
    )
    parser.add_argument(
        "--base-url",
        default=DEFAULT_PREVIEW_BASE_URL,
        help=f"GitHub PagesのベースURL (既定: {DEFAULT_PREVIEW_BASE_URL})",
    )
    parser.add_argument(
        "--issue",
        type=int,
        default=None,
        help="GitHub Issue番号",
    )
    parser.add_argument(
        "--changed-file",
        action="append",
        default=[],
        help="変更ファイル。複数指定可能。カンマ区切りも可。",
    )
    parser.add_argument(
        "--output",
        type=pathlib.Path,
        default=None,
        help="結果JSONの保存先。未指定なら標準出力のみ。",
    )
    args = parser.parse_args()

    try:
        changed_files = normalize_changed_files(args.changed_file)
        result = deploy_preview(
            source_repo=args.source.resolve(),
            preview_repo=args.preview.resolve(),
            preview_branch=args.branch,
            preview_base_url=args.base_url,
            issue_number=args.issue,
            changed_files=changed_files,
        )

        text = json.dumps(result, ensure_ascii=False, indent=2)
        print(text)

        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(text, encoding="utf-8")
            log(f"結果JSON出力: {args.output}")

    except DeployError as error:
        error_result = {
            "status": "error",
            "message": str(error),
            "deployedAt": datetime.now().isoformat(),
        }
        print(json.dumps(error_result, ensure_ascii=False, indent=2))
        log(f"[ERROR] {error}")
        raise SystemExit(1)
    except Exception as error:
        error_result = {
            "status": "error",
            "message": f"予期しないエラー: {error}",
            "deployedAt": datetime.now().isoformat(),
        }
        print(json.dumps(error_result, ensure_ascii=False, indent=2))
        log(f"[ERROR] 予期しないエラー: {error}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
