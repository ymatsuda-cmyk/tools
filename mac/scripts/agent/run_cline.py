from __future__ import annotations

import argparse
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


def fail(message: str, code: int = 1) -> None:
    print(f"[ERROR] {message}", file=sys.stderr)
    raise SystemExit(code)


def run(cmd: list[str], cwd: Path | None = None, check: bool = True, capture: bool = True) -> subprocess.CompletedProcess[str]:
    print("> " + subprocess.list2cmdline(cmd))
    result = subprocess.run(
        cmd,
        cwd=str(cwd) if cwd else None,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=capture,
        check=False,
    )
    if capture:
        if result.stdout:
            print(result.stdout.rstrip())
        if result.stderr:
            print(result.stderr.rstrip(), file=sys.stderr)
    if check and result.returncode != 0:
        fail(f"Command failed ({result.returncode}): {subprocess.list2cmdline(cmd)}", result.returncode)
    return result


def require_command(name: str) -> None:
    if not shutil.which(name):
        fail(f"'{name}' が PATH にありません。インストール後、PowerShellを開き直してください。")


def git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run(["git", *args], cwd=repo, check=check)


def slugify(text: str, max_len: int = 48) -> str:
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return (text[:max_len].rstrip("-") or "task")


def detect_repo_name(repo: Path) -> str:
    cp = git(repo, "remote", "get-url", "origin")
    url = cp.stdout.strip()
    m = re.search(r"github\.com[/:]([^/]+/[^/]+?)(?:\.git)?$", url)
    if not m:
        fail("origin から owner/repo を判定できません。--github-repo を指定してください。")
    return m.group(1)


def load_issue(issue_number: int, github_repo: str, repo: Path) -> dict:
    cp = run([
        "gh", "issue", "view", str(issue_number),
        "--repo", github_repo,
        "--json", "number,title,body,url,labels,state"
    ], cwd=repo)
    try:
        return json.loads(cp.stdout)
    except json.JSONDecodeError as exc:
        fail(f"Issue JSONを解析できません: {exc}")


def build_prompt(issue: dict, branch: str) -> str:
    labels = ", ".join(x.get("name", "") for x in issue.get("labels", [])) or "なし"
    return f"""GitHub Issue #{issue['number']} を実装してください。

タイトル: {issue['title']}
URL: {issue['url']}
ラベル: {labels}

本文:
{issue.get('body') or '(本文なし)'}

作業ルール:
1. 現在のリポジトリと既存実装を確認する。
2. Issueの要件を満たす最小限かつ保守しやすい変更を行う。
3. 既存のコーディング規約に従う。
4. 必要なテストを追加または更新し、実行する。
5. 機密情報、APIキー、.envをコミットしない。
6. git commit、git push、PR作成は行わない。これらは外側のスクリプトが実行する。
7. 最後に、変更ファイル、実装内容、テスト結果、残課題を簡潔に報告する。

作業ブランチ: {branch}
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="GitHub Issueを取得してClineで実装し、PRを作成します。")
    parser.add_argument("issue", type=int, help="GitHub Issue番号")
    parser.add_argument("--repo-path", default=os.getenv("TARGET_REPO_PATH"), help="ローカルリポジトリのパス")
    parser.add_argument("--github-repo", default=os.getenv("GITHUB_REPOSITORY"), help="owner/repo")
    parser.add_argument("--base", default=os.getenv("BASE_BRANCH", "main"), help="PRのベースブランチ")
    parser.add_argument("--test-command", default=os.getenv("TEST_COMMAND", ""), help="Cline完了後に実行するテストコマンド")
    parser.add_argument("--auto-approve", action="store_true", default=os.getenv("CLINE_AUTO_APPROVE", "false").lower() == "true", help="Clineのツール操作を自動承認")
    parser.add_argument("--no-pr", action="store_true", help="commit/push/PR作成を行わない")
    args = parser.parse_args()

    if not args.repo_path:
        fail("--repo-path または環境変数 TARGET_REPO_PATH を指定してください。")
    repo = Path(args.repo_path).expanduser().resolve()
    if not (repo / ".git").exists():
        fail(f"Gitリポジトリではありません: {repo}")

    for command in ("git", "gh", "cline"):
        require_command(command)

    run(["gh", "auth", "status"], cwd=repo)
    github_repo = args.github_repo or detect_repo_name(repo)

    dirty = git(repo, "status", "--porcelain").stdout.strip()
    if dirty:
        fail("未コミットの変更があります。commitまたはstashしてから実行してください。")

    issue = load_issue(args.issue, github_repo, repo)
    if issue.get("state") != "OPEN":
        fail(f"Issue #{args.issue} はOPENではありません。state={issue.get('state')}")

    branch = f"issue-{args.issue}-{slugify(issue['title'])}"
    git(repo, "fetch", "origin")
    git(repo, "checkout", args.base)
    git(repo, "pull", "--ff-only", "origin", args.base)

    exists = git(repo, "show-ref", "--verify", "--quiet", f"refs/heads/{branch}", check=False).returncode == 0
    if exists:
        git(repo, "checkout", branch)
    else:
        git(repo, "checkout", "-b", branch)

    prompt = build_prompt(issue, branch)
    logs = repo / ".agent-logs"
    logs.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    prompt_file = logs / f"issue-{args.issue}-{stamp}-prompt.txt"
    output_file = logs / f"issue-{args.issue}-{stamp}-cline.jsonl"
    prompt_file.write_text(prompt, encoding="utf-8")

    cline_cmd = [
        "cline", "--cwd", str(repo), "--json",
        "--auto-approve", "true" if args.auto_approve else "false",
        prompt,
    ]
    print(f"[INFO] Cline開始: Issue #{args.issue}, auto-approve={args.auto_approve}")
    with output_file.open("w", encoding="utf-8") as out:
        process = subprocess.Popen(
            cline_cmd, cwd=str(repo), text=True, encoding="utf-8", errors="replace",
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        )
        assert process.stdout is not None
        for line in process.stdout:
            print(line, end="")
            out.write(line)
        return_code = process.wait()
    if return_code != 0:
        fail(f"Clineが終了コード {return_code} で失敗しました。ログ: {output_file}", return_code)

    if args.test_command:
        print(f"[INFO] テスト実行: {args.test_command}")
        test_cmd = ["powershell", "-NoProfile", "-Command", args.test_command] if os.name == "nt" else shlex.split(args.test_command)
        run(test_cmd, cwd=repo, capture=False)

    changed = git(repo, "status", "--porcelain").stdout.strip()
    if not changed:
        print("[INFO] 変更がないため終了します。")
        return

    git(repo, "diff", "--check")
    if args.no_pr:
        print(f"[OK] 実装完了。PR作成はスキップしました。ブランチ: {branch}")
        return

    git(repo, "add", "-A")
    git(repo, "commit", "-m", f"Implement issue #{args.issue}: {issue['title']}")
    git(repo, "push", "-u", "origin", branch)

    pr_body = f"Closes #{args.issue}\n\nCline CLIによる実装です。変更内容とテスト結果をレビューしてください。"
    cp = run([
        "gh", "pr", "create", "--repo", github_repo,
        "--base", args.base, "--head", branch,
        "--title", f"#{args.issue} {issue['title']}",
        "--body", pr_body,
    ], cwd=repo)
    pr_url = cp.stdout.strip().splitlines()[-1] if cp.stdout.strip() else ""
    if pr_url:
        run(["gh", "issue", "comment", str(args.issue), "--repo", github_repo, "--body", f"実装PRを作成しました: {pr_url}"], cwd=repo)
    print(f"[OK] PR作成完了: {pr_url or '(URLを取得できませんでした)'}")
    print(f"[INFO] Clineログ: {output_file}")


if __name__ == "__main__":
    main()
