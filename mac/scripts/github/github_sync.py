#!/usr/bin/env python3

import argparse
import json
import shutil
import subprocess
from pathlib import Path


def run(cmd):
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"\nCMD : {' '.join(cmd)}"
            f"\nOUT : {result.stdout}"
            f"\nERR : {result.stderr}"
        )

    return result


def run_quiet(cmd):
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True
    )


def git(repo_root, *args):
    return [
        "git",
        "-C",
        str(repo_root),
        *args
    ]


def unmerged_files(repo_root):
    result = run_quiet(
        git(repo_root, "diff", "--name-only", "--diff-filter=U")
    )

    return [
        line.strip()
        for line in result.stdout.splitlines()
        if line.strip()
    ]


def abort_unfinished(repo_root):
    """前回の実行が残したrebase/mergeを畳む。残っているとpullが弾かれる。"""

    git_dir = Path(
        run(git(repo_root, "rev-parse", "--absolute-git-dir")).stdout.strip()
    )

    if (git_dir / "rebase-merge").exists() or (git_dir / "rebase-apply").exists():
        print("前回のrebaseが途中のままだったため中断します")
        run_quiet(git(repo_root, "rebase", "--abort"))

    elif (git_dir / "MERGE_HEAD").exists():
        print("前回のmergeが途中のままだったため中断します")
        run_quiet(git(repo_root, "merge", "--abort"))


def sync_config(config_path):

    config = json.loads(
        Path(config_path).read_text(encoding="utf-8")
    )

    if not config.get("enabled", True):
        print("無効設定のためスキップ")
        return

    source_folder = Path(config["source_folder"])

    repo_root = Path(config["repository_root"])

    target_folder = (
        repo_root /
        config["repo_subfolder"]
    )

    print(f"同期開始 : {config['id']}")

    abort_unfinished(repo_root)

    target_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    #
    # ファイルコピー
    #
    copied = 0

    for source_file in source_folder.glob("*"):

        if not source_file.is_file():
            continue

        target_file = (
            target_folder /
            source_file.name
        )

        shutil.copy2(
            source_file,
            target_file
        )

        copied += 1

    print(f"コピー : {copied}件")

    #
    # add
    #
    run([
        "git",
        "-C",
        str(repo_root),
        "add",
        config["repo_subfolder"]
    ])

    #
    # 同期対象外の衝突
    #
    # 同期フォルダの衝突は上のコピーとaddで解消済み。それ以外は手で直すしかない
    #
    conflicts = unmerged_files(repo_root)

    if conflicts:
        raise RuntimeError(
            "未解決の衝突が残っているため中止します:\n  "
            + "\n  ".join(conflicts)
        )

    #
    # 差分確認
    #
    diff = subprocess.run([
        "git",
        "-C",
        str(repo_root),
        "diff",
        "--cached",
        "--quiet"
    ])

    if diff.returncode == 0:
        print("差分なし")
        return

    #
    # add後
    #

    try:

        run([
            "git",
            "-C",
            str(repo_root),
            "pull",
            "--rebase",
            "--autostash",
            config["remote_name"],
            config["branch"]
        ])

    except Exception as e:

        # 途中のまま残すと次回の実行もpullで弾かれる
        abort_unfinished(repo_root)

        print(f"rebase失敗: {e}")
        raise

    # rebase後に差分が吸収される場合があるため、commit直前でも再確認する
    diff_after_pull = subprocess.run([
        "git",
        "-C",
        str(repo_root),
        "diff",
        "--cached",
        "--quiet"
    ])

    if diff_after_pull.returncode == 0:
        print("差分なし（pull後に同期済み）")
        return


    #
    # commit
    #

    run([
        "git",
        "-C",
        str(repo_root),
        "commit",
        "-m",
        config["commit_message"]
    ])


    #
    # push
    #

    run([
        "git",
        "-C",
        str(repo_root),
        "push",
        config["remote_name"],
        config["branch"]
    ])

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--config",
        required=True
    )

    args = parser.parse_args()

    sync_config(args.config)


if __name__ == "__main__":
    main()

