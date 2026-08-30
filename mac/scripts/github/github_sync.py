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
            f"\nERR : {result.stderr}"
        )

    return result


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

        print(f"rebase失敗: {e}")
        raise


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

