#!/usr/bin/env python3
"""
github_sync.py

source_folder の中身をそのまま repository_root/repo_subfolder にコピーし、commit / push する。
暗号化が要るときは github_aes_sync.py を使うこと。

- 設定ごとに GitHub アカウント（SSH鍵 / トークン / コミット者）を切り替えられる
  （github_aes_sync.py と同じ仕組み）。アカウントの秘密情報はリポジトリ外の
  accounts ファイルに置く。

使い方:
  python3 github_sync.py --config config/schedule.json
  python3 github_sync.py --config a.json --config b.json   # 複数
  python3 github_sync.py --config config/                  # フォルダ内の *.json 全部
  python3 github_sync.py --config a.json --dry-run         # commit/pushしない
"""

import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

# 既定のアカウント定義ファイル（リポジトリ外。トークン等はここにだけ置く）。
# github_aes_sync.py と同じ場所を見るので、1つのファイルで両方のスクリプトから使える
DEFAULT_ACCOUNTS_FILE = "~/.config/github_aes_sync/accounts.json"


# ============================================================
# git 実行（アカウントごとの環境を差し込む。github_aes_sync.py と同じ仕組み）
# ============================================================

class GitRunner:
    """アカウント設定を反映して git を実行する。

    - commit 者: -c user.name / -c user.email
    - SSH: GIT_SSH_COMMAND で鍵を固定（IdentitiesOnly=yes で agent の他の鍵を使わない）
    - トークン: GIT_CONFIG_* 環境変数で http.extraHeader を渡す
      （コマンドライン引数に載せないので ps に出ない。git 2.31 以降）
    """

    def __init__(self, repo_root, account=None):
        self.repo_root = Path(repo_root)
        self.account = account or {}
        self.env = self._build_env()
        self.opts = self._build_opts()

    def _build_opts(self):
        opts = []
        if self.account.get("user_name"):
            opts += ["-c", f"user.name={self.account['user_name']}"]
        if self.account.get("user_email"):
            opts += ["-c", f"user.email={self.account['user_email']}"]
        return opts

    def _build_env(self):
        env = dict(os.environ)
        # launchd などで認証プロンプト待ちのまま止まらないようにする
        env["GIT_TERMINAL_PROMPT"] = "0"

        acc = self.account
        auth = acc.get("auth")

        if auth == "ssh":
            key = acc.get("ssh_key")
            if not key:
                raise RuntimeError("auth=ssh には ssh_key が必要です")
            key_path = Path(os.path.expanduser(key)).resolve()
            if not key_path.exists():
                raise FileNotFoundError(f"SSH鍵が見つかりません: {key_path}")
            env["GIT_SSH_COMMAND"] = (
                f"ssh -i '{key_path}' -o IdentitiesOnly=yes -o BatchMode=yes"
            )

        elif auth == "token":
            token = load_secret(acc, "token", "トークン")
            host = acc.get("host", "github.com")
            basic = base64.b64encode(
                f"x-access-token:{token}".encode("utf-8")
            ).decode("ascii")
            pairs = [
                # キーチェーンに入っている別アカウントの資格情報を使わせない
                ("credential.helper", ""),
                (f"http.https://{host}/.extraheader", f"Authorization: Basic {basic}"),
            ]
            env["GIT_CONFIG_COUNT"] = str(len(pairs))
            for i, (k, v) in enumerate(pairs):
                env[f"GIT_CONFIG_KEY_{i}"] = k
                env[f"GIT_CONFIG_VALUE_{i}"] = v

        elif auth not in (None, "", "default"):
            raise RuntimeError(f"未対応の auth です: {auth}（ssh / token / default）")

        return env

    def cmd(self, *args):
        return ["git", *self.opts, "-C", str(self.repo_root), *args]

    def run(self, *args):
        result = subprocess.run(
            self.cmd(*args), capture_output=True, text=True, env=self.env
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"\nCMD : git {' '.join(args)}"
                f"\nOUT : {result.stdout}"
                f"\nERR : {mask(result.stderr)}"
            )
        return result

    def quiet(self, *args):
        return subprocess.run(
            self.cmd(*args), capture_output=True, text=True, env=self.env
        )


def mask(text):
    """エラーメッセージにトークンが混ざっても出さない。"""
    return text.replace("Authorization: Basic", "Authorization: Basic ***")


# ============================================================
# 秘密情報の読み込み
# ============================================================

def load_secret(conf, key, label):
    """<key>_env（環境変数名）→ <key>_file（ファイル）の順に探す。"""

    env_name = conf.get(f"{key}_env")
    if env_name and os.environ.get(env_name):
        return os.environ[env_name]

    file_name = conf.get(f"{key}_file")
    if file_name:
        path = Path(os.path.expanduser(file_name))
        if path.exists():
            text = path.read_text(encoding="utf-8").strip()
            if text:
                return text

    raise RuntimeError(
        f"{label}がありません。{key}_env（環境変数名）か "
        f"{key}_file（ファイルのパス）を確認してください"
    )


def resolve_account(config):
    """config["account"] を解決する。

    - 文字列: accounts ファイル内の名前を参照（秘密情報はリポジトリ外に置ける）
    - dict  : その場で定義（トークンは token_env / token_file で外出し推奨）
    - なし  : 従来どおり、マシンの既定の git 設定で動かす
    """

    acc = config.get("account")

    if acc is None:
        return None

    if isinstance(acc, dict):
        return acc

    accounts_file = Path(os.path.expanduser(
        config.get("accounts_file", DEFAULT_ACCOUNTS_FILE)
    ))
    if not accounts_file.exists():
        raise FileNotFoundError(
            f"アカウント定義ファイルが見つかりません: {accounts_file}"
        )

    accounts = json.loads(accounts_file.read_text(encoding="utf-8"))
    if acc not in accounts:
        raise KeyError(
            f"アカウント '{acc}' が {accounts_file} にありません"
            f"（定義済み: {', '.join(accounts.keys()) or 'なし'}）"
        )
    return accounts[acc]


# ============================================================
# git 補助
# ============================================================

def unmerged_files(g):
    result = g.quiet("diff", "--name-only", "--diff-filter=U")
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def abort_unfinished(g):
    """前回の実行が残したrebase/mergeを畳む。残っているとpullが弾かれる。"""

    git_dir = Path(g.run("rev-parse", "--absolute-git-dir").stdout.strip())

    if (git_dir / "rebase-merge").exists() or (git_dir / "rebase-apply").exists():
        print("前回のrebaseが途中のままだったため中断します")
        g.quiet("rebase", "--abort")

    elif (git_dir / "MERGE_HEAD").exists():
        print("前回のmergeが途中のままだったため中断します")
        g.quiet("merge", "--abort")


def ensure_remote(g, remote_name, remote_url):
    """remote_url 指定時は、remote がその URL を向いているようにする。"""

    if not remote_url:
        return

    current = g.quiet("remote", "get-url", remote_name)
    if current.returncode != 0:
        g.run("remote", "add", remote_name, remote_url)
        print(f"remote 追加 : {remote_name} -> {remote_url}")
    elif current.stdout.strip() != remote_url:
        g.run("remote", "set-url", remote_name, remote_url)
        print(f"remote 変更 : {remote_name} -> {remote_url}")


def has_staged_diff(g):
    return g.quiet("diff", "--cached", "--quiet").returncode != 0


def push_if_ahead(g, remote_name, branch):
    """working treeに差分が無くても、前回commitまで進んでpushだけ失敗した
    状態が残っていることがある。そのときは差分なしのまま黙って終わってしまうので、
    remoteより進んでいるcommitが無いか確認し、あればpushする。"""
    g.quiet("fetch", remote_name, branch)
    ahead = g.quiet("rev-list", f"{remote_name}/{branch}..HEAD", "--count")
    if ahead.returncode != 0 or ahead.stdout.strip() in ("", "0"):
        return
    print(f"ローカルに未pushのcommitが{ahead.stdout.strip()}件あるためpushします")
    g.run("push", remote_name, branch)
    print(f"push 完了 : {remote_name}/{branch}")


# ============================================================
# 1設定ぶんの同期
# ============================================================

REQUIRED_KEYS = ["id", "source_folder", "repository_root", "repo_subfolder", "branch",
                 "commit_message", "remote_name"]


def load_config(config_path):
    config = json.loads(Path(config_path).read_text(encoding="utf-8-sig"))
    missing = [k for k in REQUIRED_KEYS if k not in config]
    if missing:
        raise KeyError(
            f"設定ファイルに次のキーがありません: {', '.join(missing)}\n  ({config_path})"
        )
    return config


def sync_config(config_path, dry_run=False):

    config = json.loads(Path(config_path).read_text(encoding="utf-8-sig"))
    if not config.get("enabled", True):
        print(f"無効設定のためスキップ : {config_path}")
        return

    config = load_config(config_path)

    source_folder = Path(os.path.expanduser(config["source_folder"]))
    repo_root = Path(os.path.expanduser(config["repository_root"]))
    target_folder = repo_root / config["repo_subfolder"]

    account = resolve_account(config)
    g = GitRunner(repo_root, account)

    label = (account or {}).get("user_name") or "既定のgit設定"
    print(f"同期開始 : {config['id']}（アカウント: {label}）")

    if not source_folder.is_dir():
        raise FileNotFoundError(f"元フォルダが見つかりません: {source_folder}")

    abort_unfinished(g)
    ensure_remote(g, config["remote_name"], config.get("remote_url"))

    target_folder.mkdir(parents=True, exist_ok=True)

    #
    # ファイルコピー
    #
    copied = 0

    for source_file in source_folder.glob("*"):

        if not source_file.is_file():
            continue

        shutil.copy2(source_file, target_folder / source_file.name)
        copied += 1

    print(f"コピー : {copied}件")

    #
    # add
    #
    g.run("add", config["repo_subfolder"])

    #
    # 同期対象外の衝突
    #
    # 同期フォルダの衝突は上のコピーとaddで解消済み。それ以外は手で直すしかない
    #
    conflicts = unmerged_files(g)

    if conflicts:
        raise RuntimeError(
            "未解決の衝突が残っているため中止します:\n  "
            + "\n  ".join(conflicts)
        )

    #
    # 差分確認
    #
    if not has_staged_diff(g):
        print("差分なし")
        push_if_ahead(g, config["remote_name"], config["branch"])
        return

    if dry_run:
        changed = g.run("diff", "--cached", "--name-only").stdout.split()
        print("dry-run : commit/push しません。変更予定:\n  " + "\n  ".join(changed))
        return

    #
    # pull（rebase）
    #
    try:
        g.run("pull", "--rebase", "--autostash", config["remote_name"], config["branch"])

    except Exception as e:

        # 途中のまま残すと次回の実行もpullで弾かれる
        abort_unfinished(g)

        print(f"rebase失敗: {e}")
        raise

    # autostash復元でstagedが解けるのでadd し直す
    g.run("add", config["repo_subfolder"])

    # rebase後に差分が吸収される場合があるため、commit直前でも再確認する
    if not has_staged_diff(g):
        print("差分なし（pull後に同期済み）")
        push_if_ahead(g, config["remote_name"], config["branch"])
        return

    #
    # commit / push
    #
    g.run("commit", "-m", config["commit_message"])
    g.run("push", config["remote_name"], config["branch"])
    print(f"push 完了 : {config['remote_name']}/{config['branch']}")


# ============================================================
# main
# ============================================================

def expand_configs(paths):
    """--config にはファイルでもフォルダでも渡せる（フォルダは *.json を名前順）。"""
    result = []
    for p in paths:
        path = Path(os.path.expanduser(p))
        if path.is_dir():
            result += sorted(path.glob("*.json"))
        else:
            result.append(path)
    return result


def main():
    parser = argparse.ArgumentParser(description="github_sync")
    parser.add_argument("--config", required=True, action="append",
                        help="設定ファイル or フォルダ（複数指定可）")
    parser.add_argument("--dry-run", action="store_true",
                        help="add までで止め、commit/push しない")
    args = parser.parse_args()

    configs = expand_configs(args.config)
    errors = []

    # 1つの設定が失敗しても、他のアカウント・リポジトリの同期は続ける
    for config_path in configs:
        try:
            sync_config(config_path, dry_run=args.dry_run)
        except Exception as e:
            print(f"失敗 : {config_path}\n{mask(str(e))}", file=sys.stderr)
            errors.append(str(config_path))

    if errors:
        print(f"\n失敗した設定 {len(errors)}件:\n  " + "\n  ".join(errors), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

