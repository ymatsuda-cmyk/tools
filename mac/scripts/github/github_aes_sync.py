#!/usr/bin/env python3
"""
github_aes_sync.py

github_sync.py の暗号化版。source_folder のファイルを AES-256-GCM で暗号化して
repository_root/repo_subfolder に置き、commit / push する。

- 暗号化の形式は jsonbin/jsonbin_aes_sync.py と同じ（kanban-aesgcm/v1）。
  ブラウザ側は api/kanban/kanban.api.js の decrypt() をそのまま使える。
- 暗号文は salt/iv が毎回変わるため、差分確認は「リポジトリにある暗号ファイルを
  復号して、平文どうしで比べる」方式。変わっていないファイルは書き換えない。
- 設定ごとに GitHub アカウント（SSH鍵 / トークン / コミット者）を切り替えられる。
  アカウントの秘密情報はリポジトリ外の accounts ファイルに置く。

使い方:
  python3 github_aes_sync.py --config config/minutes-aes.json
  python3 github_aes_sync.py --config a.json --config b.json   # 複数
  python3 github_aes_sync.py --config config/                  # フォルダ内の *.json 全部
  python3 github_aes_sync.py --config a.json --dry-run         # commit/pushしない
  python3 github_aes_sync.py --config a.json --decrypt data/minutes/x.enc.json
"""

APP_VERSION = "rev_20261002_aes1"

import argparse
import base64
import fnmatch
import hashlib
import json
import os
import secrets
import subprocess
import sys
import tempfile
from pathlib import Path

# 暗号化の形式。api/kanban/kanban.api.js の decrypt() / jsonbin_aes_sync.py と同じにしておくこと
# （片方だけ変えるとブラウザで復号できなくなる）
ENC_NAME = "kanban-aesgcm/v1"
PBKDF2_ITER = 250000

# 既定のアカウント定義ファイル（リポジトリ外。トークン等はここにだけ置く）
DEFAULT_ACCOUNTS_FILE = "~/.config/github_aes_sync/accounts.json"

# 既定の暗号ファイル名の付け方: foo.json -> foo.enc.json / foo.csv -> foo.csv.enc.json
DEFAULT_ENC_SUFFIX = ".enc.json"


# ============================================================
# git 実行（アカウントごとの環境を差し込む）
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


def load_passphrase(config):
    """パスフレーズは設定ファイルに直接書かず、環境変数かファイルから読む。"""
    return load_secret(config, "passphrase", "暗号化用のパスフレーズ")


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
# 暗号化 / 復号（jsonbin_aes_sync.py と同じ形式）
# ============================================================

def _aesgcm():
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError as e:
        raise RuntimeError(
            "暗号化には cryptography が必要です: pip3 install cryptography"
        ) from e
    return AESGCM


def derive_key(passphrase, salt, iterations=PBKDF2_ITER):
    return hashlib.pbkdf2_hmac(
        "sha256", passphrase.encode("utf-8"), salt, iterations, dklen=32
    )


def encrypt_bytes(plain, passphrase, fmt):
    """plain を AES-256-GCM で暗号化し、封筒(JSON)の形にする。salt と iv は毎回作り直す。"""

    AESGCM = _aesgcm()
    salt = secrets.token_bytes(16)
    iv = secrets.token_bytes(12)
    key = derive_key(passphrase, salt)
    ct = AESGCM(key).encrypt(iv, plain, ENC_NAME.encode("utf-8"))

    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {
        "enc": ENC_NAME,
        "kdf": "PBKDF2-SHA256",
        "iter": PBKDF2_ITER,
        "salt": b64(salt),
        "iv": b64(iv),
        "ct": b64(ct),
        # json: 復号結果を JSON.parse すればよい / raw: 元ファイルのバイト列
        "format": fmt,
    }


def decrypt_envelope(envelope, passphrase):
    """封筒を復号して平文のバイト列を返す。形式違い・鍵違いは例外。"""

    if not isinstance(envelope, dict) or envelope.get("enc") != ENC_NAME:
        raise ValueError("暗号ファイルの形式が違います")

    from cryptography.exceptions import InvalidTag

    AESGCM = _aesgcm()
    salt = base64.b64decode(envelope["salt"])
    iv = base64.b64decode(envelope["iv"])
    ct = base64.b64decode(envelope["ct"])
    key = derive_key(passphrase, salt, int(envelope.get("iter", PBKDF2_ITER)))
    try:
        return AESGCM(key).decrypt(iv, ct, ENC_NAME.encode("utf-8"))
    except InvalidTag as e:
        raise ValueError("復号できません（パスフレーズ違い、またはファイル破損）") from e


# ============================================================
# 平文の用意
# ============================================================

def is_json_file(path):
    return path.suffix.lower() == ".json"


def build_plain(source_file):
    """暗号化する平文を作る。

    .json は一度パースしてから詰めて書き直す（jsonbin_aes_sync.py と同じ平文になる）。
    書きかけ・壊れた JSON はここで弾き、リポジトリに送らない。
    """

    data = source_file.read_bytes()

    if not is_json_file(source_file):
        return data, "raw"

    try:
        payload = json.loads(data.decode("utf-8-sig"))
        # Power Automate が JSON 文字列をそのまま書いた場合は、もう一度ほどく
        if isinstance(payload, str):
            payload = json.loads(payload)
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        raise RuntimeError(f"JSON として読めません（書き込み途中？）: {source_file.name}: {e}")

    plain = json.dumps(
        payload, ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    return plain, "json"


def encrypted_name(source_name, suffix):
    """foo.json -> foo.enc.json / foo.csv -> foo.csv.enc.json"""
    if suffix.endswith(".json") and source_name.lower().endswith(".json"):
        return source_name[: -len(".json")] + suffix
    return source_name + suffix


def write_atomic(path, text):
    """書きかけのファイルを commit しないよう、一時ファイル経由で置き換える。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".tmp_", suffix=path.suffix)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def iter_sources(source_folder, patterns):
    for source_file in sorted(source_folder.iterdir()):
        if not source_file.is_file():
            continue
        # .DS_Store / 一時ファイルなどは送らない
        if source_file.name.startswith(".") or source_file.name.startswith("~$"):
            continue
        if not any(fnmatch.fnmatch(source_file.name, p) for p in patterns):
            continue
        yield source_file


def encrypt_folder(config, source_folder, target_folder, passphrase):
    """source_folder を暗号化して target_folder に置く。中身が変わったファイルだけ書く。"""

    suffix = config.get("encrypted_suffix", DEFAULT_ENC_SUFFIX)
    patterns = config.get("include", ["*"])
    if isinstance(patterns, str):
        patterns = [patterns]

    written, unchanged, failed = [], 0, []

    for source_file in iter_sources(source_folder, patterns):
        target_file = target_folder / encrypted_name(source_file.name, suffix)

        try:
            plain, fmt = build_plain(source_file)
        except RuntimeError as e:
            # 1ファイル壊れていても他は進める（壊れたものは前回の暗号ファイルのまま）
            print(f"  スキップ : {e}")
            failed.append(source_file.name)
            continue

        # 既存の暗号ファイルを復号し、平文が同じなら書き換えない
        if target_file.exists():
            try:
                current = decrypt_envelope(
                    json.loads(target_file.read_text(encoding="utf-8")), passphrase
                )
                if current == plain:
                    unchanged += 1
                    continue
            except (ValueError, KeyError, json.JSONDecodeError) as e:
                print(f"  再暗号化 : {target_file.name}（{e}）")

        envelope = encrypt_bytes(plain, passphrase, fmt)
        write_atomic(target_file, json.dumps(envelope, ensure_ascii=False))
        written.append(target_file)

        # 平文のまま同じフォルダに残っていたら公開されたままなので知らせる
        plain_twin = target_folder / source_file.name
        if plain_twin.exists() and plain_twin != target_file:
            print(f"  注意 : 平文ファイルがリポジトリに残っています: {plain_twin}")

    print(f"暗号化 : 更新 {len(written)}件 / 変更なし {unchanged}件"
          + (f" / スキップ {len(failed)}件" if failed else ""))
    return written, failed


# ============================================================
# git 補助（github_sync.py と同じ流れ）
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


def discard_generated(g, subfolder, written):
    """今回書いた暗号ファイルを元に戻す。

    暗号文は毎回変わるので、別マシンが先に同じ内容を push していると
    pull の autostash 復元で必ず衝突する。暗号ファイルは元フォルダから
    作り直せるので、pull の前に捨てて、pull 後に平文比較で作り直す。
    """

    g.run("reset", "-q", "--", subfolder)
    for path in written:
        rel = str(path.relative_to(g.repo_root))
        tracked = g.quiet("ls-files", "--error-unmatch", "--", rel).returncode == 0
        if tracked:
            g.run("checkout", "HEAD", "--", rel)
        elif path.exists():
            path.unlink()


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
    if not (config.get("passphrase_env") or config.get("passphrase_file")):
        raise KeyError(
            f"passphrase_env か passphrase_file が必要です\n  ({config_path})"
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

    passphrase = load_passphrase(config)

    abort_unfinished(g)
    ensure_remote(g, config["remote_name"], config.get("remote_url"))

    #
    # 暗号化して配置（変わったものだけ）
    #
    written, _ = encrypt_folder(config, source_folder, target_folder, passphrase)

    #
    # add
    #
    g.run("add", config["repo_subfolder"])

    #
    # 同期対象外の衝突（手で直すしかない）
    #
    conflicts = unmerged_files(g)
    if conflicts:
        raise RuntimeError(
            "未解決の衝突が残っているため中止します:\n  " + "\n  ".join(conflicts)
        )

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
    discard_generated(g, config["repo_subfolder"], written)

    try:
        g.run("pull", "--rebase", "--autostash", config["remote_name"], config["branch"])
    except Exception as e:
        # 途中のまま残すと次回の実行もpullで弾かれる
        abort_unfinished(g)
        print(f"rebase失敗: {e}")
        raise

    # pull 後の内容に対して作り直す（同じ平文が push 済みなら書き換えない）
    encrypt_folder(config, source_folder, target_folder, passphrase)
    g.run("add", config["repo_subfolder"])

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
# 復号（確認用）
# ============================================================

def decrypt_file(config_path, enc_path):
    config = load_config(config_path)
    passphrase = load_passphrase(config)
    envelope = json.loads(Path(enc_path).read_text(encoding="utf-8"))
    plain = decrypt_envelope(envelope, passphrase)
    if envelope.get("format", "json") == "json":
        print(json.dumps(json.loads(plain.decode("utf-8")), ensure_ascii=False, indent=2))
    else:
        sys.stdout.buffer.write(plain)


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
    parser = argparse.ArgumentParser(description=f"github_aes_sync {APP_VERSION}")
    parser.add_argument("--config", required=True, action="append",
                        help="設定ファイル or フォルダ（複数指定可）")
    parser.add_argument("--dry-run", action="store_true",
                        help="暗号化と add までで止め、commit/push しない")
    parser.add_argument("--decrypt", metavar="ENC_FILE",
                        help="暗号ファイルを復号して表示する（確認用）")
    args = parser.parse_args()

    if args.decrypt:
        decrypt_file(args.config[0], args.decrypt)
        return

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
