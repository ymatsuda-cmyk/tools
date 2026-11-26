"""m2common.py — Mac 側の共通部品（設定・GitHub・ログ）

設定は ~/.minutes2/.env に置く（.env.example を参照）。
"""
import base64
import json
import os
import pathlib
import time

import requests

HOME = pathlib.Path(os.environ.get("MINUTES2_HOME", pathlib.Path.home() / ".minutes2"))
KEYS_PATH = HOME / "keys.json"


ENV_OVERRIDES: list = []   # 環境変数が、.env の値を（違う値で）上書きしたキー。気付かずに別のリポジトリなどを見る事故を、ログで知らせるため


def load_env() -> dict:
    env = {}
    p = HOME / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    for k, v in os.environ.items():
        if k.startswith(("GH_", "GAS_", "AUDIO_", "NOTION_", "DRIVE_", "WHISPER_", "MINUTES2_")):
            if k in env and env[k] != v:
                ENV_OVERRIDES.append(k)
            env[k] = v   # 環境変数が優先（launchd の EnvironmentVariables で上書きできるように）
    return env


ENV = load_env()

# ログの置き場所：既定は ~/scripts/logs（.env の MINUTES2_LOG_DIR で変えられる）
LOG_DIR = pathlib.Path(os.path.expanduser(ENV.get("MINUTES2_LOG_DIR") or "~/scripts/logs"))
LOG_PATH = LOG_DIR / "minutes2_worker.log"
LOG_MAX_BYTES = 5 * 1024 * 1024   # これを超えたら、古いログに回す（.1 .2 .3 を残す）
LOG_KEEP = 3


def _rotate_log():
    try:
        if LOG_PATH.exists() and LOG_PATH.stat().st_size > LOG_MAX_BYTES:
            for i in range(LOG_KEEP, 0, -1):
                src = LOG_PATH if i == 1 else LOG_PATH.with_name(f"{LOG_PATH.name}.{i - 1}")
                dst = LOG_PATH.with_name(f"{LOG_PATH.name}.{i}")
                if src.exists():
                    src.replace(dst)
    except OSError:
        pass


def log(msg: str):
    line = time.strftime("%Y-%m-%d %H:%M:%S ") + msg
    print(line, flush=True)
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        _rotate_log()
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def load_keys() -> dict:
    if not KEYS_PATH.exists():
        raise SystemExit(f"{KEYS_PATH} がありません。先に setup_keys.py init を実行してください。")
    return json.loads(KEYS_PATH.read_text(encoding="utf-8"))


def save_keys(keys: dict):
    HOME.mkdir(parents=True, exist_ok=True)
    tmp = KEYS_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(keys, ensure_ascii=False, indent=2), encoding="utf-8")
    os.chmod(tmp, 0o600)
    tmp.replace(KEYS_PATH)


class GitHub:
    """データ用リポジトリ（private）を contents API で読み書きする。index.json は sha で競合を検出する"""

    def __init__(self, env=None):
        env = env or ENV
        self.owner = env["GH_OWNER"]
        self.repo = env["GH_REPO"]
        self.branch = env.get("GH_BRANCH", "main")
        self.s = requests.Session()
        self.s.headers.update({
            "Authorization": "Bearer " + env["GH_TOKEN"],
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })

    def _url(self, path):
        return f"https://api.github.com/repos/{self.owner}/{self.repo}/contents/{path}"

    def check(self):
        """リポジトリに入れるかを確かめる。違っていると、ファイルが「無い」のと同じ見え方（404）になり、何もしないまま終わってしまうため"""
        r = self.s.get(f"https://api.github.com/repos/{self.owner}/{self.repo}", timeout=60)
        if r.status_code == 200:
            return
        if r.status_code in (401, 403):
            raise RuntimeError(f"GitHub に入れません（{r.status_code}）。GH_TOKEN が無効か、期限が切れている可能性があります。")
        if r.status_code == 404:
            raise RuntimeError(f"GitHub のリポジトリ {self.owner}/{self.repo} が見つかりません。GH_OWNER・GH_REPO と、"
                               "GH_TOKEN（fine-grained）の Resource owner と Repository access にこのリポジトリが入っているかを確かめてください。")
        r.raise_for_status()

    def get(self, path):
        """→ (text, sha) / 無ければ (None, None)"""
        r = self.s.get(self._url(path), params={"ref": self.branch}, timeout=60)
        if r.status_code == 404:
            return None, None
        r.raise_for_status()
        j = r.json()
        content = j.get("content")
        if not content or j.get("encoding") != "base64":
            b = self.s.get(f"https://api.github.com/repos/{self.owner}/{self.repo}/git/blobs/{j['sha']}", timeout=120)
            b.raise_for_status()
            content = b.json()["content"]
        return base64.b64decode(content).decode("utf-8"), j["sha"]

    def put(self, path, text, sha=None, message=None) -> bool:
        body = {
            "message": message or f"update {path}",
            "content": base64.b64encode(text.encode("utf-8")).decode("ascii"),
            "branch": self.branch,
        }
        if sha:
            body["sha"] = sha
        r = self.s.put(self._url(path), json=body, timeout=120)
        if r.status_code == 404 and self._is_empty_repo():
            # 空のリポジトリにはまだブランチが無いので、ブランチを指定せずに最初のコミットを作る
            body.pop("branch")
            r = self.s.put(self._url(path), json=body, timeout=120)
        if r.status_code in (409, 422):
            return False
        if r.status_code == 404:
            raise RuntimeError(
                f"GitHub に書けませんでした（404）: {self.owner}/{self.repo} が見つかりません。"
                "リポジトリ名と、書き込み用トークン（GH_TOKEN）の Repository access にこのリポジトリが入っているかを確かめてください。"
                f" GitHub の応答: {r.text[:200]}")
        if r.status_code in (401, 403):
            raise RuntimeError(
                f"GitHub に書けませんでした（{r.status_code}）: 書き込み用トークン（GH_TOKEN）の権限を確かめてください"
                "（Contents: Read and write が必要です）。"
                f" GitHub の応答: {r.text[:200]}")
        r.raise_for_status()
        return True

    def _is_empty_repo(self) -> bool:
        r = self.s.get(f"https://api.github.com/repos/{self.owner}/{self.repo}/branches", timeout=60)
        return r.status_code == 200 and r.json() == []

    def write(self, path, text, message=None):
        for _ in range(4):
            _, sha = self.get(path)
            if self.put(path, text, sha, message):
                return
            time.sleep(1)
        raise RuntimeError(f"{path} の書き込みが競合しました")

    def get_json(self, path, default=None):
        text, _ = self.get(path)
        return default if text is None else json.loads(text)

    def update_index(self, fn):
        """fn(list) で index.json を変える。GAS と同時に書いて sha がずれたら読み直してやり直す"""
        for i in range(5):
            text, sha = self.get("index.json")
            lst = json.loads(text) if text else []
            result = fn(lst)
            if self.put("index.json", json.dumps(lst, ensure_ascii=False, separators=(",", ":")), sha, "index (mac)"):
                return result
            time.sleep(1 + i)
        raise RuntimeError("index.json の更新が競合しました")
