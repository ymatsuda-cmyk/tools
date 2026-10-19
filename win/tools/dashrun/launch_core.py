# -*- coding: utf-8 -*-
"""ローカルアプリ起動の共通処理（dashrun ハンドラ / 将来の local_bridge で共用）

役割
  - exe パスの検証（ドライブ文字から始まるローカルの .exe のみ。ネットワークパス・引数は不可）
  - 許可済みリスト（%APPDATA%\\dashrun\\approved.json）の読み書き
  - 初回起動時の許可ダイアログ
  - exe の起動
  - メールを指定したEdgeプロファイルで開く（openmail）
"""
import datetime
import json
import ntpath
import os
import re
import subprocess

APP_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "dashrun")
APPROVED_FILE = os.path.join(APP_DIR, "approved.json")
LOG_FILE = os.path.join(APP_DIR, "dashrun.log")
MAIL_PROFILES_FILE = os.path.join(APP_DIR, "mail_profiles.json")

MAX_PATH_LEN = 1024
_EXE_PATH_RE = re.compile(r'^[A-Za-z]:\\(?:[^\\/:*?"<>|\r\n\x00]+\\)*[^\\/:*?"<>|\r\n\x00]+\.exe$', re.IGNORECASE)
_PROFILE_RE = re.compile(r'^[A-Za-z0-9 _.-]{1,64}$')
_HTTPS_URL_RE = re.compile(r'^https://[^\s<>"\'\x00-\x1f]{1,2000}$')

IS_WINDOWS = os.name == "nt"


# ---------------------------------------------------------------- ログ
def log(message):
    try:
        os.makedirs(APP_DIR, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{stamp} {message}\n")
    except OSError:
        pass


# ---------------------------------------------------------------- 検証（exe）
def _resolve(path):
    p = ntpath.normpath(path)
    if IS_WINDOWS:
        p = os.path.realpath(p)
        if p.startswith("\\\\?\\") and not p.startswith("\\\\?\\UNC\\"):
            p = p[4:]
    if p.startswith("\\\\"):
        return None
    return p


def validate_exe_path(raw, exists=None):
    if not isinstance(raw, str):
        return False, "パスが文字列ではありません"
    path = raw.strip()
    if len(path) >= 2 and path[0] == path[-1] == '"':
        path = path[1:-1].strip()
    if not path:
        return False, "パスが空です"
    if len(path) > MAX_PATH_LEN:
        return False, "パスが長すぎます"
    if not _EXE_PATH_RE.match(path):
        return False, "ドライブ文字から始まる .exe のパスのみ起動できます（ネットワークパス・引数は不可）"
    resolved = _resolve(path)
    if not resolved or not _EXE_PATH_RE.match(resolved):
        return False, "ローカルドライブ上の .exe ではありません"
    check = exists if exists is not None else os.path.isfile
    if not check(resolved):
        return False, "ファイルが見つかりません"
    return True, resolved


def _key(path):
    return ntpath.normcase(path)


# ---------------------------------------------------------------- 検証（メール関連）
def validate_profile_name(raw):
    if not isinstance(raw, str) or not _PROFILE_RE.match(raw):
        return False, "プロファイル名が不正です"
    return True, raw


def validate_https_url(raw):
    if not isinstance(raw, str) or not _HTTPS_URL_RE.match(raw.strip()):
        return False, "https:// で始まる正しいURLのみ開けます"
    return True, raw.strip()


def load_mail_profiles():
    try:
        with open(MAIL_PROFILES_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


# ---------------------------------------------------------------- 許可済みリスト
def load_approved():
    try:
        with open(APPROVED_FILE, encoding="utf-8") as f:
            data = json.load(f)
        items = data.get("approved", []) if isinstance(data, dict) else []
        return [x for x in items if isinstance(x, dict) and isinstance(x.get("path"), str)]
    except (OSError, ValueError):
        return []


def save_approved(items):
    os.makedirs(APP_DIR, exist_ok=True)
    tmp = APPROVED_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"approved": items}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, APPROVED_FILE)


def is_approved(path):
    k = _key(path)
    return any(_key(x["path"]) == k for x in load_approved())


def approve(path):
    items = load_approved()
    if not any(_key(x["path"]) == _key(path) for x in items):
        items.append({"path": path, "approvedAt": datetime.datetime.now().isoformat(timespec="seconds")})
        save_approved(items)


# ---------------------------------------------------------------- ダイアログ（Windows）
MB_OK = 0x0
MB_YESNO = 0x4
MB_ICONERROR = 0x10
MB_ICONWARNING = 0x30
MB_DEFBUTTON2 = 0x100
MB_SETFOREGROUND = 0x10000
MB_TOPMOST = 0x40000
IDYES = 6


def message_box(text, title, flags):
    if not IS_WINDOWS:
        print(f"[{title}] {text}")
        return 0
    import ctypes
    return ctypes.windll.user32.MessageBoxW(None, text, title, flags | MB_SETFOREGROUND | MB_TOPMOST)


def ask_approval(path):
    text = (
        "ダッシュボードから次のアプリの起動が要求されました。\n\n"
        f"{path}\n\n"
        "このアプリの起動を許可しますか？\n"
        "（許可すると次回からは確認なしで起動します）"
    )
    return message_box(text, "dashrun - 起動の許可", MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) == IDYES


def show_error(text):
    message_box(text, "dashrun", MB_OK | MB_ICONERROR)


# ---------------------------------------------------------------- 起動（exe・引数なし）
def launch_exe(path):
    workdir = ntpath.dirname(path)
    if IS_WINDOWS:
        try:
            os.startfile(path, "open", cwd=workdir)
        except TypeError:
            os.startfile(path)
    else:
        raise OSError("Windows 以外では起動できません")


def handle_launch_request(raw_path, ask=ask_approval, launcher=launch_exe, exists=None):
    ok, result = validate_exe_path(raw_path, exists=exists)
    if not ok:
        log(f"INVALID {raw_path!r}: {result}")
        return {"status": "invalid", "message": result, "path": str(raw_path)}
    path = result
    if not is_approved(path):
        if not ask(path):
            log(f"DENIED {path}")
            return {"status": "denied", "message": "起動は許可されませんでした", "path": path}
        approve(path)
        log(f"APPROVED {path}")
    try:
        launcher(path)
    except OSError as e:
        log(f"ERROR {path}: {e}")
        return {"status": "error", "message": f"起動に失敗しました: {e}", "path": path}
    log(f"STARTED {path}")
    return {"status": "started", "message": "起動しました", "path": path}


# ---------------------------------------------------------------- 起動（メール：指定プロファイルで開く）
def launch_url_in_profile(exe_path, profile, url):
    # outlook.office.com はドメイン移行中で不安定なため、新ドメインに固定してから開く
    fixed_url = url.replace("https://outlook.office.com/", "https://outlook.cloud.microsoft/", 1)
    workdir = ntpath.dirname(exe_path)
    args = [exe_path]
    if profile:
        args.append(f"--profile-directory={profile}")
    args.append(fixed_url)
    subprocess.Popen(args, cwd=workdir or None)


def handle_openmail_request(account_id, raw_url, ask=ask_approval):
    entry = load_mail_profiles().get(account_id)
    if not isinstance(entry, dict):
        log(f"NOPROFILE {account_id!r}")
        return {"status": "invalid", "message": f"mail_profiles.json に {account_id} の設定がありません"}
    return handle_openurl_request(entry.get("exe", ""), entry.get("profile", ""), raw_url, ask=ask)


def handle_openurl_request(raw_exe, raw_profile, raw_url, ask=ask_approval):
    """指定したブラウザ exe（とプロファイル）で URL を開く。profile が空ならプロファイル指定なし"""
    ok, exe = validate_exe_path(raw_exe)
    if not ok:
        return {"status": "invalid", "message": exe}
    profile = ""
    if raw_profile:
        ok, profile = validate_profile_name(raw_profile)
        if not ok:
            return {"status": "invalid", "message": profile}
    ok, url = validate_https_url(raw_url)
    if not ok:
        return {"status": "invalid", "message": url}

    mkey = f"{ntpath.normcase(exe)}|{profile.lower()}"
    if not any(_key(x.get("path", "")) == mkey for x in load_approved()):
        if not ask(f"{exe}\n（プロファイル: {profile or '指定なし'}）"):
            log(f"DENIED {mkey}")
            return {"status": "denied", "message": "起動は許可されませんでした"}
        items = load_approved()
        items.append({"path": mkey, "approvedAt": datetime.datetime.now().isoformat(timespec="seconds")})
        save_approved(items)
        log(f"APPROVED {mkey}")

    try:
        launch_url_in_profile(exe, profile, url)
    except OSError as e:
        log(f"ERROR {mkey}: {e}")
        return {"status": "error", "message": f"起動に失敗しました: {e}"}
    log(f"STARTED {mkey} -> {url}")
    return {"status": "started", "message": "開きました"}