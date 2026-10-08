# -*- coding: utf-8 -*-
"""ローカルアプリ起動の共通処理（dashrun ハンドラ / 将来の local_bridge で共用）

役割
  - exe パスの検証（ドライブ文字から始まるローカルの .exe のみ。ネットワークパス・引数は不可）
  - 許可済みリスト（%APPDATA%\\dashrun\\approved.json）の読み書き
  - 初回起動時の許可ダイアログ
  - exe の起動

戻り値は dashboard 側のランチャーと同じ形の dict:
  { "status": "started" | "denied" | "invalid" | "error", "message": str, "path": str }
local_bridge へ移行するときは handle_launch_request() をそのまま呼べばよい。
"""
import datetime
import json
import ntpath
import os
import re

APP_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "dashrun")
APPROVED_FILE = os.path.join(APP_DIR, "approved.json")
LOG_FILE = os.path.join(APP_DIR, "dashrun.log")

MAX_PATH_LEN = 1024
# ドライブ文字:\ で始まり、Windows で使えない文字を含まず、.exe で終わる
_EXE_PATH_RE = re.compile(r'^[A-Za-z]:\\(?:[^\\/:*?"<>|\r\n\x00]+\\)*[^\\/:*?"<>|\r\n\x00]+\.exe$', re.IGNORECASE)

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


# ---------------------------------------------------------------- 検証
def _resolve(path):
    """'..' を畳み、Windows ではシンボリックリンクも解決する。UNC になったら None"""
    p = ntpath.normpath(path)
    if IS_WINDOWS:
        p = os.path.realpath(p)
        if p.startswith("\\\\?\\") and not p.startswith("\\\\?\\UNC\\"):
            p = p[4:]
    if p.startswith("\\\\"):
        return None
    return p


def validate_exe_path(raw, exists=None):
    """(True, 正規化済みパス) または (False, 理由) を返す"""
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


# ---------------------------------------------------------------- 起動
def launch_exe(path):
    """引数なしで exe を起動する。ShellExecute 経由なので管理者権限が必要な exe も UAC が出る"""
    workdir = ntpath.dirname(path)
    if IS_WINDOWS:
        try:
            os.startfile(path, "open", cwd=workdir)  # Python 3.10+
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
