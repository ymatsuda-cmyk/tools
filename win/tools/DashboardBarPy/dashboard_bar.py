"""
DashboardBar（Python版）
ダッシュボードを Edge のアプリモードで開き、画面の端に AppBar（サイドバー）として固定する。

操作:
  つまみクリック: 展開/収納   つまみ上下ドラッグ: 位置移動   帯左右ドラッグ: 幅変更
  帯ホバー: 覗き見   Ctrl+Alt+D: 切替   右クリック / タスクトレイ: メニュー
  exe をつまみ・帯にドロップ: ダッシュボードの追加画面を開く

標準ライブラリのみ（ctypes + tkinter + winreg）。pythonw.exe で起動する想定。
"""
VERSION = "2.0.1"

import collections
import ctypes
import faulthandler
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.parse
import urllib.request
import winreg
from ctypes import wintypes as wt
from tkinter import messagebox, ttk

APP_NAME = "DashboardBar"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.abspath(__file__)
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")
LOG_PATH = os.path.join(BASE_DIR, "dashboard_bar.log")
ICON_PATH = os.path.join(BASE_DIR, "dashboard_bar.ico")
INSTANCE_PATH = os.path.join(BASE_DIR, ".instance")
STRIP_TITLE = "DashboardBar Strip"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

logging.basicConfig(filename=LOG_PATH, level=logging.INFO, encoding="utf-8",
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("bar")

# Python 自体が異常終了したときも原因をログに残す
try:
    _fault_log = open(LOG_PATH, "a", encoding="utf-8")
    faulthandler.enable(file=_fault_log)
except Exception:
    pass

# ---------------- DPI（物理ピクセルで扱う） ----------------
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

# ---------------- Win32 定義 ----------------
user32 = ctypes.WinDLL("user32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
dwmapi = ctypes.WinDLL("dwmapi")
gdi32 = ctypes.WinDLL("gdi32")

LRESULT = ctypes.c_ssize_t


def _fn(dll, name, res, *args):
    f = getattr(dll, name)
    f.restype = res
    f.argtypes = list(args)
    return f


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [("cbSize", wt.DWORD), ("rcMonitor", RECT), ("rcWork", RECT),
                ("dwFlags", wt.DWORD), ("szDevice", wt.WCHAR * 32)]


class APPBARDATA(ctypes.Structure):
    _fields_ = [("cbSize", wt.DWORD), ("hWnd", wt.HWND), ("uCallbackMessage", wt.UINT),
                ("uEdge", wt.UINT), ("rc", RECT), ("lParam", wt.LPARAM)]


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wt.DWORD), ("Data2", wt.WORD), ("Data3", wt.WORD), ("Data4", wt.BYTE * 8)]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [("cbSize", wt.DWORD), ("hWnd", wt.HWND), ("uID", wt.UINT), ("uFlags", wt.UINT),
                ("uCallbackMessage", wt.UINT), ("hIcon", wt.HICON), ("szTip", wt.WCHAR * 128),
                ("dwState", wt.DWORD), ("dwStateMask", wt.DWORD), ("szInfo", wt.WCHAR * 256),
                ("uVersion", wt.UINT), ("szInfoTitle", wt.WCHAR * 64), ("dwInfoFlags", wt.DWORD),
                ("guidItem", GUID), ("hBalloonIcon", wt.HICON)]


WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
MONITORENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HMONITOR, wt.HDC, ctypes.POINTER(RECT), wt.LPARAM)
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)

SHAppBarMessage = _fn(shell32, "SHAppBarMessage", ctypes.c_size_t, wt.DWORD, ctypes.POINTER(APPBARDATA))
Shell_NotifyIconW = _fn(shell32, "Shell_NotifyIconW", wt.BOOL, wt.DWORD, ctypes.POINTER(NOTIFYICONDATAW))
DragAcceptFiles = _fn(shell32, "DragAcceptFiles", None, wt.HWND, wt.BOOL)
DragQueryFileW = _fn(shell32, "DragQueryFileW", wt.UINT, wt.HANDLE, wt.UINT, wt.LPWSTR, wt.UINT)
DragFinish = _fn(shell32, "DragFinish", None, wt.HANDLE)

EnumWindows = _fn(user32, "EnumWindows", wt.BOOL, WNDENUMPROC, wt.LPARAM)
EnumDisplayMonitors = _fn(user32, "EnumDisplayMonitors", wt.BOOL, wt.HDC, ctypes.POINTER(RECT),
                          MONITORENUMPROC, wt.LPARAM)
GetMonitorInfoW = _fn(user32, "GetMonitorInfoW", wt.BOOL, wt.HMONITOR, ctypes.POINTER(MONITORINFOEXW))
GetClassNameW = _fn(user32, "GetClassNameW", ctypes.c_int, wt.HWND, wt.LPWSTR, ctypes.c_int)
GetWindowTextLengthW = _fn(user32, "GetWindowTextLengthW", ctypes.c_int, wt.HWND)
IsWindowVisible = _fn(user32, "IsWindowVisible", wt.BOOL, wt.HWND)
IsWindow = _fn(user32, "IsWindow", wt.BOOL, wt.HWND)
SetWindowPos = _fn(user32, "SetWindowPos", wt.BOOL, wt.HWND, wt.HWND,
                   ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.UINT)
ShowWindow = _fn(user32, "ShowWindow", wt.BOOL, wt.HWND, ctypes.c_int)
GetCursorPos = _fn(user32, "GetCursorPos", wt.BOOL, ctypes.POINTER(wt.POINT))
GetWindowRect = _fn(user32, "GetWindowRect", wt.BOOL, wt.HWND, ctypes.POINTER(RECT))
GetAncestor = _fn(user32, "GetAncestor", wt.HWND, wt.HWND, wt.UINT)
GetForegroundWindow = _fn(user32, "GetForegroundWindow", wt.HWND)
SetForegroundWindow = _fn(user32, "SetForegroundWindow", wt.BOOL, wt.HWND)
GetAsyncKeyState = _fn(user32, "GetAsyncKeyState", ctypes.c_short, ctypes.c_int)
RegisterWindowMessageW = _fn(user32, "RegisterWindowMessageW", wt.UINT, wt.LPCWSTR)
PostMessageW = _fn(user32, "PostMessageW", wt.BOOL, wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)
FindWindowW = _fn(user32, "FindWindowW", wt.HWND, wt.LPCWSTR, wt.LPCWSTR)
SetWindowLongPtrW = _fn(user32, "SetWindowLongPtrW", ctypes.c_void_p, wt.HWND, ctypes.c_int, ctypes.c_void_p)
CallWindowProcW = _fn(user32, "CallWindowProcW", LRESULT, ctypes.c_void_p, wt.HWND, wt.UINT,
                      wt.WPARAM, wt.LPARAM)
LoadImageW = _fn(user32, "LoadImageW", wt.HANDLE, wt.HINSTANCE, wt.LPCWSTR, wt.UINT,
                 ctypes.c_int, ctypes.c_int, wt.UINT)
GetSystemMetrics = _fn(user32, "GetSystemMetrics", ctypes.c_int, ctypes.c_int)
SetWindowRgn = _fn(user32, "SetWindowRgn", ctypes.c_int, wt.HWND, wt.HRGN, wt.BOOL)
CreatePolygonRgn = _fn(gdi32, "CreatePolygonRgn", wt.HRGN, ctypes.POINTER(wt.POINT), ctypes.c_int, ctypes.c_int)
DwmGetWindowAttribute = _fn(dwmapi, "DwmGetWindowAttribute", ctypes.c_long,
                            wt.HWND, wt.DWORD, ctypes.c_void_p, wt.DWORD)
CreateMutexW = _fn(kernel32, "CreateMutexW", wt.HANDLE, ctypes.c_void_p, wt.BOOL, wt.LPCWSTR)
CloseHandle = _fn(kernel32, "CloseHandle", wt.BOOL, wt.HANDLE)

try:
    GetDpiForSystem = _fn(user32, "GetDpiForSystem", wt.UINT)
except AttributeError:
    GetDpiForSystem = lambda: 96  # noqa: E731

ABM_NEW, ABM_REMOVE, ABM_QUERYPOS, ABM_SETPOS = 0, 1, 2, 3
ABE_LEFT, ABE_RIGHT = 0, 2
ABN_POSCHANGED = 1
SWP_NOACTIVATE = 0x10
SW_HIDE, SW_SHOWNOACTIVATE = 0, 4
HWND_TOPMOST = wt.HWND(-1)
WM_CLOSE = 0x0010
WM_DROPFILES = 0x0233
WM_LBUTTONUP, WM_RBUTTONUP, WM_CONTEXTMENU = 0x0202, 0x0205, 0x007B
WM_APP = 0x8000
WM_APP_ACTIVATE = WM_APP + 1      # 二重起動時に既存側へ送る
WM_APP_TRAY = WM_APP + 2          # タスクトレイからの通知
NIN_BALLOONUSERCLICK = 0x0405
NIM_ADD, NIM_MODIFY, NIM_DELETE = 0, 1, 2
NIF_MESSAGE, NIF_ICON, NIF_TIP, NIF_INFO = 0x1, 0x2, 0x4, 0x10
NIIF_INFO = 0x1
GA_ROOT = 2
GWLP_WNDPROC = -4
DWMWA_EXTENDED_FRAME_BOUNDS = 9
ERROR_ALREADY_EXISTS = 183
IMAGE_ICON, LR_LOADFROMFILE = 1, 0x10
SM_CXSMICON = 49
VK_CONTROL, VK_MENU, VK_D = 0x11, 0x12, 0x44
SHELL_CLASSES = {"Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"}

EXPANDED, COLLAPSED, PEEK = "expanded", "collapsed", "peek"

# ---------------- 設定 ----------------
DEFAULTS = {
    "Url": "https://ymatsuda-cmyk.github.io/tools/dashboard/",
    "Browser": "",               # 空なら Edge を自動検出
    "Monitor": "",               # 表示するモニター（"\\\\.\\DISPLAY2" など）。空ならメイン
    "Side": "left",              # left / right
    "Width": 380,
    "MinWidth": 240,
    "MaxWidth": 800,
    "StripWidth": 6,
    "TabWidth": 20,
    "TabHeight": 68,
    "TabPosition": 0.5,
    "TabOpacity": 0.55,
    "StripColor": "#C9C5BC",
    "TabColor": "#D3D1C7",
    "ArrowColor": "#B4B2A9",
    "AnimationMs": 260,
    "HoverPeek": True,
    "HideOnFullscreen": True,    # 全画面アプリ中は隠す
    "AutoRecover": True,         # ダッシュボードが閉じたら開き直す
    "RestoreState": True,        # 前回の展開/収納を復元
    "StartCollapsed": False,     # RestoreState が false のときの起動状態
    "LastState": "expanded",
    "CheckUpdateOnStart": True,
    "UpdateUrl": "https://raw.githubusercontent.com/ymatsuda-cmyk/tools/main/dashboardbar/dashboard_bar.py",
    "HideTitleBar": True,
    "TitleBarHeight": 32,
}


def load_settings():
    s = dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            s.update(json.load(f))
    except FileNotFoundError:
        pass
    except Exception as e:
        log.warning("settings.json の読込に失敗: %s", e)
    s.pop("Frameless", None)
    s.pop("TabHoverColor", None)
    s["Width"] = max(s["MinWidth"], min(s["MaxWidth"], int(s["Width"])))
    if s["Side"] not in ("left", "right"):
        s["Side"] = "left"
    return s


def save_settings(s):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning("settings.json の保存に失敗: %s", e)


# ---------------- 小物 ----------------
def with_params(url, **params):
    parts = urllib.parse.urlsplit(url)
    q = [(k, v) for k, v in urllib.parse.parse_qsl(parts.query, keep_blank_values=True) if k not in params]
    q += [(k, v) for k, v in params.items()]
    return urllib.parse.urlunsplit(parts._replace(query=urllib.parse.urlencode(q)))


def sidebar_url(url, **extra):
    """ダッシュボード側がサイドバー表示と判定できるよう ?sidebar=1 を付ける"""
    return with_params(url, sidebar="1", **extra)


def find_browser(s):
    if s.get("Browser") and os.path.exists(s["Browser"]):
        return s["Browser"]
    for r in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if r:
            p = os.path.join(r, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.exists(p):
                return p
    return None


def class_name(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def list_browser_windows():
    found = []

    def cb(h, _):
        if h and IsWindowVisible(h) and class_name(h) == "Chrome_WidgetWin_1" and GetWindowTextLengthW(h) > 0:
            found.append(h)
        return True

    EnumWindows(WNDENUMPROC(cb), 0)
    return found


def launch_app_window(browser, url, timeout=20.0):
    before = set(list_browser_windows())
    subprocess.Popen([browser, f"--app={url}"])
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(0.25)
        new = [h for h in list_browser_windows() if h not in before]
        if new:
            time.sleep(0.5)  # Edge 側の初期配置が終わるのを待つ
            return new[0]
    return None


def list_monitors():
    """メインを先頭に、左から順に並べたモニター一覧"""
    mons = []

    def cb(hmon, _hdc, _rc, _lp):
        mi = MONITORINFOEXW()
        mi.cbSize = ctypes.sizeof(MONITORINFOEXW)
        if GetMonitorInfoW(hmon, ctypes.byref(mi)):
            mons.append({
                "device": mi.szDevice,
                "rect": RECT(mi.rcMonitor.left, mi.rcMonitor.top, mi.rcMonitor.right, mi.rcMonitor.bottom),
                "work": RECT(mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom),
                "primary": bool(mi.dwFlags & 1),
            })
        return True

    EnumDisplayMonitors(None, None, MONITORENUMPROC(cb), 0)
    mons.sort(key=lambda m: (not m["primary"], m["rect"].left, m["rect"].top))
    return mons


def monitor_label(i, m):
    r = m["rect"]
    return f"{i + 1}: {m['device'].replace(chr(92) * 2 + '.' + chr(92), '')}  {r.right - r.left}×{r.bottom - r.top}" + \
        ("（メイン）" if m["primary"] else "")


def rect_of(hwnd):
    r = RECT()
    GetWindowRect(hwnd, ctypes.byref(r))
    return r


def frame_insets(hwnd):
    """Windows 11 の見えない枠（リサイズ用の透明な境界）の幅: (左, 上, 右, 下)"""
    win = rect_of(hwnd)
    vis = RECT()
    hr = DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(vis), ctypes.sizeof(RECT))
    if hr != 0 or vis.right <= vis.left:
        return None
    return (vis.left - win.left, vis.top - win.top, win.right - vis.right, win.bottom - vis.bottom)


def _bezier(p0, p1, p2, p3, n=10):
    pts = []
    for i in range(1, n + 1):
        t = i / n
        u = 1 - t
        pts.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                    u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return pts


def tab_outline(w, h, mirror=False):
    """角を丸めた台形（帯側が高く、外側が低い）の輪郭。mirror=True で左右反転（右端配置用）"""
    sx = w / 20.0
    c = min(20.0, h / 2.0)
    k = c / 20.0

    def P(x, y):
        return (x * sx, y)

    pts = [P(0, 0)]
    pts += _bezier(P(0, 0), P(4, 0), P(8, 3 * k), P(13, 7 * k))
    pts += _bezier(P(13, 7 * k), P(17, 10 * k), P(20, 14 * k), P(20, c))
    pts.append(P(20, h - c))
    pts += _bezier(P(20, h - c), P(20, h - 14 * k), P(17, h - 10 * k), P(13, h - 7 * k))
    pts += _bezier(P(13, h - 7 * k), P(8, h - 3 * k), P(4, h), P(0, h))
    if mirror:
        pts = [(w - x, y) for x, y in pts]
    return pts


def inside(p, r):
    return r.left <= p.x < r.right and r.top <= p.y < r.bottom


def parse_version(text):
    m = re.search(r'^VERSION\s*=\s*"([0-9.]+)"', text, re.M)
    return m.group(1) if m else None


def version_tuple(v):
    return tuple(int(x) for x in v.split(".") if x.isdigit())


def pythonw_path():
    exe = sys.executable
    if os.path.basename(exe).lower() == "python.exe":
        cand = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.exists(cand):
            return cand
    return exe


def autostart_command():
    return f'"{pythonw_path()}" "{SCRIPT_PATH}"'


def get_autostart():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except OSError:
        return False


def set_autostart(on):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if on:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, autostart_command())
        else:
            try:
                winreg.DeleteValue(k, APP_NAME)
            except FileNotFoundError:
                pass


def ensure_icon():
    """トレイ用アイコン（32x32）を生成。サイドバーを模した絵柄"""
    if os.path.exists(ICON_PATH):
        return ICON_PATH
    import struct
    n = 32
    px = bytearray()
    for y in range(n - 1, -1, -1):           # BMP は下から上
        for x in range(n):
            def rounded(x0, y0, x1, y1, r):
                if not (x0 <= x < x1 and y0 <= y < y1):
                    return False
                cx = min(max(x, x0 + r), x1 - 1 - r)
                cy = min(max(y, y0 + r), y1 - 1 - r)
                return (x - cx) ** 2 + (y - cy) ** 2 <= r * r + 0.5
            if rounded(1, 1, 31, 31, 6):
                if rounded(5, 6, 14, 26, 2):
                    b, g, r_, a = 0xC7, 0xD1, 0xD3, 255     # パネル（明るいグレー）
                elif 14 <= x < 16 and 12 <= y < 20:
                    b, g, r_, a = 0xC7, 0xD1, 0xD3, 255     # つまみ
                else:
                    b, g, r_, a = 0x5E, 0x6F, 0x2F, 255     # 背景（深緑）
            else:
                b, g, r_, a = 0, 0, 0, 0
            px += bytes((b, g, r_, a))
    mask = bytes(n * 4)
    bih = struct.pack("<IiiHHIIiiII", 40, n, n * 2, 1, 32, 0, len(px) + len(mask), 0, 0, 0, 0)
    data = bih + bytes(px) + mask
    header = struct.pack("<HHH", 0, 1, 1) + struct.pack("<BBBBHHII", n, n, 0, 0, 1, 32, len(data), 22)
    with open(ICON_PATH, "wb") as f:
        f.write(header + data)
    return ICON_PATH


# ---------------- 本体 ----------------
class DashboardBar:
    def __init__(self, mutex):
        self.mutex = mutex
        # Win32 のコールバックや別スレッドからは tkinter を直接触らず、ここに積んで fast_poll で実行する
        self.events = collections.deque()
        self.s = load_settings()
        self.scale = GetDpiForSystem() / 96.0

        self.state = COLLAPSED
        self.visible = 0.0          # 見えている中身の幅（DIP）
        self.reserved = None        # 予約中の幅（DIP）
        self.anchor = 0             # 予約領域の外側の端（左配置なら左端、右配置なら右端）
        self.top = 0
        self.bottom = 0
        self.mon = None
        self.mon_key = None
        self.anim_id = None
        self.insets = (0, 0, 0, 0)
        self.expected = None
        self.fix_ids = []
        self.registered = False
        self.closing = False
        self.navigating = False
        self.fullscreen_hidden = False
        self.recover_times = []
        self.settings_win = None
        self.pending_update = None  # (version, source)
        self.tray_added = False

        self.down = False
        self.dragging = False
        self.drag_x = 0
        self.drag_w = 0
        self.tab_down = False
        self.tab_moved = False
        self.tab_drag_y = 0
        self.tab_drag_pos = 0.5
        self.hover_since = None
        self.out_since = None
        self.hotkey_prev = False
        self.poll_count = 0

        # ---- 帯（tkinter） ----
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title(STRIP_TITLE)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.s["StripColor"], cursor="sb_h_double_arrow")
        self.root.report_callback_exception = self._on_tk_error

        # ---- つまみ ----
        self.tab = tk.Toplevel(self.root)
        self.tab.withdraw()
        self.tab.overrideredirect(True)
        self.tab.attributes("-topmost", True)
        self.tab.attributes("-alpha", float(self.s["TabOpacity"]))
        self.tab.configure(bg=self.s["TabColor"], cursor="hand2")
        self.tab_label = tk.Label(self.tab, text="◀", bg=self.s["TabColor"], fg=self.s["ArrowColor"],
                                  font=("Segoe UI", 8), cursor="hand2")
        self.tab_label.pack(expand=True, fill="both")
        self.tab.bind("<Enter>", lambda e: self._tab_hover(True))
        self.tab.bind("<Leave>", lambda e: self._tab_hover(False))

        # ---- メニュー（右クリック・タスクトレイ共通） ----
        self.menu = tk.Menu(self.root, tearoff=0, postcommand=self._update_menu)
        self.menu.add_command(label="収納", command=self.menu_toggle)
        self.menu.add_command(label="再読み込み", command=self.reload)
        self.menu.add_command(label="設定…", command=self.open_settings)
        self.menu.add_command(label="更新を確認", command=lambda: self.check_update(manual=True))
        self.menu.add_separator()
        self.menu.add_command(label="終了", command=self.quit)

        self.root.bind("<ButtonPress-1>", self.on_press)
        self.root.bind("<B1-Motion>", self.on_motion)
        self.root.bind("<ButtonRelease-1>", self.on_release)
        self.tab.bind("<ButtonPress-1>", self.on_tab_press)
        self.tab.bind("<B1-Motion>", self.on_tab_motion)
        self.tab.bind("<ButtonRelease-1>", self.on_tab_release)
        for w in (self.root, self.tab):
            w.bind("<Button-3>", lambda e: self.show_menu(e.x_root, e.y_root))

        # ---- Edge アプリウィンドウ ----
        self.browser = find_browser(self.s)
        if not self.browser:
            self._fatal("Edge（msedge.exe）が見つかりません。settings.json の Browser にパスを指定してください。")
        self.edge = launch_app_window(self.browser, sidebar_url(self.s["Url"]))
        if not self.edge:
            self._fatal("ダッシュボードのウィンドウを見つけられませんでした。")
        log.info("start v%s, Edge window: %s", VERSION, self.edge)
        self.update_insets()

        # ---- 帯とつまみを表示して各種登録 ----
        self.root.deiconify()
        self.tab.deiconify()
        self.root.update_idletasks()
        self.strip_hwnd = GetAncestor(self.root.winfo_id(), GA_ROOT)
        self.tab_hwnd = GetAncestor(self.tab.winfo_id(), GA_ROOT)
        self._write_instance()

        self.appbar_msg = RegisterWindowMessageW("DashboardBarPy_AppBarMsg")
        self.taskbar_created_msg = RegisterWindowMessageW("TaskbarCreated")
        self._old_procs = {}
        self._wndproc_ref = WNDPROC(self._wndproc)   # GC されないよう保持
        for h in (self.strip_hwnd, self.tab_hwnd):
            self._old_procs[int(h)] = SetWindowLongPtrW(h, GWLP_WNDPROC,
                                                        ctypes.cast(self._wndproc_ref, ctypes.c_void_p))
            DragAcceptFiles(h, True)

        self.appbar_register()
        self.tray_add()
        self._shape_tab()
        self._refresh_monitor()
        log.info("init: appbar/tray ready (tray=%s, monitor=%s)", self.tray_added, self.mon["device"])

        start_expanded = (self.s["LastState"] == EXPANDED) if self.s["RestoreState"] else not self.s["StartCollapsed"]
        if start_expanded:
            self.state = EXPANDED
            self.reserve(self.W + self.S)
            self.visible = self.W
            self.layout()
        else:
            self.state = COLLAPSED
            self.reserve(self.S)
            self.visible = 0
            self.layout()
            ShowWindow(self.edge, SW_HIDE)
        self.update_arrow()

        self.root.after(50, self.fast_poll)
        self.root.after(1000, self.slow_poll)
        log.info("init: done (%s)", self.state)
        if self.s["CheckUpdateOnStart"]:
            self.root.after(8000, lambda: self.check_update(manual=False))

    # ================= 便利プロパティ =================
    @property
    def W(self):
        return int(self.s["Width"])

    @property
    def S(self):
        return int(self.s["StripWidth"])

    @property
    def left(self):
        return self.s["Side"] != "right"

    def px(self, dip):
        return int(round(dip * self.scale))

    def _abd(self):
        abd = APPBARDATA()
        abd.cbSize = ctypes.sizeof(APPBARDATA)
        abd.hWnd = self.strip_hwnd
        return abd

    # ================= Win32 メッセージ（帯・つまみのウィンドウをサブクラス化） =================
    def _wndproc(self, hwnd, msg, wp, lp):
        try:
            if msg == WM_APP_TRAY:
                ev = lp & 0xFFFF
                if ev == WM_LBUTTONUP:
                    self.events.append(self.menu_toggle)
                elif ev in (WM_RBUTTONUP, WM_CONTEXTMENU):
                    self.events.append(self.show_menu_at_cursor)
                elif ev == NIN_BALLOONUSERCLICK and self.pending_update:
                    self.events.append(self.prompt_update)
                return 0
            if msg == WM_DROPFILES:
                files = []
                hdrop = wt.HANDLE(wp)
                count = DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
                for i in range(count):
                    buf = ctypes.create_unicode_buffer(1024)
                    DragQueryFileW(hdrop, i, buf, 1024)
                    files.append(buf.value)
                DragFinish(hdrop)
                self.events.append(lambda: self.on_drop(files))
                return 0
            if msg == WM_APP_ACTIVATE:
                self.events.append(self.on_second_instance)
                return 0
            if msg == self.appbar_msg:
                if wp == ABN_POSCHANGED:
                    self.events.append(self.relayout)
                return 0
            if msg == self.taskbar_created_msg:
                self.events.append(self.on_taskbar_created)
        except Exception:
            log.exception("wndproc")
        return CallWindowProcW(self._old_procs.get(int(hwnd)), hwnd, msg, wp, lp)

    def _write_instance(self):
        try:
            with open(INSTANCE_PATH, "w") as f:
                f.write(str(int(self.strip_hwnd)))
        except Exception:
            pass

    # ================= タスクトレイ =================
    def _nid(self):
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.strip_hwnd
        nid.uID = 1
        return nid

    def tray_add(self):
        try:
            size = GetSystemMetrics(SM_CXSMICON) or 16
            self.hicon = LoadImageW(None, ensure_icon(), IMAGE_ICON, size, size, LR_LOADFROMFILE)
            nid = self._nid()
            nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
            nid.uCallbackMessage = WM_APP_TRAY
            nid.hIcon = self.hicon
            nid.szTip = f"{APP_NAME} v{VERSION}"
            self.tray_added = bool(Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)))
        except Exception:
            log.exception("tray_add")

    def tray_remove(self):
        if self.tray_added:
            Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self._nid()))
            self.tray_added = False

    def notify(self, title, text):
        """タスクトレイから Windows の通知を出す"""
        log.info("notify: %s / %s", title, text)
        if not self.tray_added:
            return
        nid = self._nid()
        nid.uFlags = NIF_INFO
        nid.szInfoTitle = title[:63]
        nid.szInfo = text[:255]
        nid.dwInfoFlags = NIIF_INFO
        Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid))

    def on_taskbar_created(self):
        """エクスプローラーが再起動したら、トレイと AppBar を登録し直す"""
        log.info("TaskbarCreated")
        self.tray_added = False
        self.tray_add()
        self.appbar_register()
        self.reserve(self.reserved or self.S, force=True)
        self.layout()

    def _update_menu(self):
        self.menu.entryconfigure(0, label="収納" if self.state == EXPANDED else "展開")

    def show_menu(self, x, y):
        try:
            self.menu.tk_popup(x, y)
        finally:
            self.menu.grab_release()

    def show_menu_at_cursor(self):
        p = wt.POINT()
        GetCursorPos(ctypes.byref(p))
        SetForegroundWindow(self.strip_hwnd)   # メニュー外クリックで閉じるために必要
        self.show_menu(p.x, p.y)

    def menu_toggle(self):
        if self.fullscreen_hidden:
            return
        self.goto(EXPANDED) if self.state != EXPANDED else self.goto(COLLAPSED)

    # ================= 二重起動 =================
    def on_second_instance(self):
        self.notify(APP_NAME, "すでに起動しています。既存の帯を表示しました。")
        if not self.fullscreen_hidden and self.state != EXPANDED:
            self.goto(EXPANDED)

    # ================= AppBar（領域の予約） =================
    def appbar_register(self):
        if self.registered:
            SHAppBarMessage(ABM_REMOVE, ctypes.byref(self._abd()))
        abd = self._abd()
        abd.uCallbackMessage = self.appbar_msg
        SHAppBarMessage(ABM_NEW, ctypes.byref(abd))
        self.registered = True
        self.reserved = None

    def _refresh_monitor(self):
        mons = list_monitors()
        target = next((m for m in mons if m["device"] == self.s["Monitor"]), None) or \
            next((m for m in mons if m["primary"]), mons[0])
        self.mon = target
        r, w = target["rect"], target["work"]
        self.mon_key = (target["device"], r.left, r.top, r.right, r.bottom, w.top, w.bottom, len(mons))
        self.top, self.bottom = w.top, w.bottom   # タスクバーに重ならない高さ
        return target

    def reserve(self, dip, force=False):
        if not self.registered or (not force and dip == self.reserved):
            return
        self.reserved = dip
        m = self._refresh_monitor()["rect"]
        w = self.px(dip)
        abd = self._abd()
        if self.left:
            abd.uEdge = ABE_LEFT
            abd.rc = RECT(m.left, m.top, m.left + w, m.bottom)
        else:
            abd.uEdge = ABE_RIGHT
            abd.rc = RECT(m.right - w, m.top, m.right, m.bottom)
        SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(abd))
        if self.left:
            abd.rc.right = abd.rc.left + w
        else:
            abd.rc.left = abd.rc.right - w
        SHAppBarMessage(ABM_SETPOS, ctypes.byref(abd))
        self.anchor = abd.rc.left if self.left else abd.rc.right
        self._schedule_fix()

    def _schedule_fix(self):
        """予約変更で作業領域が変わると、Edge が自分で位置を直してしまう。少し後に置き直す"""
        for i in self.fix_ids:
            self.root.after_cancel(i)
        self.fix_ids = [self.root.after(ms, self._fix_position) for ms in (60, 250, 600, 1200)]

    def _fix_position(self):
        if self.closing or self.anim_id is not None or self.dragging or self.state == COLLAPSED \
                or self.fullscreen_hidden:
            return
        self.layout()

    def relayout(self):
        if self.closing:
            return
        self.reserve(self.reserved or self.S, force=True)
        if not self.fullscreen_hidden:
            self.layout()

    def reposition(self):
        """モニターや左右を変えたとき：予約を作り直して配置し直す"""
        self.cancel_anim()
        self.appbar_register()
        self._shape_tab()
        if self.state == EXPANDED:
            self.reserve(self.W + self.S)
            self.visible = self.W
        else:
            self.reserve(self.S)
            self.visible = 0 if self.state == COLLAPSED else self.W
        self.update_arrow()
        self.layout()

    # ================= 配置 =================
    def update_insets(self):
        ins = frame_insets(self.edge) if IsWindowVisible(self.edge) else None
        if ins and all(0 <= v <= 40 for v in ins):
            if ins != self.insets:
                log.info("frame insets: %s", ins)
            self.insets = ins

    def strip_x(self):
        spx = self.px(self.S)
        v = self.px(self.visible)
        return self.anchor + v if self.left else self.anchor - v - spx

    def layout(self):
        if self.fullscreen_hidden:
            return
        wpx, spx = self.px(self.W), self.px(self.S)
        sx = self.strip_x()
        h = self.bottom - self.top
        tb = self.px(self.s["TitleBarHeight"]) if self.s.get("HideTitleBar") else 0
        il, _, ir, ib = self.insets
        # 中身は幅を保ったまま、帯の外側から出入りする。見えない枠とタイトルバーの分は外へ逃がす
        ex = (sx - wpx - il) if self.left else (sx + spx - il)
        self.expected = (ex, self.top - tb, wpx + il + ir, h + tb + ib)
        SetWindowPos(self.edge, HWND_TOPMOST, *self.expected, SWP_NOACTIVATE)
        self.root.geometry(f"{spx}x{h}+{sx}+{self.top}")
        self.place_tab()

    def place_tab(self):
        spx = self.px(self.S)
        sx = self.strip_x()
        h = self.bottom - self.top
        tw, th = self.px(self.s["TabWidth"]), self.px(self.s["TabHeight"])
        pos = max(0.0, min(1.0, float(self.s["TabPosition"])))
        tx = sx + spx if self.left else sx - tw
        self.tab.geometry(f"{tw}x{th}+{tx}+{self.top + int((h - th) * pos)}")

    def _shape_tab(self):
        """つまみを角の丸い台形に切り抜く（右配置では左右反転）"""
        tw, th = self.px(self.s["TabWidth"]), self.px(self.s["TabHeight"])
        pts = [(round(x), round(y)) for x, y in tab_outline(tw, th, mirror=not self.left)]
        arr = (wt.POINT * len(pts))(*[wt.POINT(a, b) for a, b in pts])
        SetWindowRgn(self.tab_hwnd, CreatePolygonRgn(arr, len(pts), 2), True)
        self.tab_label.pack_configure(padx=(0, 3) if self.left else (3, 0))

    def update_arrow(self):
        out = "◀" if self.left else "▶"     # 収納する向き
        inn = "▶" if self.left else "◀"     # 展開する向き
        self.tab_label.configure(text=out if self.state == EXPANDED else inn)

    def _tab_hover(self, on):
        if not self.tab_down:
            self.tab.attributes("-alpha", 1.0 if on else float(self.s["TabOpacity"]))

    # ================= 状態遷移 =================
    def goto(self, target):
        if target == self.state:
            return
        self.state = target
        self.update_arrow()
        self.out_since = None
        if target in (EXPANDED, COLLAPSED) and self.s["LastState"] != target:
            self.s["LastState"] = target
            save_settings(self.s)

        if target == EXPANDED:
            self.reserve(self.W + self.S)          # 予約を広げてから
            ShowWindow(self.edge, SW_SHOWNOACTIVATE)
            self.animate(self.W, None)             # スライドイン
        elif target == COLLAPSED:
            def done():
                self.reserve(self.S)               # スライドアウト後に予約を縮める
                ShowWindow(self.edge, SW_HIDE)
            self.animate(0, done)
        elif target == PEEK:
            self.reserve(self.S)
            ShowWindow(self.edge, SW_SHOWNOACTIVATE)
            self.animate(self.W, None)

    def toggle(self):
        self.goto(COLLAPSED if self.state == EXPANDED else EXPANDED)

    def cancel_anim(self):
        if self.anim_id is not None:
            self.root.after_cancel(self.anim_id)
            self.anim_id = None

    def animate(self, to, done):
        self.cancel_anim()
        frm = self.visible
        ms = int(self.s["AnimationMs"])
        if ms <= 0 or abs(to - frm) < 1:
            self.visible = to
            self.layout()
            if done:
                done()
            return
        t0 = time.perf_counter()

        def step():
            t = min(1.0, (time.perf_counter() - t0) * 1000.0 / ms)
            eased = 1 - (1 - t) ** 3
            self.visible = frm + (to - frm) * eased
            self.layout()
            if t < 1.0:
                self.anim_id = self.root.after(10, step)
            else:
                self.anim_id = None
                if done:
                    done()

        step()

    # ================= 帯：クリック / 左右ドラッグで幅 =================
    def on_press(self, e):
        self.down = True
        self.dragging = False
        self.drag_x = e.x_root
        self.drag_w = self.W
        self.hover_since = None

    def on_motion(self, e):
        if not self.down or self.state == COLLAPSED:
            return
        dx = e.x_root - self.drag_x
        if not self.dragging and abs(dx) < 4:
            return
        if not self.dragging:
            self.dragging = True
            self.cancel_anim()
        if not self.left:
            dx = -dx
        w = int(round(self.drag_w + dx / self.scale))
        self.s["Width"] = max(self.s["MinWidth"], min(self.s["MaxWidth"], w))
        self.visible = self.W
        self.layout()                               # 予約は離した時に確定

    def on_release(self, e):
        if not self.down:
            return
        self.down = False
        if self.dragging:
            self.dragging = False
            if self.state == EXPANDED:
                self.reserve(self.W + self.S)
            save_settings(self.s)
        elif self.state == PEEK:
            self.goto(EXPANDED)
        else:
            self.toggle()

    def set_width(self, w):
        self.s["Width"] = max(self.s["MinWidth"], min(self.s["MaxWidth"], int(w)))
        save_settings(self.s)
        if self.state != COLLAPSED:
            self.visible = self.W
            if self.state == EXPANDED:
                self.reserve(self.W + self.S)
            self.layout()

    # ================= つまみ：クリック / 上下ドラッグで位置 =================
    def on_tab_press(self, e):
        self.tab_down = True
        self.tab_moved = False
        self.tab_drag_y = e.y_root
        self.tab_drag_pos = float(self.s["TabPosition"])
        self.hover_since = None
        self.tab.attributes("-alpha", 1.0)

    def on_tab_motion(self, e):
        if not self.tab_down:
            return
        dy = e.y_root - self.tab_drag_y
        if not self.tab_moved and abs(dy) < 4:
            return
        self.tab_moved = True
        span = (self.bottom - self.top) - self.px(self.s["TabHeight"])
        if span > 0:
            self.s["TabPosition"] = max(0.0, min(1.0, self.tab_drag_pos + dy / span))
            self.place_tab()

    def on_tab_release(self, e):
        if not self.tab_down:
            return
        self.tab_down = False
        if self.tab_moved:
            save_settings(self.s)
        elif self.state == PEEK:
            self.goto(EXPANDED)
        else:
            self.toggle()
        p = wt.POINT()
        GetCursorPos(ctypes.byref(p))
        self._tab_hover(inside(p, rect_of(self.tab_hwnd)))

    # ================= ダッシュボードを開き直す（再読み込み・ドロップ・自動復旧） =================
    def open_url(self, url):
        """新しいアプリウィンドウで開き、差し替える（拡張機能なしでページを移動するための方法）"""
        if self.navigating:
            return False
        self.navigating = True
        old = self.edge
        try:
            new = launch_app_window(self.browser, url)
            if not new:
                self.notify(APP_NAME, "ダッシュボードを開けませんでした。")
                return False
            self.edge = new
            self.update_insets()
            if old and IsWindow(old):
                PostMessageW(old, WM_CLOSE, 0, 0)
            if self.fullscreen_hidden:
                ShowWindow(self.edge, SW_HIDE)
            elif self.state == COLLAPSED:
                ShowWindow(self.edge, SW_HIDE)
                self.layout()
            else:
                self.layout()
            return True
        finally:
            self.navigating = False

    def reload(self):
        self.open_url(sidebar_url(self.s["Url"]))

    def on_drop(self, files):
        if not files:
            return
        path = files[0]
        if not path.lower().endswith(".exe"):
            self.notify(APP_NAME, "いまドロップで追加できるのは exe だけです。")
            return
        title = os.path.splitext(os.path.basename(path))[0]
        log.info("drop: %s", path)
        if self.open_url(sidebar_url(self.s["Url"], add="exe", title=title, path=path)):
            if self.state != EXPANDED:
                self.goto(EXPANDED)
            if len(files) > 1:
                self.notify(APP_NAME, "複数のファイルがドロップされたため、最初の1つだけ追加します。")

    def recover(self):
        """ダッシュボードのウィンドウが消えたら開き直す（2分間に3回失敗したらあきらめる）"""
        now = time.time()
        self.recover_times = [t for t in self.recover_times if now - t < 120] + [now]
        if len(self.recover_times) > 3:
            self.notify(APP_NAME, "ダッシュボードを開き直せなかったため終了します。")
            self.quit()
            return
        log.info("recover: reopening dashboard")
        new = launch_app_window(self.browser, sidebar_url(self.s["Url"]))
        if not new:
            return
        self.edge = new
        self.update_insets()
        if self.state == COLLAPSED or self.fullscreen_hidden:
            ShowWindow(self.edge, SW_HIDE)
        self.layout()
        self.notify(APP_NAME, "ダッシュボードを開き直しました。")

    # ================= 全画面アプリ中は隠す =================
    def foreground_is_fullscreen(self):
        fg = GetForegroundWindow()
        if not fg:
            return False
        own = {self.edge, self.strip_hwnd, self.tab_hwnd}
        if self.settings_win is not None:
            try:
                own.add(GetAncestor(self.settings_win.winfo_id(), GA_ROOT))
            except Exception:
                pass
        if fg in own or class_name(fg) in SHELL_CLASSES:
            return False
        r, m = rect_of(fg), self.mon["rect"]
        return r.left <= m.left and r.top <= m.top and r.right >= m.right and r.bottom >= m.bottom

    def set_fullscreen_hidden(self, hidden):
        if hidden == self.fullscreen_hidden:
            return
        self.fullscreen_hidden = hidden
        log.info("fullscreen hidden: %s", hidden)
        if hidden:
            self.cancel_anim()
            if self.state == PEEK:
                self.state = COLLAPSED
                self.visible = 0
                self.update_arrow()
            ShowWindow(self.edge, SW_HIDE)
            self.tab.withdraw()
            self.root.withdraw()
        else:
            self.root.deiconify()
            self.tab.deiconify()
            self._shape_tab()
            if self.state != COLLAPSED:
                ShowWindow(self.edge, SW_SHOWNOACTIVATE)
            self.layout()

    # ================= 設定画面 =================
    def open_settings(self):
        if self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.deiconify()
            self.settings_win.lift()
            return
        SettingsWindow(self)

    def apply_tab_opacity(self, v):
        self.s["TabOpacity"] = round(float(v), 2)
        self.tab.attributes("-alpha", self.s["TabOpacity"])

    # ================= 更新 =================
    def check_update(self, manual):
        url = self.s.get("UpdateUrl")
        if not url:
            return

        def work():
            try:
                req = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    src = r.read().decode("utf-8")
                latest = parse_version(src)
                if not latest:
                    raise ValueError("VERSION が見つかりません")
                self.events.append(lambda: self._update_checked(latest, src, manual))
            except Exception as e:
                log.warning("update check failed: %s", e)
                if manual:
                    msg = f"更新を確認できませんでした（{e}）"
                    self.events.append(lambda: self.notify(APP_NAME, msg))

        threading.Thread(target=work, daemon=True).start()

    def _update_checked(self, latest, src, manual):
        if version_tuple(latest) > version_tuple(VERSION):
            self.pending_update = (latest, src)
            if manual:
                self.prompt_update()
            else:
                self.notify(APP_NAME, f"新しいバージョン v{latest} があります。クリックして更新します。")
        elif manual:
            self.notify(APP_NAME, f"最新です（v{VERSION}）。")

    def prompt_update(self):
        if not self.pending_update:
            return
        latest, src = self.pending_update
        if not messagebox.askyesno(APP_NAME, f"新しいバージョン v{latest} があります（現在 v{VERSION}）。\n"
                                             f"更新して再起動しますか？", parent=self.root):
            return
        try:
            compile(src, SCRIPT_PATH, "exec")          # 壊れたファイルで上書きしない
            backup = SCRIPT_PATH + ".bak"
            with open(SCRIPT_PATH, "rb") as f:
                old = f.read()
            with open(backup, "wb") as f:
                f.write(old)
            with open(SCRIPT_PATH, "w", encoding="utf-8", newline="") as f:
                f.write(src)
            log.info("updated to v%s (backup: %s)", latest, backup)
        except Exception as e:
            log.exception("update failed")
            messagebox.showerror(APP_NAME, f"更新に失敗しました。\n{e}", parent=self.root)
            return
        self.restart()

    def restart(self):
        self.quit(close_edge=True, release_mutex=True)
        subprocess.Popen([pythonw_path(), SCRIPT_PATH], cwd=BASE_DIR)

    # ================= 定期処理 =================
    def fast_poll(self):
        if self.closing:
            return
        while self.events:
            fn = self.events.popleft()
            try:
                fn()
            except Exception:
                log.exception("event")
            if self.closing:
                return
        try:
            self.poll_count += 1
            # 全画面アプリの検出（約 0.25 秒ごと）
            if self.poll_count % 5 == 0 and self.s["HideOnFullscreen"] and not self.navigating:
                self.set_fullscreen_hidden(self.foreground_is_fullscreen())
            elif not self.s["HideOnFullscreen"] and self.fullscreen_hidden:
                self.set_fullscreen_hidden(False)

            if not self.fullscreen_hidden:
                # Ctrl+Alt+D
                pressed = all(GetAsyncKeyState(k) & 0x8000 for k in (VK_CONTROL, VK_MENU, VK_D))
                if pressed and not self.hotkey_prev:
                    self.goto(EXPANDED) if self.state == PEEK else self.toggle()
                self.hotkey_prev = pressed

                p = wt.POINT()
                GetCursorPos(ctypes.byref(p))
                in_strip = inside(p, rect_of(self.strip_hwnd))   # 覗き見の開始は帯だけ
                in_tab = inside(p, rect_of(self.tab_hwnd))
                busy = self.down or self.tab_down
                now = time.time()

                if self.s["HoverPeek"] and self.state == COLLAPSED and not busy and in_strip:
                    if self.hover_since is None:
                        self.hover_since = now
                    elif now - self.hover_since > 0.25:
                        self.hover_since = None
                        self.goto(PEEK)
                else:
                    self.hover_since = None

                if self.state == PEEK and not busy:
                    if in_strip or in_tab or inside(p, rect_of(self.edge)):
                        self.out_since = None
                    elif self.out_since is None:
                        self.out_since = now
                    elif now - self.out_since > 0.4:
                        self.goto(COLLAPSED)
        except Exception:
            log.exception("fast_poll")
        self.root.after(50, self.fast_poll)

    def slow_poll(self):
        if self.closing:
            return
        try:
            if not self.navigating and not IsWindow(self.edge):
                log.info("Edge window closed")
                if self.s["AutoRecover"]:
                    self.recover()
                else:
                    self.quit()
                    return

            if not self.fullscreen_hidden:
                old = self.insets
                self.update_insets()
                if self.insets != old and self.state != COLLAPSED and self.anim_id is None:
                    self.layout()

                # Edge が勝手に動いていたら置き直す
                if (self.state != COLLAPSED and self.anim_id is None and not self.dragging
                        and not self.navigating and self.expected and IsWindowVisible(self.edge)):
                    r = rect_of(self.edge)
                    ex = self.expected
                    if (abs(r.left - ex[0]) > 2 or abs(r.top - ex[1]) > 2
                            or abs((r.right - r.left) - ex[2]) > 2):
                        log.info("Edge がずれていたので置き直し: %s -> %s", (r.left, r.top, r.right - r.left), ex)
                        self.layout()

            # モニター構成や解像度の変化
            old_key = self.mon_key
            self._refresh_monitor()
            if self.mon_key != old_key and self.anim_id is None and not self.dragging:
                log.info("monitor changed: %s", self.mon_key)
                self.reposition()
        except Exception:
            log.exception("slow_poll")
        self.root.after(1000, self.slow_poll)

    # ================= 終了 =================
    def cleanup(self):
        if self.registered:
            self.registered = False
            try:
                SHAppBarMessage(ABM_REMOVE, ctypes.byref(self._abd()))
            except Exception:
                log.exception("ABM_REMOVE")
        try:
            self.tray_remove()
        except Exception:
            pass

    def quit(self, close_edge=True, release_mutex=False):
        if self.closing:
            return
        self.closing = True
        self.cancel_anim()
        self.cleanup()
        try:
            if close_edge and IsWindow(self.edge):
                PostMessageW(self.edge, WM_CLOSE, 0, 0)
        except Exception:
            pass
        try:
            os.remove(INSTANCE_PATH)
        except Exception:
            pass
        if release_mutex and self.mutex:
            CloseHandle(self.mutex)
            self.mutex = None
        self.root.destroy()

    def _on_tk_error(self, exc, val, tb):
        log.error("tk callback error", exc_info=(exc, val, tb))

    def _fatal(self, msg):
        log.error(msg)
        messagebox.showerror(APP_NAME, msg)
        self.root.destroy()
        sys.exit(1)

    def run(self):
        try:
            self.root.mainloop()
        finally:
            self.cleanup()


# ---------------- 設定画面 ----------------
class SettingsWindow:
    def __init__(self, bar):
        self.bar = bar
        s = bar.s
        self.win = w = tk.Toplevel(bar.root)
        bar.settings_win = w
        w.title(f"{APP_NAME} の設定")
        w.attributes("-topmost", True)
        w.resizable(False, False)
        try:
            w.iconbitmap(ensure_icon())
        except Exception:
            pass
        w.protocol("WM_DELETE_WINDOW", self.close)

        f = ttk.Frame(w, padding=14)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        row = 0

        def section(text):
            nonlocal row
            ttk.Label(f, text=text, font=("Segoe UI", 9, "bold")).grid(row=row, column=0, columnspan=2,
                                                                      sticky="w", pady=(10 if row else 0, 4))
            row += 1

        def line(label, widget):
            nonlocal row
            ttk.Label(f, text=label).grid(row=row, column=0, sticky="w", padx=(8, 12), pady=3)
            widget.grid(row=row, column=1, sticky="we", pady=3)
            row += 1

        def check(label, key=None, var=None, cmd=None):
            nonlocal row
            v = var or tk.BooleanVar(value=bool(s[key]))
            ttk.Checkbutton(f, text=label, variable=v,
                            command=cmd or (lambda: self.set_flag(key, v.get()))).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=8, pady=2)
            row += 1
            return v

        # ---- 表示 ----
        section("表示")
        self.mons = list_monitors()
        labels = [monitor_label(i, m) for i, m in enumerate(self.mons)]
        cur = next((i for i, m in enumerate(self.mons) if m["device"] == bar.mon["device"]), 0)
        self.mon_var = tk.StringVar(value=labels[cur])
        cb = ttk.Combobox(f, textvariable=self.mon_var, values=labels, state="readonly", width=30)
        cb.bind("<<ComboboxSelected>>", lambda e: self.set_monitor(labels.index(self.mon_var.get())))
        line("モニター", cb)

        side = ttk.Frame(f)
        self.side_var = tk.StringVar(value=s["Side"])
        ttk.Radiobutton(side, text="左端", value="left", variable=self.side_var,
                        command=self.set_side).pack(side="left")
        ttk.Radiobutton(side, text="右端", value="right", variable=self.side_var,
                        command=self.set_side).pack(side="left", padx=(12, 0))
        line("表示位置", side)

        self.width_var = tk.IntVar(value=s["Width"])
        sp = ttk.Spinbox(f, from_=s["MinWidth"], to=s["MaxWidth"], increment=10, width=8,
                         textvariable=self.width_var, command=self.set_width)
        sp.bind("<Return>", lambda e: self.set_width())
        sp.bind("<FocusOut>", lambda e: self.set_width())
        line("展開時の幅（px）", sp)

        self.op_var = tk.DoubleVar(value=s["TabOpacity"])
        ttk.Scale(f, from_=0.2, to=1.0, variable=self.op_var,
                  command=lambda v: bar.apply_tab_opacity(v)).grid(row=row, column=1, sticky="we", pady=3)
        ttk.Label(f, text="つまみの不透明度").grid(row=row, column=0, sticky="w", padx=(8, 12), pady=3)
        row += 1

        self.url_var = tk.StringVar(value=s["Url"])
        line("ダッシュボードのURL", ttk.Entry(f, textvariable=self.url_var, width=34))

        # ---- 動作 ----
        section("動作")
        check("帯にマウスを乗せたら覗き見表示する", "HoverPeek")
        check("全画面アプリの使用中は隠す", "HideOnFullscreen")
        check("ダッシュボードが閉じたら自動で開き直す", "AutoRecover")

        # ---- 起動と更新 ----
        section("起動と更新")
        self.auto_var = tk.BooleanVar(value=get_autostart())
        check("Windows の起動時に自動で起動する", var=self.auto_var, cmd=self.set_autostart)
        check("前回の展開・収納の状態を復元する", "RestoreState")
        check("起動時に更新を確認する", "CheckUpdateOnStart")
        upd = ttk.Frame(f)
        ttk.Label(upd, text=f"バージョン {VERSION}").pack(side="left")
        ttk.Button(upd, text="今すぐ確認", command=lambda: bar.check_update(manual=True)).pack(side="left", padx=10)
        upd.grid(row=row, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 0))
        row += 1

        btns = ttk.Frame(f)
        ttk.Button(btns, text="閉じる", command=self.close).pack(side="right")
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(14, 0))

        # 対象モニターの中央に表示
        w.update_idletasks()
        m = bar.mon["work"]
        x = m.left + ((m.right - m.left) - w.winfo_width()) // 2
        y = m.top + ((m.bottom - m.top) - w.winfo_height()) // 2
        w.geometry(f"+{x}+{y}")
        w.focus_force()

    def set_flag(self, key, value):
        self.bar.s[key] = bool(value)
        save_settings(self.bar.s)

    def set_monitor(self, i):
        self.bar.s["Monitor"] = self.mons[i]["device"]
        save_settings(self.bar.s)
        self.bar.reposition()

    def set_side(self):
        self.bar.s["Side"] = self.side_var.get()
        save_settings(self.bar.s)
        self.bar.reposition()

    def set_width(self):
        try:
            self.bar.set_width(int(self.width_var.get()))
        except (ValueError, tk.TclError):
            pass
        self.width_var.set(self.bar.W)

    def set_autostart(self):
        try:
            set_autostart(self.auto_var.get())
        except Exception as e:
            log.exception("autostart")
            messagebox.showerror(APP_NAME, f"自動起動の設定に失敗しました。\n{e}", parent=self.win)
            self.auto_var.set(get_autostart())

    def close(self):
        bar = self.bar
        save_settings(bar.s)
        new_url = self.url_var.get().strip()
        url_changed = new_url and new_url != bar.s["Url"]
        bar.settings_win = None
        self.win.destroy()
        if url_changed:
            bar.s["Url"] = new_url
            save_settings(bar.s)
            bar.reload()


# ---------------- 起動 ----------------
def activate_existing():
    """既に起動している DashboardBar に知らせる"""
    hwnd = None
    try:
        with open(INSTANCE_PATH) as f:
            hwnd = int(f.read().strip())
        if not IsWindow(hwnd):
            hwnd = None
    except Exception:
        hwnd = None
    if not hwnd:
        hwnd = FindWindowW(None, STRIP_TITLE)
    if hwnd:
        PostMessageW(hwnd, WM_APP_ACTIVATE, 0, 0)


def main():
    mutex = CreateMutexW(None, False, "Local\\DashboardBarPy_SingleInstance")
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        log.info("already running")
        activate_existing()
        return
    try:
        DashboardBar(mutex).run()
    except SystemExit:
        raise
    except Exception:
        log.exception("fatal")
        raise


if __name__ == "__main__":
    main()
