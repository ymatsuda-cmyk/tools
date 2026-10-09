"""
DashboardBar（Python版）
- Edge のアプリモードでダッシュボードを表示し、画面左端に AppBar として固定する
- つまみクリック: 展開/収納  つまみ上下ドラッグ: 位置移動  帯左右ドラッグ: 幅変更
- ホバー: 覗き見  Ctrl+Alt+D: 切替  右クリック: メニュー
- 標準ライブラリのみ（ctypes + tkinter）。pythonw.exe で起動する想定
"""
import ctypes
import json
import logging
import os
import subprocess
import sys
import time
import tkinter as tk
import urllib.parse
from ctypes import wintypes as wt
from tkinter import messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")
LOG_PATH = os.path.join(BASE_DIR, "dashboard_bar.log")

logging.basicConfig(filename=LOG_PATH, level=logging.INFO, encoding="utf-8",
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("bar")

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
dwmapi = ctypes.WinDLL("dwmapi")
gdi32 = ctypes.WinDLL("gdi32")


def _fn(dll, name, res, *args):
    f = getattr(dll, name)
    f.restype = res
    f.argtypes = list(args)
    return f


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wt.DWORD), ("rcMonitor", RECT), ("rcWork", RECT), ("dwFlags", wt.DWORD)]


class APPBARDATA(ctypes.Structure):
    _fields_ = [("cbSize", wt.DWORD), ("hWnd", wt.HWND), ("uCallbackMessage", wt.UINT),
                ("uEdge", wt.UINT), ("rc", RECT), ("lParam", wt.LPARAM)]


WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

SHAppBarMessage = _fn(shell32, "SHAppBarMessage", ctypes.c_size_t, wt.DWORD, ctypes.POINTER(APPBARDATA))
EnumWindows = _fn(user32, "EnumWindows", wt.BOOL, WNDENUMPROC, wt.LPARAM)
GetClassNameW = _fn(user32, "GetClassNameW", ctypes.c_int, wt.HWND, wt.LPWSTR, ctypes.c_int)
GetWindowTextLengthW = _fn(user32, "GetWindowTextLengthW", ctypes.c_int, wt.HWND)
IsWindowVisible = _fn(user32, "IsWindowVisible", wt.BOOL, wt.HWND)
IsWindow = _fn(user32, "IsWindow", wt.BOOL, wt.HWND)
GetWindowLongW = _fn(user32, "GetWindowLongW", ctypes.c_long, wt.HWND, ctypes.c_int)
SetWindowLongW = _fn(user32, "SetWindowLongW", ctypes.c_long, wt.HWND, ctypes.c_int, ctypes.c_long)
SetWindowPos = _fn(user32, "SetWindowPos", wt.BOOL, wt.HWND, wt.HWND,
                   ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.UINT)
ShowWindow = _fn(user32, "ShowWindow", wt.BOOL, wt.HWND, ctypes.c_int)
GetCursorPos = _fn(user32, "GetCursorPos", wt.BOOL, ctypes.POINTER(wt.POINT))
GetWindowRect = _fn(user32, "GetWindowRect", wt.BOOL, wt.HWND, ctypes.POINTER(RECT))
GetAncestor = _fn(user32, "GetAncestor", wt.HWND, wt.HWND, wt.UINT)
MonitorFromPoint = _fn(user32, "MonitorFromPoint", wt.HMONITOR, wt.POINT, wt.DWORD)
GetMonitorInfoW = _fn(user32, "GetMonitorInfoW", wt.BOOL, wt.HMONITOR, ctypes.POINTER(MONITORINFO))
GetAsyncKeyState = _fn(user32, "GetAsyncKeyState", ctypes.c_short, ctypes.c_int)
RegisterWindowMessageW = _fn(user32, "RegisterWindowMessageW", wt.UINT, wt.LPCWSTR)
DwmGetWindowAttribute = _fn(dwmapi, "DwmGetWindowAttribute", ctypes.c_long,
                            wt.HWND, wt.DWORD, ctypes.c_void_p, wt.DWORD)
CreateRoundRectRgn = _fn(gdi32, "CreateRoundRectRgn", wt.HRGN,
                         ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int)
CreatePolygonRgn = _fn(gdi32, "CreatePolygonRgn", wt.HRGN,
                       ctypes.POINTER(wt.POINT), ctypes.c_int, ctypes.c_int)
SetWindowRgn = _fn(user32, "SetWindowRgn", ctypes.c_int, wt.HWND, wt.HRGN, wt.BOOL)
PostMessageW = _fn(user32, "PostMessageW", wt.BOOL, wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)

try:
    GetDpiForSystem = _fn(user32, "GetDpiForSystem", wt.UINT)
except AttributeError:
    GetDpiForSystem = lambda: 96  # noqa: E731

ABM_NEW, ABM_REMOVE, ABM_QUERYPOS, ABM_SETPOS = 0, 1, 2, 3
ABE_LEFT = 0
GWL_STYLE = -16
WS_CAPTION, WS_THICKFRAME = 0x00C00000, 0x00040000
SWP_NOSIZE, SWP_NOMOVE, SWP_NOZORDER, SWP_NOACTIVATE, SWP_FRAMECHANGED = 0x1, 0x2, 0x4, 0x10, 0x20
SW_HIDE, SW_SHOWNOACTIVATE = 0, 4
HWND_TOPMOST = wt.HWND(-1)
DWMWA_EXTENDED_FRAME_BOUNDS = 9
WM_CLOSE = 0x0010
GA_ROOT = 2
VK_CONTROL, VK_MENU, VK_D = 0x11, 0x12, 0x44

EXPANDED, COLLAPSED, PEEK = "expanded", "collapsed", "peek"

# ---------------- 設定 ----------------
DEFAULTS = {
    "Url": "https://ymatsuda-cmyk.github.io/tools/dashboard/",
    "Browser": "",            # 空なら Edge を自動検出
    "Width": 380,
    "MinWidth": 240,
    "MaxWidth": 800,
    "StripWidth": 6,          # 細い帯の幅
    "TabWidth": 20,           # つまみの幅
    "TabHeight": 68,          # つまみの高さ
    "TabPosition": 0.5,       # つまみの縦位置（0=上端, 0.5=中央, 1=下端）
    "StripColor": "#C9C5BC",
    "TabColor": "#D3D1C7",
    "TabHoverColor": "#D3D1C7",
    "ArrowColor": "#B4B2A9",  # 薄いグレー
    "TabOpacity": 0.55,       # つまみの不透明度（マウスを乗せると 1.0）
    "AnimationMs": 260,
    "StartCollapsed": False,
    "HoverPeek": True,
    "HideTitleBar": True,     # Edge のタイトルバーを画面外に追い出して隠す
    "TitleBarHeight": 32,     # 隠すタイトルバーの高さ（DIP）。ずれる場合はここを調整
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
    s["Width"] = max(s["MinWidth"], min(s["MaxWidth"], int(s["Width"])))
    return s


def save_settings(s):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning("settings.json の保存に失敗: %s", e)


def sidebar_url(url):
    """ダッシュボード側がサイドバー表示と判定できるよう ?sidebar=1 を付ける"""
    parts = urllib.parse.urlsplit(url)
    q = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
    if not any(k == "sidebar" for k, _ in q):
        q.append(("sidebar", "1"))
    return urllib.parse.urlunsplit(parts._replace(query=urllib.parse.urlencode(q)))


def find_browser(s):
    if s.get("Browser") and os.path.exists(s["Browser"]):
        return s["Browser"]
    roots = [os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")]
    for r in roots:
        if r:
            p = os.path.join(r, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.exists(p):
                return p
    return None


def list_browser_windows():
    found = []

    def cb(h, _):
        if h and IsWindowVisible(h):
            buf = ctypes.create_unicode_buffer(256)
            GetClassNameW(h, buf, 256)
            if buf.value == "Chrome_WidgetWin_1" and GetWindowTextLengthW(h) > 0:
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


def primary_monitor():
    mon = MonitorFromPoint(wt.POINT(0, 0), 1)  # MONITOR_DEFAULTTOPRIMARY
    mi = MONITORINFO()
    mi.cbSize = ctypes.sizeof(MONITORINFO)
    GetMonitorInfoW(mon, ctypes.byref(mi))
    return mi.rcMonitor, mi.rcWork


def rect_of(hwnd):
    r = RECT()
    GetWindowRect(hwnd, ctypes.byref(r))
    return r


def frame_insets(hwnd):
    """Windows 11 の見えない枠（リサイズ用の透明な境界）の幅を返す: (左, 上, 右, 下)"""
    win = rect_of(hwnd)
    vis = RECT()
    hr = DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
                               ctypes.byref(vis), ctypes.sizeof(RECT))
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


def tab_outline(w, h):
    """角を丸めた台形（左＝帯側が高く、右が低い）の輪郭。20x68 の形を基準に、角の大きさは保ったまま縦に伸ばす"""
    sx = w / 20.0
    c = min(20.0, h / 2.0)  # 上下の曲線部分の高さ
    k = c / 20.0

    def P(x, y):
        return (x * sx, y)

    pts = [P(0, 0)]
    pts += _bezier(P(0, 0), P(4, 0), P(8, 3 * k), P(13, 7 * k))
    pts += _bezier(P(13, 7 * k), P(17, 10 * k), P(20, 14 * k), P(20, c))
    pts.append(P(20, h - c))
    pts += _bezier(P(20, h - c), P(20, h - 14 * k), P(17, h - 10 * k), P(13, h - 7 * k))
    pts += _bezier(P(13, h - 7 * k), P(8, h - 3 * k), P(4, h), P(0, h))
    return pts


def inside(p, r):
    return r.left <= p.x < r.right and r.top <= p.y < r.bottom


# ---------------- 本体 ----------------
class DashboardBar:
    def __init__(self):
        self.s = load_settings()
        self.scale = GetDpiForSystem() / 96.0

        self.state = COLLAPSED
        self.visible = 0.0          # 見えている中身の幅（DIP）
        self.reserved = None        # 予約中の幅（DIP）
        self.left = 0
        self.top = 0
        self.bottom = 0
        self.mon_key = None
        self.anim_id = None
        self.insets = (0, 0, 0, 0)  # Edge の見えない枠
        self.registered = False
        self.closing = False

        self.down = False
        self.dragging = False
        self.drag_x = 0
        self.drag_w = 0
        self.tab_down = False
        self.tab_moved = False
        self.tab_drag_y = 0
        self.tab_drag_pos = 0.5
        self.expected = None        # Edge を置いたはずの位置（ずれ検出用）
        self.fix_ids = []
        self.hover_since = None
        self.out_since = None
        self.hotkey_prev = False

        # 帯（tkinter）
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.s["StripColor"], cursor="sb_h_double_arrow")
        self.root.report_callback_exception = self._on_tk_error

        # つまみ（帯から右に飛び出すタブ）
        self.tab = tk.Toplevel(self.root)
        self.tab.withdraw()
        self.tab.overrideredirect(True)
        self.tab.attributes("-topmost", True)
        self.tab.configure(bg=self.s["TabColor"], cursor="hand2")
        self.tab_label = tk.Label(self.tab, text="◀", bg=self.s["TabColor"], fg=self.s["ArrowColor"],
                                  font=("Segoe UI", 8), cursor="hand2", padx=0)
        self.tab_label.pack(expand=True, fill="both", padx=(0, 3))
        self.tab.attributes("-alpha", float(self.s["TabOpacity"]))
        self.tab.bind("<Enter>", lambda e: self._tab_hover(True))
        self.tab.bind("<Leave>", lambda e: self._tab_hover(False))

        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="幅を初期値に戻す", command=self.reset_width)
        self.menu.add_separator()
        self.menu.add_command(label="終了", command=self.quit)

        # 帯：左右ドラッグで幅変更（クリックで切替）
        self.root.bind("<ButtonPress-1>", self.on_press)
        self.root.bind("<B1-Motion>", self.on_motion)
        self.root.bind("<ButtonRelease-1>", self.on_release)
        # つまみ：上下ドラッグで位置移動（クリックで切替）
        self.tab.bind("<ButtonPress-1>", self.on_tab_press)
        self.tab.bind("<B1-Motion>", self.on_tab_motion)
        self.tab.bind("<ButtonRelease-1>", self.on_tab_release)
        for w in (self.root, self.tab):
            w.bind("<Button-3>", lambda e: self.menu.tk_popup(e.x_root, e.y_root))

        # Edge アプリウィンドウを起動
        browser = find_browser(self.s)
        if not browser:
            self._fatal("Edge（msedge.exe）が見つかりません。settings.json の Browser にパスを指定してください。")
        self.edge = launch_app_window(browser, sidebar_url(self.s["Url"]))
        if not self.edge:
            self._fatal("ダッシュボードのウィンドウを見つけられませんでした。")
        log.info("Edge window: %s", self.edge)
        self.update_insets()

        # 帯を表示して AppBar 登録
        self.root.deiconify()
        self.tab.deiconify()
        self.root.update_idletasks()
        self.strip_hwnd = GetAncestor(self.root.winfo_id(), GA_ROOT)
        self.tab_hwnd = GetAncestor(self.tab.winfo_id(), GA_ROOT)
        self._shape_tab()
        abd = self._abd()
        abd.uCallbackMessage = RegisterWindowMessageW("DashboardBarPy_AppBarMsg")
        SHAppBarMessage(ABM_NEW, ctypes.byref(abd))
        self.registered = True

        self._refresh_monitor()

        if self.s["StartCollapsed"]:
            self.state = COLLAPSED
            self.reserve(self.S)
            self.visible = 0
            self.layout()
            ShowWindow(self.edge, SW_HIDE)
        else:
            self.state = EXPANDED
            self.reserve(self.W + self.S)
            self.visible = self.W
            self.layout()
        self.update_arrow()

        self.root.after(50, self.fast_poll)
        self.root.after(1000, self.slow_poll)

    # ---- 便利プロパティ ----
    @property
    def W(self):
        return int(self.s["Width"])

    @property
    def S(self):
        return int(self.s["StripWidth"])

    def px(self, dip):
        return int(round(dip * self.scale))

    def _abd(self):
        abd = APPBARDATA()
        abd.cbSize = ctypes.sizeof(APPBARDATA)
        abd.hWnd = self.strip_hwnd
        return abd

    # ---- 予約と配置 ----
    def _refresh_monitor(self):
        m, w = primary_monitor()
        self.mon_key = (m.left, m.top, m.right, m.bottom, w.top, w.bottom)
        self.top, self.bottom = w.top, w.bottom   # タスクバーに重ならない高さ
        return m

    def reserve(self, dip, force=False):
        if not self.registered or (not force and dip == self.reserved):
            return
        self.reserved = dip
        m = self._refresh_monitor()
        abd = self._abd()
        abd.uEdge = ABE_LEFT
        abd.rc = RECT(m.left, m.top, m.left + self.px(dip), m.bottom)
        SHAppBarMessage(ABM_QUERYPOS, ctypes.byref(abd))
        abd.rc.right = abd.rc.left + self.px(dip)
        SHAppBarMessage(ABM_SETPOS, ctypes.byref(abd))
        self.left = abd.rc.left
        self._schedule_fix()

    def _schedule_fix(self):
        """予約変更で作業領域が変わると、Edge が自分で位置を直してしまう。少し後に置き直す"""
        for i in self.fix_ids:
            self.root.after_cancel(i)
        self.fix_ids = [self.root.after(ms, self._fix_position) for ms in (60, 250, 600, 1200)]

    def _fix_position(self):
        if self.closing or self.anim_id is not None or self.dragging or self.state == COLLAPSED:
            return
        self.layout()

    def update_insets(self):
        ins = frame_insets(self.edge) if IsWindowVisible(self.edge) else None
        if ins and all(0 <= v <= 40 for v in ins):
            if ins != self.insets:
                log.info("frame insets: %s", ins)
            self.insets = ins

    def layout(self):
        wpx, spx = self.px(self.W), self.px(self.S)
        x = self.left + self.px(self.visible)
        h = self.bottom - self.top
        # タイトルバーの分だけ上にずらし、画面外に追い出して隠す
        tb = self.px(self.s["TitleBarHeight"]) if self.s.get("HideTitleBar") else 0
        # 中身は幅を保ったまま、左端から出入りする
        # 見えない枠の分を外側に広げ、見た目の端を帯にぴったり合わせる
        il, _, ir, ib = self.insets
        ex = (x - wpx - il, self.top - tb, wpx + il + ir, h + tb + ib)
        self.expected = ex
        SetWindowPos(self.edge, HWND_TOPMOST, ex[0], ex[1], ex[2], ex[3], SWP_NOACTIVATE)
        self.root.geometry(f"{spx}x{h}+{x}+{self.top}")
        self.place_tab()

    def place_tab(self):
        spx = self.px(self.S)
        x = self.left + self.px(self.visible)
        h = self.bottom - self.top
        tw, th = self.px(self.s["TabWidth"]), self.px(self.s["TabHeight"])
        pos = max(0.0, min(1.0, float(self.s["TabPosition"])))
        self.tab.geometry(f"{tw}x{th}+{x + spx}+{self.top + int((h - th) * pos)}")

    def _shape_tab(self):
        """つまみを角の丸い台形に切り抜く"""
        tw, th = self.px(self.s["TabWidth"]), self.px(self.s["TabHeight"])
        pts = [(round(px_), round(py)) for px_, py in tab_outline(tw, th)]
        arr = (wt.POINT * len(pts))(*[wt.POINT(a, b) for a, b in pts])
        rgn = CreatePolygonRgn(arr, len(pts), 2)  # WINDING
        SetWindowRgn(self.tab_hwnd, rgn, True)

    def _tab_hover(self, on):
        if not self.tab_down:
            self.tab.attributes("-alpha", 1.0 if on else float(self.s["TabOpacity"]))

    # ---- つまみ：クリック / 上下ドラッグ ----
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

    # ---- 状態遷移 ----
    def goto(self, target):
        if target == self.state:
            return
        self.state = target
        self.update_arrow()
        self.out_since = None

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

    def update_arrow(self):
        self.tab_label.configure(text="◀" if self.state == EXPANDED else "▶")

    # ---- アニメーション ----
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
            eased = 1 - (1 - t) ** 3       # ease-out cubic
            self.visible = frm + (to - frm) * eased
            self.layout()
            if t < 1.0:
                self.anim_id = self.root.after(10, step)
            else:
                self.anim_id = None
                if done:
                    done()

        step()

    # ---- 帯：クリック / ドラッグ ----
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
            self.goto(EXPANDED)                     # 覗き見中のクリックは固定
        else:
            self.toggle()

    def reset_width(self):
        self.s["Width"] = DEFAULTS["Width"]
        save_settings(self.s)
        if self.state != COLLAPSED:
            self.visible = self.W
            if self.state == EXPANDED:
                self.reserve(self.W + self.S)
            self.layout()

    # ---- 定期処理 ----
    def fast_poll(self):
        if self.closing:
            return
        try:
            # Ctrl+Alt+D
            pressed = all(GetAsyncKeyState(k) & 0x8000 for k in (VK_CONTROL, VK_MENU, VK_D))
            if pressed and not self.hotkey_prev:
                self.goto(EXPANDED) if self.state == PEEK else self.toggle()
            self.hotkey_prev = pressed

            p = wt.POINT()
            GetCursorPos(ctypes.byref(p))
            in_strip = inside(p, rect_of(self.strip_hwnd))   # 帯だけ（覗き見の開始判定）
            in_tab = inside(p, rect_of(self.tab_hwnd))
            now = time.time()

            # 帯へのホバーで覗き見（つまみへのホバーでは開かない）
            busy = self.down or self.tab_down
            if self.s["HoverPeek"] and self.state == COLLAPSED and not busy and in_strip:
                if self.hover_since is None:
                    self.hover_since = now
                elif now - self.hover_since > 0.25:
                    self.hover_since = None
                    self.goto(PEEK)
            else:
                self.hover_since = None

            # 覗き見中：外に 0.4 秒出たら収納
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
            if not IsWindow(self.edge):
                log.info("Edge window closed")
                self.quit()
                return
            old = self.insets
            self.update_insets()
            if self.insets != old and self.state != COLLAPSED and self.anim_id is None:
                self.layout()
            # Edge が勝手に動いていたら置き直す
            if (self.state != COLLAPSED and self.anim_id is None and not self.dragging
                    and self.expected and IsWindowVisible(self.edge)):
                r = rect_of(self.edge)
                ex = self.expected
                if (abs(r.left - ex[0]) > 2 or abs(r.top - ex[1]) > 2
                        or abs((r.right - r.left) - ex[2]) > 2):
                    log.info("Edge がずれていたので置き直し: %s -> %s",
                             (r.left, r.top, r.right - r.left), ex)
                    self.layout()
            m, w = primary_monitor()
            key = (m.left, m.top, m.right, m.bottom, w.top, w.bottom)
            if key != self.mon_key and self.anim_id is None and not self.dragging:
                self.reserve(self.reserved or self.S, force=True)
                self.layout()
        except Exception:
            log.exception("slow_poll")
        self.root.after(1000, self.slow_poll)

    # ---- 終了 ----
    def cleanup(self):
        if self.registered:
            self.registered = False
            try:
                SHAppBarMessage(ABM_REMOVE, ctypes.byref(self._abd()))
            except Exception:
                log.exception("ABM_REMOVE")

    def quit(self):
        if self.closing:
            return
        self.closing = True
        self.cancel_anim()
        self.cleanup()
        try:
            if IsWindow(self.edge):
                PostMessageW(self.edge, WM_CLOSE, 0, 0)
        except Exception:
            pass
        self.root.destroy()

    def _on_tk_error(self, exc, val, tb):
        log.error("tk callback error", exc_info=(exc, val, tb))

    def _fatal(self, msg):
        log.error(msg)
        messagebox.showerror("DashboardBar", msg)
        self.root.destroy()
        sys.exit(1)

    def run(self):
        try:
            self.root.mainloop()
        finally:
            self.cleanup()


if __name__ == "__main__":
    try:
        DashboardBar().run()
    except SystemExit:
        raise
    except Exception:
        log.exception("fatal")
        raise
