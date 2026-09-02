"""
quota_client.py — Kaggleノートブック側の稼働時間レポーター

既存の gas('started', ...) / heartbeat 処理に組み込んで使う。
GAS 側の kaggle_quota.gs が消費時間を積み上げるための「打刻」を送る。

【重要】
  ・GAS への HTTP は Python デフォルト User-Agent が Google にブロック
    されるので必ず User-Agent を差し替える
  ・打刻失敗でセッション全体を落とさないこと（例外は握り潰す）
  ・heartbeat が STALE_HEARTBEAT_SEC (既定300s) 途絶えると GAS 側が
    セッションを締めるので、間隔は 60s 程度に保つ
"""

import atexit
import json
import os
import re
import signal
import threading
import time
import urllib.parse
import urllib.request

GAS_URL = "https://script.google.com/macros/s/XXXXXXXXXXXX/exec"
KERNEL_SLUG = "matsuda2026/qwen3-coder-30b"
ACCEL = "GPU"
HEARTBEAT_SEC = 60
UA = "Mozilla/5.0 (X11; Linux x86_64) kaggle-quota-client/1.0"


def detect_session_id() -> str:
    """コンテナ名からセッションIDを抽出。

    ダッシュ区切りの5桁以上のセグメントだけを対象にする
    （文字列内の最初の数字を拾うと誤検出する）。
    """
    candidates = []
    try:
        with open("/proc/self/cgroup", "r") as f:
            candidates.append(f.read())
    except Exception:
        pass
    candidates.append(os.environ.get("HOSTNAME", ""))

    for text in candidates:
        for seg in re.split(r"[-/_.]", text):
            if re.fullmatch(r"\d{5,}", seg):
                return seg
    return str(int(time.time()))


SESSION_ID = detect_session_id()


def _call(action: str, **params) -> dict:
    """GAS を叩く。失敗しても例外を投げない。"""
    params.update({"action": action, "session_id": SESSION_ID})
    url = GAS_URL + "?" + urllib.parse.urlencode(
        {k: v for k, v in params.items() if v is not None}
    )
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    for attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=30) as res:
                body = res.read().decode("utf-8", "replace")
            return json.loads(body)
        except Exception as e:
            if attempt == 4:
                print(f"[QUOTA] {action} failed: {e}", flush=True)
                return {"ok": False, "error": str(e)}
            time.sleep(2 * (attempt + 1))
    return {"ok": False}


def quota_start() -> dict:
    r = _call("quota_session_start", kernel=KERNEL_SLUG, accel=ACCEL)
    print(f"[QUOTA] start {SESSION_ID} -> {r}", flush=True)
    return r


def quota_beat() -> dict:
    return _call("quota_heartbeat", kernel=KERNEL_SLUG, accel=ACCEL)


def quota_end(note: str = "exit") -> dict:
    r = _call("quota_end", note=note)
    print(f"[QUOTA] end {SESSION_ID} -> {r}", flush=True)
    return r


def quota_remaining() -> dict:
    """自分の残量を取得（ノートブック内で自己停止判断に使える）。"""
    return _call("quota", accel=ACCEL)


_stop = threading.Event()


def _loop():
    while not _stop.wait(HEARTBEAT_SEC):
        quota_beat()


def start_reporter(auto_stop_below_hours: float = 0.0):
    """打刻を開始。auto_stop_below_hours > 0 なら残量が切れたら自己終了。"""
    quota_start()

    def runner():
        while not _stop.wait(HEARTBEAT_SEC):
            quota_beat()
            if auto_stop_below_hours > 0:
                st = quota_remaining()
                rem = st.get("remaining_hours")
                if isinstance(rem, (int, float)) and rem <= auto_stop_below_hours:
                    print(f"[QUOTA] remaining {rem}h — shutting down", flush=True)
                    quota_end("auto-stop: quota exhausted")
                    os._exit(0)

    threading.Thread(target=runner, daemon=True).start()

    def _bye(*_a):
        _stop.set()
        quota_end("exit")

    atexit.register(_bye)
    for sig in (signal.SIGTERM, signal.SIGINT):
        try:
            signal.signal(sig, lambda s, f: (_bye(), os._exit(0)))
        except Exception:
            pass


if __name__ == "__main__":
    start_reporter()
    print(json.dumps(quota_remaining(), indent=2, ensure_ascii=False))
    time.sleep(180)
