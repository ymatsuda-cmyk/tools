#!/usr/bin/env python3
"""
drive_inbox.py — Google Drive ミラーフォルダを監視して文字起こし→Notion登録

入り口は3つとも同じフォルダに集約される:
  1. minutes-viewer の upload.js   → 音声 + サイドカー JSON
  2. iOS 共有シート → ドライブ      → 音声のみ（ファイル名か更新日時から推定）
  3. Finder にドラッグ             → 同上

処理に成功した音声は削除する。失敗したものは failed/ に退避して原因を残す。

  crontab ではなく launchd から起動する（ユーザーコンテキストが必要）。
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
import time
import traceback
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

# ---------------------------------------------------------------- config

JST = timezone(timedelta(hours=9))

ENV_FILE = Path.home() / ".plaud_notion_sync.env"
if ENV_FILE.exists():
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        # setdefault なので既存の環境変数が優先される。
        # MINUTES_INDEX_PATH をここに書いているとスクリプト側の既定値を上書きする点に注意。
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

INBOX = Path(
    os.environ.get(
        "AUDIO_INBOX",
        str(Path.home() / "Google Drive/マイドライブ/audio-inbox"),
    )
)
FAILED_DIR = INBOX / "failed"
LOCK_FILE = Path("/tmp/drive_inbox.lock")

NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
NOTION_DB_ID = os.environ.get("NOTION_DB_ID", "28b0e7a535dc805697c6d4b9f8032d18")
NOTION_VERSION = "2022-06-28"

# Notion のプロパティ名。DB 側の実際の名前に合わせて変更する。
PROP_TITLE = os.environ.get("PROP_TITLE", "名前")
PROP_DATE = os.environ.get("PROP_DATE", "日付")
PROP_CATEGORY = os.environ.get("PROP_CATEGORY", "カテゴリー")
PROP_PERMISSION = os.environ.get("PROP_PERMISSION", "権限")
PROP_STATUS = os.environ.get("PROP_STATUS", "状態")
STATUS_DONE = os.environ.get("STATUS_DONE", "文字起こし")

AUDIO_EXT = {".m4a", ".mp3", ".wav", ".mp4", ".aac", ".flac", ".ogg"}
SKIP_EXT = {".tmp", ".partial", ".download", ".crdownload", ".gdoc"}
STABLE_SECONDS = int(os.environ.get("STABLE_SECONDS", "60"))
BLOCK_LIMIT = 1900  # rich_text は 2000 文字まで。改行の都合で少し余裕を持たせる
DRY_RUN = os.environ.get("DRY_RUN") == "1"
KEEP_DONE = os.environ.get("KEEP_DONE") == "1"  # 1 なら削除せず done/ へ移動

# 20260903-1400_定例_a1b2  /  20260903-140000-定例  のどちらも拾う
FNAME_RE = re.compile(r"^(\d{8})-(\d{4})(?:\d{2})?[_-](.*)$")


def log(*a):
    print(f"[{datetime.now(JST):%Y-%m-%d %H:%M:%S}]", *a, flush=True)


# ---------------------------------------------------------------- collect


def is_stable(p: Path) -> bool:
    """Drive のミラーは書き込み完了前にファイルが見えることがある。
    更新から一定時間経ち、サイズが変化しないことを確認する。"""
    try:
        s1 = p.stat()
    except FileNotFoundError:
        return False
    if time.time() - s1.st_mtime < STABLE_SECONDS:
        return False
    time.sleep(2)
    try:
        return p.stat().st_size == s1.st_size and s1.st_size > 0
    except FileNotFoundError:
        return False


def collect() -> list[tuple[Path, dict, Path | None]]:
    """(音声, メタデータ, サイドカーJSON or None) のリストを返す"""
    if not INBOX.is_dir():
        log(f"!! inbox が見つかりません: {INBOX}")
        log("   Drive for desktop がミラーモードで同期されているか確認してください。")
        return []

    jobs: list[tuple[Path, dict, Path | None]] = []
    claimed: set[str] = set()

    # 1) サイドカー JSON があるもの
    for j in sorted(INBOX.glob("*.json")):
        try:
            meta = json.loads(j.read_text(encoding="utf-8"))
        except Exception as e:
            log(f"-- JSON が壊れています: {j.name} ({e})")
            continue
        name = meta.get("audio")
        if not name:
            log(f"-- audio キーがありません: {j.name}")
            continue
        claimed.add(name)
        audio = INBOX / name
        if not audio.exists():
            # 音声が先に消えている（前回の削除漏れ）。JSON だけ掃除する。
            log(f"-- 音声不在。JSON を削除: {j.name}")
            j.unlink()
            continue
        if is_stable(audio):
            jobs.append((audio, meta, j))

    # 2) JSON なし → ファイル名か更新日時から日時を推定
    for a in sorted(INBOX.iterdir()):
        if not a.is_file() or a.name in claimed:
            continue
        ext = a.suffix.lower()
        if ext in SKIP_EXT or ext not in AUDIO_EXT or a.name.startswith("."):
            continue
        if not is_stable(a):
            continue
        m = FNAME_RE.match(a.stem)
        if m:
            dt = datetime.strptime(m[1] + m[2], "%Y%m%d%H%M").replace(tzinfo=JST)
            title = m[3].replace("_", " ").strip() or a.stem
        else:
            dt = datetime.fromtimestamp(a.stat().st_mtime, JST)
            title = a.stem
        jobs.append((a, {"title": title, "meetingAt": dt.isoformat(), "source": "drive"}, None))

    return jobs


# ---------------------------------------------------------------- transcribe


def transcribe(audio: Path) -> str:
    """既存の transcribe_engines を使う。エンジン切替も既存設定に従う。"""
    sys.path.insert(0, str(Path.home() / "scripts"))
    from transcribe_engines import transcribe as _t  # type: ignore

    # Qwen3-ASR は diarize を省略すると暗黙的に True になる。必ず明示する。
    return _t(str(audio), diarize=False)


# ---------------------------------------------------------------- notion


def notion(method: str, path: str, payload: dict | None = None) -> dict:
    if not NOTION_TOKEN:
        raise RuntimeError("NOTION_TOKEN が未設定です")
    res = requests.request(
        method,
        f"https://api.notion.com/v1/{path}",
        headers={
            "Authorization": f"Bearer {NOTION_TOKEN}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=60,
    )
    if not res.ok:
        raise RuntimeError(f"Notion {method} {path} -> {res.status_code}: {res.text[:400]}")
    return res.json()


def chunk_text(text: str, size: int = BLOCK_LIMIT) -> list[str]:
    """改行を優先しつつ size 以下に分割する"""
    out: list[str] = []
    buf = ""
    for para in text.replace("\r\n", "\n").split("\n"):
        while len(para) > size:
            out.append(para[:size])
            para = para[size:]
        if len(buf) + len(para) + 1 > size:
            if buf:
                out.append(buf)
            buf = para
        else:
            buf = f"{buf}\n{para}" if buf else para
    if buf:
        out.append(buf)
    return out or [""]


def para(text: str) -> dict:
    return {
        "object": "block",
        "type": "paragraph",
        "paragraph": {"rich_text": [{"type": "text", "text": {"content": text}}]},
    }


def create_page(meta: dict, transcript: str) -> str:
    props: dict = {
        PROP_TITLE: {"title": [{"text": {"content": meta.get("title", "無題")[:200]}}]},
        PROP_DATE: {"date": {"start": meta["meetingAt"]}},
        PROP_STATUS: {"status": {"name": STATUS_DONE}},
    }
    cats = meta.get("category") or []
    if isinstance(cats, str):
        cats = [cats]
    if cats:
        props[PROP_CATEGORY] = {"multi_select": [{"name": c} for c in cats]}
    if meta.get("permission"):
        props[PROP_PERMISSION] = {"select": {"name": meta["permission"]}}

    children: list[dict] = []
    if meta.get("note"):
        children.append(
            {
                "object": "block",
                "type": "callout",
                "callout": {
                    "rich_text": [{"type": "text", "text": {"content": meta["note"][:BLOCK_LIMIT]}}],
                    "icon": {"emoji": "📝"},
                },
            }
        )
    children.append(
        {
            "object": "block",
            "type": "heading_2",
            "heading_2": {"rich_text": [{"type": "text", "text": {"content": "文字起こし"}}]},
        }
    )

    chunks = chunk_text(transcript)
    page = notion(
        "POST",
        "pages",
        {
            "parent": {"database_id": NOTION_DB_ID},
            "properties": props,
            # ページ作成時に渡せる children は 100 件まで
            "children": children + [para(c) for c in chunks[:100 - len(children)]],
        },
    )
    page_id = page["id"]

    rest = chunks[100 - len(children):]
    for i in range(0, len(rest), 100):
        notion("PATCH", f"blocks/{page_id}/children", {"children": [para(c) for c in rest[i : i + 100]]})

    return page_id


def refresh_index() -> None:
    """index.json を再生成して GitHub Pages に反映する。
    既存スクリプトに関数があればそれを呼ぶ。無ければ手動実行を促す。"""
    sys.path.insert(0, str(Path.home() / "scripts"))
    try:
        import plaud_transcribe_notion as p  # type: ignore
    except Exception as e:
        log(f"-- index 再生成をスキップ（import 失敗: {e}）")
        return
    for name in ("build_index", "generate_index", "rebuild_index", "write_index"):
        fn = getattr(p, name, None)
        if callable(fn):
            fn()
            log(f"-- index.json を再生成しました（{name}）")
            return
    log("-- index 再生成の関数が見つかりません。plaud_transcribe_notion.py の関数名を refresh_index() に追記してください。")


# ---------------------------------------------------------------- main


def finish(audio: Path, sidecar: Path | None) -> None:
    if KEEP_DONE:
        done = INBOX / "done"
        done.mkdir(exist_ok=True)
        shutil.move(str(audio), done / audio.name)
        if sidecar:
            shutil.move(str(sidecar), done / sidecar.name)
        log(f"-- done/ へ移動: {audio.name}")
    else:
        audio.unlink(missing_ok=True)
        if sidecar:
            sidecar.unlink(missing_ok=True)
        log(f"-- 削除: {audio.name}")


def fail(audio: Path, sidecar: Path | None, err: str) -> None:
    FAILED_DIR.mkdir(exist_ok=True)
    shutil.move(str(audio), FAILED_DIR / audio.name)
    if sidecar:
        shutil.move(str(sidecar), FAILED_DIR / sidecar.name)
    (FAILED_DIR / f"{audio.stem}.error.txt").write_text(err, encoding="utf-8")
    log(f"!! failed/ へ退避: {audio.name}")


def main() -> int:
    # launchd の実行が重なると同じファイルを二重処理するので排他をかける
    if LOCK_FILE.exists():
        age = time.time() - LOCK_FILE.stat().st_mtime
        if age < 7200:
            log(f"-- 前回の実行が継続中（{int(age)}秒）。終了します。")
            return 0
        log("-- 古いロックを破棄します")
    LOCK_FILE.write_text(str(os.getpid()))

    try:
        jobs = collect()
        if not jobs:
            log("処理対象なし")
            return 0
        log(f"{len(jobs)} 件を処理します")

        ok = 0
        for audio, meta, sidecar in jobs:
            log(f">> {audio.name}  ({audio.stat().st_size / 1048576:.1f} MB)  {meta.get('meetingAt')}")
            if DRY_RUN:
                log("   DRY_RUN のため実処理はしません")
                continue
            try:
                t0 = time.time()
                text = transcribe(audio)
                log(f"   文字起こし完了 {len(text):,} 文字 / {time.time() - t0:.0f}秒")
                if not text.strip():
                    raise RuntimeError("文字起こし結果が空です")
                page_id = create_page(meta, text)
                log(f"   Notion 登録完了 {page_id}")
                # Notion 登録が成功してからファイルを消す。順序を逆にしないこと。
                finish(audio, sidecar)
                ok += 1
            except Exception:
                fail(audio, sidecar, traceback.format_exc())

        if ok:
            refresh_index()
        log(f"完了: {ok}/{len(jobs)} 件")
        return 0
    finally:
        LOCK_FILE.unlink(missing_ok=True)


if __name__ == "__main__":
    sys.exit(main())
