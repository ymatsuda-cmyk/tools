#!/usr/bin/env python3
"""video_inbox.py — Drive の inbox に置かれた動画を文字起こしして Notion へ登録する。

コンテンツナレッジ(beta/contentsstock)の取り込み口。文書側の contents_watch.py と
役割は同じで、扱うものが動画・音声になっている。

処理の流れ:
  1. Drive for desktop がミラーした inbox フォルダを見る
  2. 書き込み中でないことを確かめる(サイズが安定するまで待つ)
  3. mlx-whisper で文字起こしする([mm:ss] 付きの行にする)
  4. Notion にページを作り、本文へ文字起こしを流し込む
  5. 済んだ動画は contents フォルダへ移す(同じDrive内なのでリンクは変わらない)

アプリ(ブラウザ)からアップロードした場合は、同じ名前の .json が一緒に置かれる。
中に タイトル・タグ・DriveファイルID が入っているので、それを優先して使う。
JSONが無い(Finderから直接置いた)場合はファイル名をタイトルにする。

環境変数 (既定で ~/.contentsstock.env からも読む):
  NOTION_TOKEN          Notion Integration Token(必須)
  CONTENTS_DB_ID        コンテンツDBのID(既定は下の DEFAULT_DB_ID)
  CONTENTS_INBOX        監視するフォルダ(必須。Driveミラーの inbox)
  CONTENTS_STORE        済んだ動画の置き場(既定: inboxの隣の contents)
  CONTENTS_ENV_FILE     環境変数ファイルのパス
  WHISPER_MODEL         既定: mlx-community/whisper-large-v2-mlx
  WHISPER_LANGUAGE      既定: ja

使い方:
    python3 video_inbox.py --once            # cron / launchd 向け。1回だけ見る
    python3 video_inbox.py --interval 300    # 常駐して5分ごとに見る
    python3 video_inbox.py --once --dry-run  # 対象の確認だけ
"""
from __future__ import annotations

import argparse
import fcntl
import json
import os
import shutil
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

JST = timezone(timedelta(hours=9))
ENV_FILE = Path(os.environ.get("CONTENTS_ENV_FILE", str(Path.home() / ".contentsstock.env")))


def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_env()

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
DEFAULT_DB_ID = "d600e7a535dc83caadf381afe7abea03"
CONTENTS_DB_ID = os.environ.get("CONTENTS_DB_ID", DEFAULT_DB_ID)

INBOX = Path(os.environ.get("CONTENTS_INBOX", str(Path.home() / "Google Drive/マイドライブ/contents-inbox")))
STORE = Path(os.environ.get("CONTENTS_STORE", str(INBOX.parent / "contents")))
FAILED = INBOX / "failed"
LOCK_FILE = Path(os.environ.get("CONTENTS_LOCK_FILE", str(Path.home() / ".contentsstock_video_inbox.lock")))

WHISPER_MODEL = os.environ.get("WHISPER_MODEL", "mlx-community/whisper-large-v2-mlx")
WHISPER_LANGUAGE = os.environ.get("WHISPER_LANGUAGE", "ja")

# Notion側のカラム名。gas/Code.gs と揃えること
PROP_TITLE = "タイトル"
PROP_FILE = "ファイル名"
PROP_DRIVE = "Driveリンク"
PROP_KIND = "種別"
PROP_TAGS = "タグ"
PROP_STATUS = "状態"
PROP_RAW_COUNT = "原文文字数"
STATUS_DONE = "完了"  # 本文が入った状態。ここからアプリ側で要約する

VIDEO_EXT = {".mp4", ".mov", ".m4v", ".webm", ".mkv", ".avi", ".mp3", ".m4a", ".wav", ".aac"}
SKIP_EXT = {".tmp", ".partial", ".download", ".crdownload"}
STABLE_SECONDS = int(os.environ.get("CONTENTS_STABLE_SECONDS", "60"))
BLOCK_LIMIT = 1900  # rich_text は2000文字まで。改行の都合で少し余裕を持たせる


def log(*a):
    print(f"[{datetime.now(JST):%Y-%m-%d %H:%M:%S}]", *a, flush=True)


def acquire_lock():
    """多重起動を防ぐ。

    常駐と手動実行が重なると同じ動画を両方が拾い、先に終わった側が
    ファイルを移すため、もう片方が FileNotFoundError で落ちる。
    文字起こしは数分かかるので、interval より長引くと自分自身とも重なる。
    このハンドルはプロセスが終わるまで持ち続けること。
    """
    handle = open(LOCK_FILE, "w")  # noqa: SIM115
    try:
        fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        handle.close()
        return None
    handle.write(str(os.getpid()))
    handle.flush()
    return handle


# ---------------------------------------------------------------- 収集


def is_stable(path: Path) -> bool:
    """Driveのミラーは書き込み完了前にファイルが見えることがある。
    更新から一定時間経ち、サイズが変化しないことを確認する。"""
    try:
        first = path.stat()
    except FileNotFoundError:
        return False
    if time.time() - first.st_mtime < STABLE_SECONDS:
        return False
    time.sleep(2)
    try:
        return path.stat().st_size == first.st_size and first.st_size > 0
    except FileNotFoundError:
        return False


def collect() -> list[tuple[Path, dict, Path | None]]:
    """(動画, メタデータ, サイドカーJSON or None) のリストを返す"""
    if not INBOX.is_dir():
        log(f"!! inbox が見つかりません: {INBOX}")
        log("   Drive for desktop がミラーモードで同期されているか確認してください。")
        return []

    jobs: list[tuple[Path, dict, Path | None]] = []
    claimed: set[str] = set()

    # 1) アプリからのアップロード(サイドカーJSONがある)
    for sidecar in sorted(INBOX.glob("*.json")):
        try:
            meta = json.loads(sidecar.read_text(encoding="utf-8"))
        except Exception as e:  # noqa: BLE001  書き込み途中のJSONなど
            log(f"-- JSONを読めません({sidecar.name}): {e}")
            continue
        video = INBOX / str(meta.get("filename") or "")
        if not video.is_file():
            continue
        if not is_stable(video):
            log(f"-- まだ書き込み中のようです: {video.name}")
            continue
        claimed.add(video.name)
        jobs.append((video, meta, sidecar))

    # 2) Finder などから直接置かれたもの
    for video in sorted(INBOX.iterdir()):
        if not video.is_file() or video.name in claimed:
            continue
        ext = video.suffix.lower()
        if ext in SKIP_EXT or ext not in VIDEO_EXT:
            continue
        if not is_stable(video):
            log(f"-- まだ書き込み中のようです: {video.name}")
            continue
        jobs.append((video, {"title": video.stem}, None))

    return jobs


# ---------------------------------------------------------------- 文字起こし


def format_timestamp(seconds: float) -> str:
    total = int(seconds)
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def resolve_ffmpeg_cmd():
    """mlx-whisper は内部で PATH から "ffmpeg" を探す。

    launchd / cron の PATH は /usr/bin:/bin だけなので、Homebrew に入れていても
    見つからず FileNotFoundError になる。見つけた場所を PATH に通しておく。
    """
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg
    for candidate in (Path("/opt/homebrew/bin/ffmpeg"), Path("/usr/local/bin/ffmpeg")):
        if candidate.exists() and os.access(candidate, os.X_OK):
            os.environ["PATH"] = str(candidate.parent) + os.pathsep + os.environ.get("PATH", "")
            return str(candidate)
    return None


def transcribe(video: Path) -> str:
    """mlx-whisper で文字起こしし、[mm:ss] 付きの行にする。

    動画のままでも mlx-whisper が内部の ffmpeg で音声を取り出すので、
    こちらで変換はしない。
    """
    try:
        import mlx_whisper
    except ImportError as e:  # noqa: F841
        raise RuntimeError("mlx-whisper が未導入です (pip install mlx-whisper)")

    if not resolve_ffmpeg_cmd():
        raise RuntimeError("ffmpeg が見つかりません (brew install ffmpeg)")

    result = mlx_whisper.transcribe(
        str(video),
        path_or_hf_repo=WHISPER_MODEL,
        language=WHISPER_LANGUAGE,
        condition_on_previous_text=False,
        verbose=False,
    )

    segments = [s for s in (result.get("segments") or []) if (s.get("text") or "").strip()]
    if segments:
        return "\n".join(
            f"[{format_timestamp(s.get('start') or 0.0)}] {(s.get('text') or '').strip()}" for s in segments
        )
    return (result.get("text") or "").strip()


# ---------------------------------------------------------------- Notion


def notion(method: str, path: str, payload: dict | None = None) -> dict:
    if not NOTION_TOKEN:
        raise RuntimeError("NOTION_TOKEN が未設定です")
    res = requests.request(
        method,
        f"{NOTION_API}/{path}",
        headers={
            "Authorization": f"Bearer {NOTION_TOKEN}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=120,
    )
    if not res.ok:
        raise RuntimeError(f"Notion {method} {path} -> {res.status_code}: {res.text[:400]}")
    return res.json()


def chunk_text(text: str, size: int = BLOCK_LIMIT) -> list[str]:
    """改行を優先しつつ size 以下に分割する"""
    out: list[str] = []
    buf = ""
    for line in text.split("\n"):
        while len(line) > size:  # 1行が長すぎる場合はそこで切る
            out.append(line[:size])
            line = line[size:]
        if len(buf) + len(line) + 1 > size:
            if buf:
                out.append(buf)
            buf = line
        else:
            buf = f"{buf}\n{line}" if buf else line
    if buf:
        out.append(buf)
    return out


_schema_cache: dict | None = None


def db_schema() -> dict:
    """DBのプロパティ定義。無いカラムを送るとNotionが400を返すため先に確かめる"""
    global _schema_cache
    if _schema_cache is None:
        _schema_cache = notion("GET", f"databases/{CONTENTS_DB_ID}").get("properties", {})
    return _schema_cache


def title_prop_name() -> str:
    """title型のカラム名。DBごとに名前が違うことがある"""
    for name, spec in db_schema().items():
        if spec.get("type") == "title":
            return name
    return PROP_TITLE


def keep_existing(props: dict) -> dict:
    """DBに無いカラムは落とす。任意のカラムが未作成でも取り込みだけは通す"""
    schema = db_schema()
    dropped = [k for k in props if k not in schema]
    if dropped:
        log(f"   -- DBに無いカラムは送りません: {', '.join(dropped)}")
    return {k: v for k, v in props.items() if k in schema}


def create_page(video: Path, meta: dict, text: str) -> str:
    title = str(meta.get("title") or video.stem)
    tags = [t for t in (meta.get("tags") or []) if t]
    drive_id = meta.get("driveFileId")

    props = {
        title_prop_name(): {"title": [{"text": {"content": title[:2000]}}]},
        PROP_FILE: {"rich_text": [{"text": {"content": video.name[:2000]}}]},
        PROP_KIND: {"select": {"name": video.suffix.lower().lstrip(".")}},
        PROP_STATUS: {"select": {"name": STATUS_DONE}},
        PROP_RAW_COUNT: {"number": len(text)},
    }
    if tags:
        props[PROP_TAGS] = {"multi_select": [{"name": str(t)[:100]} for t in tags]}
    if drive_id:
        props[PROP_DRIVE] = {"url": f"https://drive.google.com/file/d/{drive_id}/view"}
    props = keep_existing(props)

    children = [
        {
            "object": "block",
            "type": "paragraph",
            "paragraph": {"rich_text": [{"text": {"content": part}}]},
        }
        for part in chunk_text(text)[:100]  # 1リクエストは100ブロックまで
    ]

    page = notion("POST", "pages", {
        "parent": {"database_id": CONTENTS_DB_ID},
        "properties": props,
        "children": children,
    })

    # 100ブロックを超える分は追記する
    rest = chunk_text(text)[100:]
    for i in range(0, len(rest), 100):
        notion("PATCH", f"blocks/{page['id']}/children", {
            "children": [
                {
                    "object": "block",
                    "type": "paragraph",
                    "paragraph": {"rich_text": [{"text": {"content": part}}]},
                }
                for part in rest[i:i + 100]
            ]
        })

    return page["id"]


# ---------------------------------------------------------------- 実行


def move_to(path: Path, dest_dir: Path):
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / path.name
    if target.exists():
        target = dest_dir / f"{path.stem}_{int(time.time())}{path.suffix}"
    shutil.move(str(path), str(target))
    return target


def run_once(dry_run: bool = False) -> int:
    jobs = collect()
    if not jobs:
        log("対象はありません")
        return 0

    done = 0
    for video, meta, sidecar in jobs:
        log(f"== {video.name}")
        if dry_run:
            log(f"   (dry-run) タイトル: {meta.get('title') or video.stem}")
            continue
        text = ""
        try:
            text = transcribe(video)
            if not text.strip():
                raise RuntimeError("文字起こしの結果が空でした")
            page_id = create_page(video, meta, text)
            log(f"   Notionに登録しました({len(text)}字): {page_id}")
            move_to(video, STORE)
            if sidecar:
                sidecar.unlink(missing_ok=True)
            done += 1
        except Exception as e:  # noqa: BLE001  1件失敗しても次へ進む
            # 取り込み中に消えたなら、別のプロセスが処理し終えて移したとみなす
            if not video.exists():
                log("   -- 他で処理済みのようなので飛ばします")
                continue
            log(f"   !! 失敗: {type(e).__name__}: {e}")
            try:
                FAILED.mkdir(parents=True, exist_ok=True)
                # 文字起こしはやり直すと時間がかかるので、取れていれば残す
                if text.strip():
                    (FAILED / f"{video.stem}.txt").write_text(text, encoding="utf-8")
                move_to(video, FAILED)
                if sidecar:
                    move_to(sidecar, FAILED)
            except Exception as move_err:  # noqa: BLE001
                log(f"   !! 退避にも失敗: {move_err}")
    return done


def main():
    ap = argparse.ArgumentParser(description="Driveのinboxにある動画を文字起こししてNotionへ登録する")
    ap.add_argument("--once", action="store_true", help="1回だけ実行する")
    ap.add_argument("--interval", type=int, default=300, help="常駐時の確認間隔(秒)")
    ap.add_argument("--dry-run", action="store_true", help="対象の確認だけ行う")
    args = ap.parse_args()

    if not NOTION_TOKEN:
        log("!! NOTION_TOKEN が未設定です")
        return 1

    if not args.dry_run:
        lock = acquire_lock()
        if not lock:
            log("-- すでに別の取り込みが走っているので終了します")
            return 0

    if args.once or args.dry_run:
        run_once(args.dry_run)
        return 0

    log(f"監視を開始します: {INBOX} ({args.interval}秒ごと)")
    while True:
        try:
            run_once()
        except Exception as e:  # noqa: BLE001  常駐は落とさない
            log(f"!! 予期しないエラー: {type(e).__name__}: {e}")
        time.sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
