#!/usr/bin/env python3
"""監視フォルダに置かれた文書をNotionのコンテンツDBへ取り込む。

置かれたファイルから本文を抜き出してNotionにページを作り、済んだファイルは
old/ へ移す。移し終えたものは二度と読まないので、文字起こしのやり直しは無い
(作り直したいときは old/ から戻す)。

対応する形式は extract_text.py を参照(PDF / テキスト / Word / Excel / PowerPoint)。

環境変数 (既定で ~/.contentsstock.env からも読む):
  NOTION_TOKEN        Notion Integration Token(必須)
  CONTENTS_DB_ID      コンテンツDBのID(既定は下の DEFAULT_DB_ID)
  CONTENTS_WATCH_DIR  監視するフォルダ(必須)
  CONTENTS_DONE_DIR   済んだファイルの置き場(既定: 監視フォルダ/old)
  CONTENTS_ENV_FILE   環境変数ファイルのパス

使い方:
    python3 contents_watch.py --once            # cron向け。1回だけ見る
    python3 contents_watch.py --interval 300    # 常駐して5分ごとに見る
    python3 contents_watch.py --once --dry-run  # 対象の確認だけ
"""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract_text import extract, is_supported  # noqa: E402

JST = timezone(timedelta(hours=9))
ENV_FILE = Path(os.environ.get("CONTENTS_ENV_FILE", str(Path.home() / ".contentsstock.env")))


def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env()

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
DEFAULT_DB_ID = "d600e7a535dc83caadf381afe7abea03"
CONTENTS_DB_ID = os.environ.get("CONTENTS_DB_ID", DEFAULT_DB_ID)
WATCH_DIR = os.environ.get("CONTENTS_WATCH_DIR", "")
DONE_DIR = os.environ.get("CONTENTS_DONE_DIR", "")

# Notion側のカラム名。gas/Code.gs や build_contentsstock_json.py と揃えること
PROP_TITLE = "タイトル"
PROP_FILE = "ファイル名"
PROP_KIND = "種別"
PROP_STATUS = "状態"
PROP_RAW_COUNT = "原文文字数"
STATUS_DONE = "完了"  # 本文が入った状態。ここからアプリ側で要約する

# 書き込み中のファイルを掴まないための猶予(秒)
SETTLE_SECONDS = int(os.environ.get("CONTENTS_SETTLE_SECONDS", "20"))


def log(msg):
    print(f"[{datetime.now(JST).strftime('%H:%M:%S')}] {msg}", flush=True)


# ---------------------------------------------------------------- Notion


def headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def notion(method, path, payload=None):
    resp = requests.request(method, f"{NOTION_API}/{path}", headers=headers(), json=payload, timeout=120)
    if resp.status_code >= 300:
        raise RuntimeError(f"Notion API {resp.status_code}: {resp.text[:300]}")
    return resp.json()


def db_property_types(db_id):
    """無いカラムへ書こうとすると400になるので、先に形を見ておく。"""
    data = notion("get", f"databases/{db_id}")
    return {name: prop.get("type") for name, prop in (data.get("properties") or {}).items()}


def title_prop_of(types):
    if types.get(PROP_TITLE) == "title":
        return PROP_TITLE
    for name, kind in types.items():
        if kind == "title":
            return name
    raise RuntimeError("コンテンツDBにtitle型のカラムがありません")


def already_registered(db_id, types, filename):
    """同じファイル名が既にあれば取り込まない。old/ から戻したときの二重登録を防ぐ。"""
    prop = PROP_FILE if types.get(PROP_FILE) == "rich_text" else None
    if not prop:
        return False
    data = notion("post", f"databases/{db_id}/query",
                  {"page_size": 1, "filter": {"property": prop, "rich_text": {"equals": filename}}})
    return bool(data.get("results"))


def text_to_blocks(text):
    """2000文字ごとに分割した段落ブロックへ変換する。"""
    blocks = []
    for para in [p.strip() for p in str(text).split("\n") if p.strip()]:
        while para:
            chunk, para = para[:2000], para[2000:]
            blocks.append({"object": "block", "type": "paragraph",
                           "paragraph": {"rich_text": [{"type": "text", "text": {"content": chunk}}]}})
    return blocks


def create_page(db_id, types, title_prop, path, text, engine):
    props = {title_prop: {"title": [{"text": {"content": path.stem[:2000]}}]}}
    if types.get(PROP_FILE) == "rich_text":
        props[PROP_FILE] = {"rich_text": [{"text": {"content": path.name[:2000]}}]}
    if types.get(PROP_KIND) == "select":
        props[PROP_KIND] = {"select": {"name": path.suffix.lower().lstrip(".") or "その他"}}
    if types.get(PROP_STATUS) == "select":
        props[PROP_STATUS] = {"select": {"name": STATUS_DONE}}
    if types.get(PROP_RAW_COUNT) == "number":
        props[PROP_RAW_COUNT] = {"number": len(text)}

    ts = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    head = [
        {"object": "block", "type": "heading_2",
         "heading_2": {"rich_text": [{"type": "text", "text": {"content": f"📄 {path.name}"}}]}},
        {"object": "block", "type": "callout",
         "callout": {"icon": {"type": "emoji", "emoji": "📝"},
                     "rich_text": [{"type": "text",
                                    "text": {"content": f"取り込み: {ts}（抽出方法: {engine}）"}}]}},
    ]
    blocks = head + text_to_blocks(text)

    page = notion("post", "pages", {
        "parent": {"database_id": db_id},
        "properties": props,
        "children": blocks[:100],
    })
    rest = blocks[100:]
    for i in range(0, len(rest), 100):
        notion("patch", f"blocks/{page['id']}/children", {"children": rest[i:i + 100]})
    return page


# ---------------------------------------------------------------- ファイルの扱い


def pending_files(watch_dir, done_dir):
    """監視フォルダ直下の対象ファイル。old/ と隠しファイルは見ない。"""
    out = []
    for path in sorted(Path(watch_dir).iterdir()):
        if not path.is_file() or path.name.startswith("."):
            continue
        if done_dir in path.parents:
            continue
        if not is_supported(path):
            continue
        out.append(path)
    return out


def settled(path):
    """コピー中のファイルを半端に読まないよう、更新が止まってから扱う。"""
    return time.time() - path.stat().st_mtime >= SETTLE_SECONDS


def move_to_done(path, done_dir):
    done_dir.mkdir(parents=True, exist_ok=True)
    dest = done_dir / path.name
    if dest.exists():
        stamp = datetime.now(JST).strftime("%Y%m%d-%H%M%S")
        dest = done_dir / f"{path.stem}_{stamp}{path.suffix}"
    path.rename(dest)
    return dest


# ---------------------------------------------------------------- 本処理


def process(path, db_id, types, title_prop, done_dir, dry_run):
    if already_registered(db_id, types, path.name):
        log(f"  ⏭ Notionに同じファイル名があります: {path.name}")
        if not dry_run:
            move_to_done(path, done_dir)
        return "dup"

    if dry_run:
        log(f"  → 取り込み対象: {path.name}")
        return "ok"

    text, engine = extract(path, log=lambda m: log(m))
    if text is None:
        log(f"  ❌ {path.name}: {engine}")
        return "error"
    if not text.strip():
        log(f"  ⚠️ {path.name}: 本文を取り出せませんでした（中身が空？）")
        return "error"

    create_page(db_id, types, title_prop, path, text, engine)
    dest = move_to_done(path, done_dir)
    log(f"  ✅ {path.name} → Notion ({len(text):,}字 / {engine})  移動先: {dest.parent.name}/")
    return "ok"


def run_once(db_id, types, title_prop, watch_dir, done_dir, dry_run):
    files = pending_files(watch_dir, done_dir)
    if not files:
        return
    waiting = [p for p in files if not settled(p)]
    ready = [p for p in files if settled(p)]
    if waiting:
        log(f"{len(waiting)}件はまだ書き込み中かもしれないので次回に回します")
    if not ready:
        return

    log(f"{len(ready)}件を取り込みます" + ("（--dry-run）" if dry_run else ""))
    counts = {"ok": 0, "dup": 0, "error": 0}
    for path in ready:
        try:
            counts[process(path, db_id, types, title_prop, done_dir, dry_run)] += 1
        except Exception as e:  # noqa: BLE001  1件失敗しても残りは続ける
            log(f"  ❌ {path.name}: {type(e).__name__}: {e}")
            counts["error"] += 1
    log(f"完了: 取り込み {counts['ok']}件 / 重複 {counts['dup']}件 / エラー {counts['error']}件")


def main():
    ap = argparse.ArgumentParser(description="監視フォルダの文書をNotionへ取り込む")
    ap.add_argument("--dir", default=WATCH_DIR, help="監視するフォルダ")
    ap.add_argument("--done-dir", default=DONE_DIR, help="済んだファイルの置き場(既定: 監視フォルダ/old)")
    ap.add_argument("--db-id", default=CONTENTS_DB_ID, help="コンテンツDBのID")
    ap.add_argument("--once", action="store_true", help="1回だけ見て終わる(cron向け)")
    ap.add_argument("--interval", type=int, default=300, help="常駐時の確認間隔(秒)")
    ap.add_argument("--dry-run", action="store_true", help="対象を表示するだけ")
    args = ap.parse_args()

    if not NOTION_TOKEN:
        print("❌ NOTION_TOKEN が設定されていません", file=sys.stderr)
        return 1
    if not args.dir:
        print(f"❌ 監視フォルダが未設定です(CONTENTS_WATCH_DIR か --dir。{ENV_FILE} でも可)", file=sys.stderr)
        return 1

    watch_dir = Path(args.dir).expanduser()
    if not watch_dir.is_dir():
        print(f"❌ 監視フォルダがありません: {watch_dir}", file=sys.stderr)
        return 1
    done_dir = Path(args.done_dir).expanduser() if args.done_dir else watch_dir / "old"

    try:
        types = db_property_types(args.db_id)
        title_prop = title_prop_of(types)
    except Exception as e:  # noqa: BLE001
        print(f"❌ コンテンツDBを読めませんでした: {e}\n"
              "   NOTION_TOKEN と CONTENTS_DB_ID、DBが統合に共有されているかを確認してください",
              file=sys.stderr)
        return 1
    log(f"コンテンツDBのカラム: {', '.join(sorted(types))}")

    if args.once:
        run_once(args.db_id, types, title_prop, watch_dir, done_dir, args.dry_run)
        return 0

    log(f"監視開始: {watch_dir}（{args.interval}秒ごと / 済んだファイルは {done_dir}）")
    while True:
        try:
            run_once(args.db_id, types, title_prop, watch_dir, done_dir, args.dry_run)
        except Exception as e:  # noqa: BLE001  常駐は止めない
            log(f"⚠️ {type(e).__name__}: {e}")
        time.sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
