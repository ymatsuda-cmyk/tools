#!/usr/bin/env python3
"""Notionの動画DBを読み、動画ナレッジ(clipstock)の一覧用JSONを書き出す。

アプリは開くたびにGAS経由でNotionを全件クエリしていて、件数が増えるほど
最初の描画までが遅かった。cronでこのスクリプトを回してJSONを先に用意しておき、
画面はそれを読むだけにする。詳細(タブごとの本文)は従来どおりNotionから取るので、
ここでは一覧とアイデア一覧に要るものだけを書き出す。

出力:
  index.json  カード表示・検索・絞り込みに要る項目(長文は有無のフラグだけ)
  ideas.json  応用と活用アイデアの本文(アイデア一覧画面が使う)

環境変数:
  NOTION_TOKEN       Notion Integration Token(必須)
  VIDEO_ENV_FILE     環境変数を読み込むファイルのパス(既定: ~/.video_notion_sync.env)
  VIDEO_DB_ID        対象データベースID
  CLIPSTOCK_OUT_DIR  出力先ディレクトリ(--out より弱い)

使い方:
    python3 build_clipstock_json.py
    python3 build_clipstock_json.py --out ~/Claude/app/clipstock/data
"""
import argparse
import json
import os
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

JST = timezone(timedelta(hours=9))
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[2]

ENV_FILE = Path(os.environ.get("VIDEO_ENV_FILE", str(Path.home() / ".video_notion_sync.env")))


def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env()

NOTION_API = "https://api.notion.com/v1"
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
VIDEO_DB_ID = os.environ.get("VIDEO_DB_ID", "3630e7a535dc8154ac62d41f7611540f")

# Notion側のカラム名。gas/Code.gs の PROP_* と一致させること
PROP_TITLE = "動画タイトル"
PROP_URL = "URL"
PROP_THUMB = "サムネイル"
PROP_TAGS = "タグ"
PROP_STATUS = "状態"
PROP_SUMMARY = "要約"
PROP_MINDMAP = "マインドマップ"
PROP_FIELDS = "分野別要約"
PROP_APPLY = "応用"
PROP_IDEAS = "活用アイデア"
PROP_MEMO = "メモ"
PROP_MODEL = "要約モデル"
PROP_GENERATED = "要約日時"
PROP_RAW_COUNT = "原文文字数"
PROP_CREATED = "作成日時"

STATUS_NEW = "新規"

# ---------------------------------------------------------------- Notion


def notion_headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }


def query_all_pages():
    """DBの全ページを作成日時の新しい順で取得する。"""
    pages = []
    cursor = None
    while True:
        payload = {
            "page_size": 100,
            "sorts": [{"property": PROP_CREATED, "direction": "descending"}],
        }
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(
            f"{NOTION_API}/databases/{VIDEO_DB_ID}/query",
            headers=notion_headers(),
            json=payload,
            timeout=60,
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Notion API {resp.status_code}: {resp.text[:300]}")
        data = resp.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
    return pages


# ---------------------------------------------------------------- プロパティの読み出し


def plain_text(rich):
    return "".join(part.get("plain_text", "") for part in (rich or []))


def title_of(props, name):
    return plain_text((props.get(name) or {}).get("title"))


def rich_of(props, name):
    return plain_text((props.get(name) or {}).get("rich_text"))


def url_of(props, name):
    return (props.get(name) or {}).get("url") or ""


def select_of(props, name):
    sel = (props.get(name) or {}).get("select")
    return (sel or {}).get("name") or ""


def multi_select_of(props, name):
    return [o.get("name", "") for o in (props.get(name) or {}).get("multi_select") or []]


def number_of(props, name):
    value = (props.get(name) or {}).get("number")
    return value if isinstance(value, (int, float)) else 0


def date_of(props, name):
    date = (props.get(name) or {}).get("date")
    return (date or {}).get("start")


def created_of(page, props):
    prop = props.get(PROP_CREATED) or {}
    return prop.get("created_time") or page.get("created_time")


# ---------------------------------------------------------------- 変換


def to_item(page):
    """gas/Code.gs の listVideos_ と同じ形にする。"""
    p = page.get("properties", {})
    return {
        "key": page["id"],
        "title": title_of(p, PROP_TITLE) or "(タイトル未取得)",
        "url": url_of(p, PROP_URL),
        "thumb": url_of(p, PROP_THUMB),
        "status": select_of(p, PROP_STATUS) or STATUS_NEW,
        "tags": multi_select_of(p, PROP_TAGS),
        "createdAt": created_of(page, p),
        "editedAt": page.get("last_edited_time"),
        "summary": rich_of(p, PROP_SUMMARY),
        "model": rich_of(p, PROP_MODEL) or None,
        "generatedAt": date_of(p, PROP_GENERATED),
        "rawCount": number_of(p, PROP_RAW_COUNT),
        "has": {
            "mindmap": bool(rich_of(p, PROP_MINDMAP)),
            "fields": bool(rich_of(p, PROP_FIELDS)),
            "apply": bool(rich_of(p, PROP_APPLY)),
            "ideas": bool(rich_of(p, PROP_IDEAS)),
            "memo": bool(rich_of(p, PROP_MEMO)),
        },
    }


def to_idea(page):
    """gas/Code.gs の listIdeas_ と同じ形。応用も活用も無いページは None を返す。"""
    p = page.get("properties", {})
    apply_text = rich_of(p, PROP_APPLY)
    ideas_text = rich_of(p, PROP_IDEAS)
    if not apply_text and not ideas_text:
        return None
    return {
        "key": page["id"],
        "title": title_of(p, PROP_TITLE),
        "url": url_of(p, PROP_URL),
        "thumb": url_of(p, PROP_THUMB),
        "tags": multi_select_of(p, PROP_TAGS),
        "status": select_of(p, PROP_STATUS),
        "apply": apply_text,
        "ideas": ideas_text,
    }


# ---------------------------------------------------------------- 出力


def write_json(path, payload):
    """書き込み中のファイルを画面に読ませないよう、一時ファイルを作ってから差し替える。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def main():
    parser = argparse.ArgumentParser(description="動画ナレッジの一覧JSONを書き出す")
    parser.add_argument(
        "--out",
        default=os.environ.get("CLIPSTOCK_OUT_DIR", str(REPO_ROOT / "data" / "clipstock")),
        help="出力先ディレクトリ",
    )
    parser.add_argument("--dry-run", action="store_true", help="件数だけ表示して書き出さない")
    args = parser.parse_args()

    if not NOTION_TOKEN:
        print("NOTION_TOKEN が未設定です", file=sys.stderr)
        return 1

    pages = query_all_pages()
    items = [to_item(page) for page in pages]
    ideas = [idea for idea in (to_idea(page) for page in pages) if idea]
    generated_at = datetime.now(JST).isoformat()

    print(f"動画 {len(items)}件 / アイデアのあるもの {len(ideas)}件")
    if args.dry_run:
        return 0

    out_dir = Path(args.out).expanduser()
    write_json(out_dir / "index.json", {"generatedAt": generated_at, "items": items})
    write_json(out_dir / "ideas.json", {"generatedAt": generated_at, "items": ideas})
    print(f"書き出しました: {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
