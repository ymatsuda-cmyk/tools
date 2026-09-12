#!/usr/bin/env python3
"""NotionのコンテンツDBを読み、contentsstockの一覧用JSONを書き出す。

アプリを開くたびにGAS経由でNotionを全件クエリすると、件数が増えるほど
最初の描画が遅くなる。cronでこれを回してJSONを先に用意しておき、
画面はそれを読むだけにする。詳細(タブごとの本文)は従来どおりNotionから取る。

出力:
  index-doc.json  カード表示・検索・絞り込みに要る項目(長文は有無のフラグだけ)
  idea-doc.json   応用と活用アイデアの本文(アイデア一覧画面が使う)

環境変数 (既定で ~/.contentsstock.env からも読む):
  NOTION_TOKEN          Notion Integration Token(必須)
  CONTENTS_DB_ID        コンテンツDBのID
  CONTENTS_OUT_DIR      出力先ディレクトリ(--out より弱い)
  CONTENTS_ENV_FILE     環境変数ファイルのパス

使い方:
    python3 build_contentsstock_json.py
    python3 build_contentsstock_json.py --out ~/scripts/data/contentsstock
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
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
CONTENTS_DB_ID = os.environ.get("CONTENTS_DB_ID", "d600e7a535dc83caadf381afe7abea03")

# Notion側のカラム名。gas/Code.gs の PROP_* と一致させること
PROP_TITLE = "タイトル"
PROP_FILE = "ファイル名"
PROP_KIND = "種別"
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
PROP_PUBLIC = "公開"
PROP_CREATED = "作成日時"

STATUS_DONE = "完了"

# ---------------------------------------------------------------- Notion


def notion_headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }


def query_all_pages(db_id):
    """DBの全ページを作成時刻の新しい順で取得する。"""
    pages = []
    cursor = None
    while True:
        payload = {"page_size": 100, "sorts": [{"timestamp": "created_time", "direction": "descending"}]}
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(f"{NOTION_API}/databases/{db_id}/query",
                             headers=notion_headers(), json=payload, timeout=60)
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


def rich_of(props, name):
    return plain_text((props.get(name) or {}).get("rich_text"))


def select_of(props, name):
    sel = (props.get(name) or {}).get("select")
    return (sel or {}).get("name") or ""


def multi_select_of(props, name):
    return [o.get("name", "") for o in (props.get(name) or {}).get("multi_select") or []]


def number_of(props, name):
    value = (props.get(name) or {}).get("number")
    return value if isinstance(value, (int, float)) else 0


def checkbox_of(props, name):
    return bool((props.get(name) or {}).get("checkbox"))


def date_of(props, name):
    date = (props.get(name) or {}).get("date")
    return (date or {}).get("start")


def created_of(page, props):
    prop = props.get(PROP_CREATED) or {}
    return prop.get("created_time") or page.get("created_time")


def title_any_of(props):
    """title型のプロパティを探す。DBによって名前が違うことがあるため。"""
    for value in props.values():
        if isinstance(value, dict) and value.get("type") == "title":
            text = plain_text(value.get("title"))
            if text:
                return text
    return ""


# ---------------------------------------------------------------- 変換


def to_item(page):
    """gas/Code.gs の listContents_ と同じ形にする。"""
    p = page.get("properties", {})
    return {
        "key": page["id"],
        "source": "doc",
        "title": title_any_of(p) or rich_of(p, PROP_FILE) or "(タイトル未設定)",
        "file": rich_of(p, PROP_FILE),
        "kind": select_of(p, PROP_KIND),
        "status": select_of(p, PROP_STATUS) or STATUS_DONE,
        "tags": multi_select_of(p, PROP_TAGS),
        "createdAt": created_of(page, p),
        "editedAt": page.get("last_edited_time"),
        "summary": rich_of(p, PROP_SUMMARY),
        "model": rich_of(p, PROP_MODEL) or None,
        "generatedAt": date_of(p, PROP_GENERATED),
        "rawCount": number_of(p, PROP_RAW_COUNT),
        "isPublic": checkbox_of(p, PROP_PUBLIC),
        "has": {
            "mindmap": bool(rich_of(p, PROP_MINDMAP)),
            "fields": bool(rich_of(p, PROP_FIELDS)),
            "apply": bool(rich_of(p, PROP_APPLY)),
            "ideas": bool(rich_of(p, PROP_IDEAS)),
            "memo": bool(rich_of(p, PROP_MEMO)),
        },
    }


def to_idea(page):
    """応用も活用も無いページは None を返す。"""
    p = page.get("properties", {})
    apply_text = rich_of(p, PROP_APPLY)
    ideas_text = rich_of(p, PROP_IDEAS)
    if not apply_text and not ideas_text:
        return None
    return {
        "key": page["id"],
        "source": "doc",
        "title": title_any_of(p),
        "kind": select_of(p, PROP_KIND),
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
    parser = argparse.ArgumentParser(description="コンテンツDBの一覧JSONを書き出す")
    parser.add_argument(
        "--out",
        default=os.environ.get("CONTENTS_OUT_DIR", str(REPO_ROOT / "data" / "contentsstock")),
        help="出力先ディレクトリ",
    )
    parser.add_argument("--db-id", default=CONTENTS_DB_ID, help="コンテンツDBのID")
    parser.add_argument("--dry-run", action="store_true", help="件数だけ表示して書き出さない")
    args = parser.parse_args()

    if not NOTION_TOKEN:
        print("NOTION_TOKEN が未設定です", file=sys.stderr)
        return 1

    pages = query_all_pages(args.db_id)
    items = [to_item(page) for page in pages]
    items.sort(key=lambda i: i.get("createdAt") or "", reverse=True)
    ideas = [idea for idea in (to_idea(page) for page in pages) if idea]
    generated_at = datetime.now(JST).isoformat()

    print(f"文書 {len(items)}件 / アイデアのあるもの {len(ideas)}件")
    if args.dry_run:
        return 0

    out_dir = Path(args.out).expanduser()
    write_json(out_dir / "index-doc.json", {"generatedAt": generated_at, "items": items})
    write_json(out_dir / "idea-doc.json", {"generatedAt": generated_at, "items": ideas})
    print(f"書き出しました: {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
