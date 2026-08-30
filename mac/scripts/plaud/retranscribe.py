#!/usr/bin/env python3
"""Notion上の既存ページを対象に、PLAUD音声を再ダウンロードして再文字起こしする

対象の指定方法（どちらか一方）:
  --page-id  <Notion page ID>   特定の1ページだけを対象にする
  --status   <状態の値>          その状態のページをまとめて対象にする（複数指定可、既定: 再取得）

既存ページの本文（transcript部分）はアーカイブ（archived）してから、
新しい文字起こし結果を追記する。ミーティング名・日時などのプロパティは変更しない。
処理完了後、状態プロパティを --set-status（既定: 文字起こし）に更新する。

使い方:
    python3 retranscribe.py --status 再取得 --label qwen3-asr
    python3 retranscribe.py --page-id 1a2b3c4d... --label qwen3-asr
    python3 retranscribe.py --status 再取得 --dry-run    # 対象一覧を見るだけ
"""
import argparse
import importlib.util
import os
import sys
from datetime import datetime
from pathlib import Path

import requests

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

# メインスクリプトのファイル名が環境によって異なる場合に対応する
# （例: plaud_transcribe_notion.py / plaud_transcribe_v2.py）
_MAIN_SCRIPT_CANDIDATES = [
    os.environ.get("PLAUD_MAIN_SCRIPT", ""),
    "plaud_transcribe_notion.py",
    "plaud_transcribe_v2.py",
]


def _load_main_module():
    for name in _MAIN_SCRIPT_CANDIDATES:
        if not name:
            continue
        path = SCRIPT_DIR / name
        if path.exists():
            spec = importlib.util.spec_from_file_location("plaud_transcribe_notion", path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module
    tried = ", ".join(n for n in _MAIN_SCRIPT_CANDIDATES if n)
    raise FileNotFoundError(
        f"メインスクリプトが見つかりません（{SCRIPT_DIR} 内に {tried} のいずれも無い）。\n"
        f"  PLAUD_MAIN_SCRIPT=実際のファイル名.py を指定するか、"
        f"ファイル名を plaud_transcribe_notion.py に変更してください。"
    )


PN = _load_main_module()  # noqa: E402  既存関数・定数を再利用
import transcribe_engines  # noqa: E402


def query_pages_by_status(statuses):
    """状態プロパティが statuses のいずれかに一致するページを全件取得する"""
    if not statuses:
        return []
    filter_ = {"or": [{"property": "状態", "select": {"equals": s}} for s in statuses]}
    pages, has_more, cursor = [], True, None
    while has_more:
        payload = {"page_size": 100, "filter": filter_}
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(
            f"https://api.notion.com/v1/databases/{PN.NOTION_DB_ID}/query",
            headers=PN.notion_headers(), json=payload, timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ Notion検索失敗: {resp.status_code} {resp.text[:200]}")
            break
        data = resp.json()
        pages.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return pages


def fetch_page(page_id):
    resp = requests.get(f"https://api.notion.com/v1/pages/{page_id}",
                        headers=PN.notion_headers(), timeout=60)
    if resp.status_code != 200:
        print(f"  ❌ ページ取得失敗: {resp.status_code} {resp.text[:200]}")
        return None
    return resp.json()


def page_title(page):
    items = page.get("properties", {}).get("ミーティング名", {}).get("title", [])
    return items[0].get("plain_text", "") if items else "(無題)"


def get_all_children(block_id):
    """ページ直下の子ブロックを全件取得（ページネーション対応）"""
    children, has_more, cursor = [], True, None
    while has_more:
        url = f"https://api.notion.com/v1/blocks/{block_id}/children?page_size=100"
        if cursor:
            url += f"&start_cursor={cursor}"
        resp = requests.get(url, headers=PN.notion_headers(), timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ ブロック取得失敗: {resp.status_code} {resp.text[:200]}")
            break
        data = resp.json()
        children.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return children


def archive_all_children(page_id):
    """既存の本文ブロックをすべてアーカイブ（旧内容を除去）する"""
    children = get_all_children(page_id)
    ok = 0
    for block in children:
        block_id = block.get("id")
        resp = requests.patch(f"https://api.notion.com/v1/blocks/{block_id}",
                              headers=PN.notion_headers(), json={"archived": True}, timeout=60)
        if resp.status_code == 200:
            ok += 1
        else:
            print(f"    ⚠️ ブロック削除失敗: {resp.status_code} {resp.text[:200]}")
    return ok, len(children)


def append_transcript(page_id, filename, transcript_text, engine_label):
    """再文字起こし結果を本文として追記する"""
    ts = datetime.now(PN.JST).strftime("%Y-%m-%d %H:%M:%S JST")
    children = [
        {"object": "block", "type": "heading_2", "heading_2": {"rich_text": [
            {"type": "text", "text": {"content": f"🎙️ {filename}"}}]}},
        {"object": "block", "type": "callout", "callout": {
            "icon": {"type": "emoji", "emoji": "🔁"},
            "rich_text": [{"type": "text", "text": {
                "content": f"再文字起こし: {ts}（エンジン: {engine_label}）"}}]}},
    ]
    children.extend(PN.text_to_blocks(transcript_text))

    resp = requests.patch(f"https://api.notion.com/v1/blocks/{page_id}/children",
                          headers=PN.notion_headers(), json={"children": children[:100]}, timeout=60)
    if resp.status_code != 200:
        print(f"    ❌ 本文追記失敗: {resp.status_code} {resp.text[:200]}")
        return False
    remaining = children[100:]
    while remaining:
        batch, remaining = remaining[:100], remaining[100:]
        requests.patch(f"https://api.notion.com/v1/blocks/{page_id}/children",
                       headers=PN.notion_headers(), json={"children": batch}, timeout=60)
    return True


def set_status(page_id, status_name):
    resp = requests.patch(f"https://api.notion.com/v1/pages/{page_id}",
                          headers=PN.notion_headers(),
                          json={"properties": {"状態": {"select": {"name": status_name}}}},
                          timeout=60)
    if resp.status_code != 200:
        print(f"    ⚠️ 状態更新失敗: {resp.status_code} {resp.text[:200]}")
        return False
    return True


def resolve_targets(page_id=None, statuses=None):
    if page_id:
        page = fetch_page(page_id)
        return [page] if page else []
    return query_pages_by_status(statuses or ["再取得"])


def process_page(page, label, settings, set_status_name):
    title = page_title(page)
    page_id = page.get("id")
    url_val = PN._rich_text(page.get("properties", {}), "URL")
    file_id = PN.plaud_id_from_url(url_val)
    if not file_id:
        print(f"  ❌ PLAUDのURLが見つからないためスキップ: {title}")
        return False

    print(f"  対象: {title}  (page_id={page_id}, plaud_id={file_id})")

    temp_url = PN.get_download_url(file_id)
    if not temp_url:
        print("  ❌ ダウンロードURL取得失敗。スキップ")
        return False

    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        filename = f"{file_id}.ogg"
        audio_path = Path(tmpdir) / filename
        print("  → ダウンロード中...")
        if not PN.download_audio(temp_url, str(audio_path)):
            print("  ❌ ダウンロード失敗。スキップ")
            return False
        print(f"  ✅ {audio_path.stat().st_size/1024/1024:.1f} MB")

        transcript = transcribe_engines.transcribe(str(audio_path), label=label, settings=settings)
        if not transcript:
            print("  ❌ 文字起こし失敗。スキップ")
            return False
        print(f"  ✅ 文字起こし完了 ({len(transcript)}文字)")

    print("  → 既存の本文を削除中...")
    ok, total = archive_all_children(page_id)
    print(f"  ✅ {ok}/{total} ブロックを削除")

    engine = transcribe_engines.get_engine(settings, label)
    if not append_transcript(page_id, filename, transcript, engine.get("label")):
        return False
    print("  ✅ 新しい文字起こしを追記")

    if set_status_name:
        if set_status(page_id, set_status_name):
            print(f"  ✅ 状態を「{set_status_name}」に更新")

    return True


def main():
    ap = argparse.ArgumentParser(description="Notionページの再文字起こし")
    ap.add_argument("--page-id", help="対象のNotionページID（単体指定。--statusより優先）")
    ap.add_argument("--status", nargs="+", default=["再取得"],
                    help="対象とする状態（複数可、既定: 再取得）")
    ap.add_argument("--label", help="settings.json のエンジンlabel（例: whisper / qwen3-asr）。"
                                    "省略時は対話端末なら選択メニュー、cron等では既定値")
    ap.add_argument("--settings", help="settings.json のパス")
    ap.add_argument("--set-status", default="文字起こし",
                    help="処理完了後に設定する状態（既定: 文字起こし。空文字で更新しない）")
    ap.add_argument("--dry-run", action="store_true", help="対象一覧を表示するだけで実行しない")
    args = ap.parse_args()

    if not PN.PLAUD_TOKEN or not PN.NOTION_TOKEN:
        print("❌ トークンが設定されていません")
        return

    settings = transcribe_engines.load_settings(args.settings)
    engine = transcribe_engines.choose_engine(settings, args.label)
    label = engine.get("label")

    targets = resolve_targets(page_id=args.page_id, statuses=args.status)
    if not targets:
        print("対象ページが見つかりませんでした。")
        return

    print(f"対象: {len(targets)}件  エンジン: {label} (type={engine.get('type')})")
    for p in targets:
        print(f"  - {page_title(p)}  ({p.get('id')})")

    if args.dry_run:
        print("\n--dry-run のため処理は行いません。")
        return

    set_status_name = args.set_status or None
    ok_count = 0
    for i, page in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}]")
        if process_page(page, label, settings, set_status_name):
            ok_count += 1

    print(f"\n完了: {ok_count}/{len(targets)}件")

    print("\nindex.json を更新中...")
    PN.build_index()


if __name__ == "__main__":
    main()