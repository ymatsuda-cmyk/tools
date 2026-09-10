#!/usr/bin/env python3
"""Notion上の既存ページを対象に、PLAUD音声を再ダウンロードして再文字起こしする

対象の指定方法（どちらか一方）:
  --page-id  <Notion page ID>   特定の1ページだけを対象にする
  --status   <状態の値>          その状態のページをまとめて対象にする（複数指定可、既定: 再取得）

処理順序（重要）:
  1. 既存ブロックのIDとテキストを退避（ローカルにバックアップも保存）
  2. 新しい文字起こしを追記
  3. 追記が成功した場合のみ、退避しておいた旧ブロックを削除
  この順序により、追記失敗時に本文が消える事故を防ぐ。

処理完了後、状態プロパティを --set-status（既定: 文字起こし）に更新する。

使い方:
    python3 retranscribe.py --status 再取得 --label qwen3-asr
    python3 retranscribe.py --page-id 1a2b3c4d... --label qwen3-asr
    python3 retranscribe.py --status 再取得 --dry-run    # 対象一覧を見るだけ
    python3 retranscribe.py --status 再取得 --keep-old   # 旧本文を残したまま追記
"""
import argparse
import importlib.util
import json
import os
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

import requests

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

BACKUP_DIR = SCRIPT_DIR / "logs" / "retranscribe_backup"

# Notion API は平均 3リクエスト/秒 が上限。連投すると 429 が返る。
NOTION_MIN_INTERVAL = 0.34
NOTION_MAX_RETRY = 5
_last_notion_call = 0.0

# メインスクリプトのファイル名が環境によって異なる場合に対応する
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


# ── Notion 呼び出し共通（レート制限・リトライ付き） ─────────────
def notion_request(method, url, *, json_body=None, timeout=60):
    """全Notion呼び出しをここに集約し、429/5xx を指数バックオフで再試行する。"""
    global _last_notion_call
    for attempt in range(NOTION_MAX_RETRY):
        wait = NOTION_MIN_INTERVAL - (time.monotonic() - _last_notion_call)
        if wait > 0:
            time.sleep(wait)
        try:
            resp = requests.request(method, url, headers=PN.notion_headers(),
                                    json=json_body, timeout=timeout)
        except requests.RequestException as e:
            if attempt == NOTION_MAX_RETRY - 1:
                print(f"    ⚠️ Notion通信エラー: {type(e).__name__}")
                return None
            time.sleep(2 ** attempt)
            continue
        finally:
            _last_notion_call = time.monotonic()

        if resp.status_code == 429:
            retry_after = float(resp.headers.get("Retry-After", 2 ** attempt))
            print(f"    ⏳ レート制限。{retry_after:.0f}秒待機して再試行")
            time.sleep(retry_after)
            continue
        if 500 <= resp.status_code < 600 and attempt < NOTION_MAX_RETRY - 1:
            time.sleep(2 ** attempt)
            continue
        return resp
    return None


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
        resp = notion_request(
            "POST", f"https://api.notion.com/v1/databases/{PN.NOTION_DB_ID}/query",
            json_body=payload)
        if resp is None or resp.status_code != 200:
            code = resp.status_code if resp is not None else "-"
            body = resp.text[:200] if resp is not None else ""
            print(f"    ⚠️ Notion検索失敗: {code} {body}")
            break
        data = resp.json()
        pages.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return pages


def fetch_page(page_id):
    resp = notion_request("GET", f"https://api.notion.com/v1/pages/{page_id}")
    if resp is None or resp.status_code != 200:
        code = resp.status_code if resp is not None else "-"
        body = resp.text[:200] if resp is not None else ""
        print(f"  ❌ ページ取得失敗: {code} {body}")
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
        resp = notion_request("GET", url)
        if resp is None or resp.status_code != 200:
            code = resp.status_code if resp is not None else "-"
            body = resp.text[:200] if resp is not None else ""
            print(f"    ⚠️ ブロック取得失敗: {code} {body}")
            break
        data = resp.json()
        children.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return children


def block_to_text(block):
    """ブロックからプレーンテキストを抽出（バックアップ用）"""
    btype = block.get("type", "")
    payload = block.get(btype)
    if not isinstance(payload, dict):
        return ""
    rich = payload.get("rich_text") or payload.get("text") or []
    return "".join(r.get("plain_text", "") for r in rich if isinstance(r, dict))


def save_backup(page_id, title, blocks):
    """削除前の本文をローカルに保存する。Notion側の操作に失敗しても復元できる。"""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(PN.JST).strftime("%Y%m%d_%H%M%S")
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in title)[:48]
    path = BACKUP_DIR / f"{ts}_{safe}_{page_id[:8]}.json"
    data = {
        "page_id": page_id,
        "title": title,
        "saved_at": datetime.now(PN.JST).isoformat(),
        "text": "\n".join(t for t in (block_to_text(b) for b in blocks) if t),
        "blocks": blocks,
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def delete_blocks(block_ids):
    """指定したブロックだけを削除する。IDを事前に確定させてから呼ぶこと。"""
    ok = 0
    for block_id in block_ids:
        resp = notion_request("DELETE", f"https://api.notion.com/v1/blocks/{block_id}")
        if resp is not None and resp.status_code == 200:
            ok += 1
        else:
            code = resp.status_code if resp is not None else "-"
            body = resp.text[:200] if resp is not None else ""
            print(f"    ⚠️ ブロック削除失敗: {code} {body}")
    return ok


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

    remaining = children
    appended = 0
    while remaining:
        batch, remaining = remaining[:100], remaining[100:]
        resp = notion_request("PATCH",
                              f"https://api.notion.com/v1/blocks/{page_id}/children",
                              json_body={"children": batch})
        if resp is None or resp.status_code != 200:
            code = resp.status_code if resp is not None else "-"
            body = resp.text[:200] if resp is not None else ""
            print(f"    ❌ 本文追記失敗: {code} {body}")
            # 部分追記で終わった場合も旧本文は消さない（呼び出し元でFalse判定）
            return False
        appended += len(batch)
    return appended > 0


def set_status(page_id, status_name):
    resp = notion_request("PATCH", f"https://api.notion.com/v1/pages/{page_id}",
                          json_body={"properties": {"状態": {"select": {"name": status_name}}}})
    if resp is None or resp.status_code != 200:
        code = resp.status_code if resp is not None else "-"
        body = resp.text[:200] if resp is not None else ""
        print(f"    ⚠️ 状態更新失敗: {code} {body}")
        return False
    return True


def resolve_targets(page_id=None, statuses=None):
    if page_id:
        page = fetch_page(page_id)
        return [page] if page else []
    return query_pages_by_status(statuses or ["再取得"])


def apply_glossary(transcript):
    """メインスクリプト側の用語辞書があれば適用する（無ければ素通し）"""
    fn = getattr(PN, "apply_replacements", None)
    if not callable(fn):
        return transcript
    try:
        result = fn(transcript)
    except Exception:
        return transcript
    if isinstance(result, tuple):
        text, n = result
        if n:
            print(f"  ✏️ 用語辞書で {n}箇所を補正")
        return text
    return result


def process_page(page, engine, settings, set_status_name, keep_old=False):
    title = page_title(page)
    page_id = page.get("id")
    label = engine.get("label")
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

    transcript = apply_glossary(transcript)

    # ── 破壊的操作の順序 ──────────────────────────────
    # 旧: 削除 → 追記（追記に失敗すると本文が失われる）
    # 新: ID退避＋バックアップ → 追記 → 成功時のみ削除
    old_blocks = get_all_children(page_id)
    old_ids = [b.get("id") for b in old_blocks if b.get("id")]

    if old_ids and not keep_old:
        backup_path = save_backup(page_id, title, old_blocks)
        print(f"  💾 旧本文をバックアップ: {backup_path.name}")

    if not append_transcript(page_id, filename, transcript, label):
        print("  ⚠️ 追記に失敗したため、旧本文は削除せず残します")
        return False
    print("  ✅ 新しい文字起こしを追記")

    if old_ids and not keep_old:
        print("  → 既存の本文を削除中...")
        ok = delete_blocks(old_ids)
        print(f"  ✅ {ok}/{len(old_ids)} ブロックを削除")
        if ok < len(old_ids):
            print("     一部残っています。ページを確認してください")

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
    ap.add_argument("--keep-old", action="store_true",
                    help="旧本文を削除せず、新しい文字起こしを下に追記する")
    ap.add_argument("--dry-run", action="store_true", help="対象一覧を表示するだけで実行しない")
    args = ap.parse_args()

    if not PN.PLAUD_TOKEN or not PN.NOTION_TOKEN:
        print("❌ トークンが設定されていません")
        return 1

    settings = transcribe_engines.load_settings(args.settings)
    engine = transcribe_engines.choose_engine(settings, args.label)
    label = engine.get("label")

    targets = resolve_targets(page_id=args.page_id, statuses=args.status)
    if not targets:
        print("対象ページが見つかりませんでした。")
        return 1

    print(f"対象: {len(targets)}件  エンジン: {label} (type={engine.get('type')})")
    for p in targets:
        print(f"  - {page_title(p)}  ({p.get('id')})")

    if args.dry_run:
        print("\n--dry-run のため処理は行いません。")
        return 0

    set_status_name = args.set_status or None
    ok_count = 0
    failed = []
    for i, page in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}]")
        try:
            if process_page(page, engine, settings, set_status_name, keep_old=args.keep_old):
                ok_count += 1
            else:
                failed.append(page_title(page))
        except KeyboardInterrupt:
            print("\n中断されました。")
            break
        except Exception as e:
            # 1件の失敗で全体を止めない
            print(f"  ❌ 想定外のエラー: {type(e).__name__}: {e}")
            failed.append(page_title(page))

    print(f"\n完了: {ok_count}/{len(targets)}件")
    if failed:
        print("失敗したページ:")
        for t in failed:
            print(f"  - {t}")

    print("\nindex.json を更新中...")
    PN.build_index()
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
