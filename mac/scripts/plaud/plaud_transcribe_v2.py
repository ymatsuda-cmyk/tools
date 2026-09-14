#!/usr/bin/env python3
import argparse, json, os, subprocess, sys, tempfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
import transcribe_engines  # noqa: E402

ENV_FILE = Path.home() / ".plaud_notion_sync.env"
def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())
load_env()

DEFAULT_PLAUD_DOMAIN = "https://api-apne1.plaud.ai"
DEFAULT_PLAUD_WS_ID  = "ws_clQPe6Vll0"
PLAUD_MAX_ACCOUNTS   = 5

def load_plaud_accounts():
    """PLAUD_TOKEN, PLAUD_TOKEN_2 … を順に読み、複数アカウント分の設定を返す。
    2つ目以降で PLAUD_WS_ID_n が未設定なら workspaceId を付けずに既定ワークスペースを見る。"""
    accounts = []
    for idx in range(1, PLAUD_MAX_ACCOUNTS + 1):
        sfx = "" if idx == 1 else f"_{idx}"
        token = os.environ.get(f"PLAUD_TOKEN{sfx}", "").strip()
        if not token:
            continue
        accounts.append({
            "name": f"account{idx}",
            "token": token,
            "domain": (os.environ.get(f"PLAUD_DOMAIN{sfx}", "").strip() or DEFAULT_PLAUD_DOMAIN).rstrip("/"),
            "workspace": os.environ.get(
                f"PLAUD_WS_ID{sfx}", DEFAULT_PLAUD_WS_ID if idx == 1 else "").strip(),
        })
    return accounts

PLAUD_ACCOUNTS = load_plaud_accounts()
# retranscribe.py など外部から参照されるため1つ目のトークンは残す
PLAUD_TOKEN  = PLAUD_ACCOUNTS[0]["token"] if PLAUD_ACCOUNTS else ""
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
NOTION_DB_ID = os.environ.get("NOTION_DS_ID", "28b0e7a535dc805697c6d4b9f8032d18")
JST = timezone(timedelta(hours=9))

# ★ ローカルの tools リポジトリのクローン先に合わせて変更してください
MINUTES_INDEX_PATH = Path(os.environ.get(
    "MINUTES_INDEX_PATH",
    str(Path.home() / "tools" / "data" / "minutes" / "index.json")
))
GIT_AUTO_PUSH = os.environ.get("MINUTES_GIT_PUSH", "0") == "1"

def plaud_headers(account):
    return {
        "Authorization": account["token"],
        "Origin": "https://web.plaud.ai",
        "Referer": "https://web.plaud.ai/",
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/148.0.0.0 Safari/537.36"
    }

def find_account(name):
    for acc in PLAUD_ACCOUNTS:
        if acc["name"] == name:
            return acc
    return None

def ms_to_hms(ms):
    s = int(ms / 1000)
    h, r = divmod(s, 3600)
    m, sec = divmod(r, 60)
    if h > 0: return f"{h}時間 {m}分 {sec}秒"
    elif m > 0: return f"{m}分 {sec}秒"
    else: return f"{sec}秒"

def get_plaud_files_for_account(account):
    all_files = []
    page = 1
    while True:
        url = f"{account['domain']}/file/simple/web?pageSize=50&pageNum={page}"
        if account["workspace"]:
            url += f"&workspaceId={account['workspace']}"
        resp = requests.get(url, headers=plaud_headers(account), timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ [{account['name']}] 一覧取得失敗: {resp.status_code} {resp.text[:120]}")
            break
        data = resp.json()
        files = data.get("data_file_list", [])
        total = data.get("data_file_total", 0)
        if not files: break
        for f in files:
            f["_account"] = account["name"]
        all_files.extend(files)
        if len(all_files) >= total or len(files) < 50: break
        page += 1
    return all_files

def get_plaud_files(accounts=None):
    """全アカウントのファイルを結合する。同一IDは先に見つかったアカウントを採用。"""
    accounts = PLAUD_ACCOUNTS if accounts is None else accounts
    merged, seen = [], set()
    for acc in accounts:
        files = get_plaud_files_for_account(acc)
        print(f"    [{acc['name']}] {len(files)}件")
        for f in files:
            fid = f.get("id")
            if not fid or fid in seen:
                continue
            seen.add(fid)
            merged.append(f)
    return merged

def get_download_url(file_id, account=None):
    """account 未指定時は全アカウントを順に試す（再文字起こし用）。"""
    for acc in ([account] if account else PLAUD_ACCOUNTS):
        try:
            resp = requests.get(f"{acc['domain']}/file/temp-url/{file_id}",
                                headers=plaud_headers(acc), timeout=60)
        except requests.RequestException:
            continue
        if resp.status_code == 200:
            data = resp.json()
            url = data.get("temp_url") or data.get("temp_url_opus")
            if url:
                return url
    return None

def download_audio(temp_url, dest_path):
    resp = requests.get(temp_url, timeout=300, stream=True)
    if resp.status_code == 200:
        with open(dest_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    return False

def transcribe(audio_path, label=None, settings=None):
    """settings.json の label に従って whisper / ローカルLLM を切り替える"""
    return transcribe_engines.transcribe(audio_path, label=label, settings=settings)

def notion_headers():
    return {"Authorization": f"Bearer {NOTION_TOKEN}", "Notion-Version": "2022-06-28", "Content-Type": "application/json"}

def fetch_notion_pages():
    pages, has_more, cursor = [], True, None
    while has_more:
        payload = {"page_size": 100}
        if cursor: payload["start_cursor"] = cursor
        resp = requests.post(f"https://api.notion.com/v1/databases/{NOTION_DB_ID}/query",
                             headers=notion_headers(), json=payload, timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ Notion取得失敗: {resp.status_code} {resp.text[:200]}")
            break
        data = resp.json()
        pages.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return pages

def plaud_id_from_url(url_val):
    if url_val and "/file/" in url_val:
        return url_val.split("/file/")[-1].strip()
    return None

def get_registered_ids(pages):
    ids = set()
    for p in pages:
        rt = p.get("properties", {}).get("URL", {}).get("rich_text", [])
        fid = plaud_id_from_url(rt[0].get("plain_text", "")) if rt else None
        if fid: ids.add(fid)
    return ids

def text_to_blocks(text):
    blocks = []
    for para in [p.strip() for p in text.split('\n') if p.strip()]:
        while para:
            chunk, para = para[:2000], para[2000:]
            blocks.append({"object":"block","type":"paragraph","paragraph":{"rich_text":[{"type":"text","text":{"content":chunk}}]}})
    return blocks or [{"object":"block","type":"paragraph","paragraph":{"rich_text":[{"type":"text","text":{"content":"（文字起こし結果なし）"}}]}}]

def create_notion_page(f, transcript_text):
    dt_jst = datetime.fromtimestamp(f.get("start_time", 0) / 1000, tz=JST)
    utc_iso = dt_jst.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    name = f.get("filename") or dt_jst.strftime("%Y-%m-%d %H:%M:%S")
    children = [{"object":"block","type":"heading_2","heading_2":{"rich_text":[{"type":"text","text":{"content":f"🎙️ {f.get('fullname', f['id'] + '.ogg')}"}}]}}]
    children.extend(text_to_blocks(transcript_text))

    payload = {
        "parent": {"database_id": NOTION_DB_ID},
        "properties": {
            "ミーティング名": {"title": [{"text": {"content": name[:100]}}]},
            "日時": {"date": {"start": utc_iso}},
            "会議時間": {"rich_text": [{"text": {"content": ms_to_hms(f.get("duration", 0))}}]},
            "状態": {"select": {"name": "文字起こし"}},
            "URL": {"rich_text": [{"text": {"content": f"https://web.plaud.ai/file/{f['id']}"}}]}
        },
        "children": children[:100]
    }
    resp = requests.post("https://api.notion.com/v1/pages", headers=notion_headers(), json=payload, timeout=60)
    if resp.status_code != 200:
        print(f"  ❌ Notion登録失敗: {resp.status_code} {resp.text[:200]}")
        return None
    page_id = resp.json().get("id")
    remaining = children[100:]
    while remaining:
        batch, remaining = remaining[:100], remaining[100:]
        requests.patch(f"https://api.notion.com/v1/blocks/{page_id}/children", headers=notion_headers(), json={"children": batch}, timeout=60)
    return page_id

def _rich_text(props, name):
    items = props.get(name, {}).get("rich_text", [])
    return items[0].get("plain_text", "") if items else ""

def _multi_select_tags(props, name):
    items = props.get(name, {}).get("multi_select", [])
    return [it.get("name", "") for it in items if it.get("name")]

def to_iso_z(s):
    if not s: return ""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return s
    if dt.tzinfo is None: dt = dt.replace(tzinfo=JST)
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")

def page_to_entry(page):
    props = page.get("properties", {})
    url_val = _rich_text(props, "URL")
    key = plaud_id_from_url(url_val)
    if not key: return None            # PLAUD由来でないページは除外
    if page.get("archived") or page.get("in_trash"): return None
    title_items = props.get("ミーティング名", {}).get("title", [])
    return {
        "key": key,
        "title": title_items[0].get("plain_text", "") if title_items else "",
        "date": to_iso_z((props.get("日時", {}).get("date") or {}).get("start") or ""),
        "duration": _rich_text(props, "会議時間"),
        "status": (props.get("状態", {}).get("select") or {}).get("name", ""),
        "tags": _multi_select_tags(props, "カテゴリー"),
        "notionPageId": page.get("id", ""),
        "updatedAt": None,
    }

def write_minutes_index(pages):
    entries = [e for e in (page_to_entry(p) for p in pages) if e]

    previous = {}
    if MINUTES_INDEX_PATH.exists():
        try:
            for old in json.loads(MINUTES_INDEX_PATH.read_text(encoding="utf-8")):
                previous[old.get("key")] = old
        except (json.JSONDecodeError, OSError):
            previous = {}

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
    for e in entries:
        old = previous.get(e["key"])
        unchanged = old and all(old.get(k) == e[k] for k in ("title", "date", "duration", "status", "tags", "notionPageId"))
        e["updatedAt"] = (old.get("updatedAt") or now_iso) if unchanged else now_iso

    entries.sort(key=lambda e: e["date"], reverse=True)

    MINUTES_INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = MINUTES_INDEX_PATH.with_name(MINUTES_INDEX_PATH.name + ".tmp")
    tmp.write_text(json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(MINUTES_INDEX_PATH)
    return entries

def push_minutes_index():
    repo = str(MINUTES_INDEX_PATH.parent)
    branch = "main"
    branch_result = subprocess.run(
        ["git", "-C", repo, "branch", "--show-current"],
        capture_output=True,
        text=True,
    )
    if branch_result.returncode == 0 and branch_result.stdout.strip():
        branch = branch_result.stdout.strip()

    try:
        subprocess.run(["git", "-C", repo, "add", str(MINUTES_INDEX_PATH)],
                       check=True, capture_output=True, text=True)
        if subprocess.run(["git", "-C", repo, "diff", "--cached", "--quiet"],
                          capture_output=True, text=True).returncode == 0:
            print("    差分なしのためcommitをスキップ"); return
        msg = f"chore: update minutes index ({datetime.now(JST).strftime('%Y-%m-%d %H:%M')})"
        subprocess.run(["git", "-C", repo, "commit", "-m", msg], check=True, capture_output=True, text=True)
        try:
            subprocess.run(["git", "-C", repo, "push", "origin", branch], check=True, capture_output=True, text=True, timeout=180)
            print("    ✅ git push 完了")
        except subprocess.CalledProcessError as e:
            err = (e.stderr or "")
            non_ff = ("fetch first" in err.lower() or "non-fast-forward" in err.lower())
            if not non_ff:
                raise
            print("    ℹ️ リモートが先行しているため rebase 後に再 push します")
            subprocess.run(
                ["git", "-C", repo, "pull", "--rebase", "--autostash", "origin", branch],
                check=True,
                capture_output=True,
                text=True,
                timeout=180,
            )
            subprocess.run(["git", "-C", repo, "push", "origin", branch], check=True, capture_output=True, text=True, timeout=180)
            print("    ✅ git push 完了 (rebase後)")
    except subprocess.CalledProcessError as e:
        print(f"    ⚠️ git操作に失敗: {(e.stderr or '')[:200]}")
    except subprocess.TimeoutExpired:
        print("    ⚠️ git push がタイムアウトしました")

def build_index(pages=None):
    if pages is None:
        pages = fetch_notion_pages()
    entries = write_minutes_index(pages)
    print(f"    ✅ {MINUTES_INDEX_PATH} ({len(entries)}件)")
    if GIT_AUTO_PUSH:
        push_minutes_index()

def main():
    ap = argparse.ArgumentParser(description="PLAUD→文字起こし→Notion")
    ap.add_argument("--label", help="settings.json のエンジンlabel（例: whisper / qwen3-asr）。"
                                    "省略時は対話端末なら選択メニュー、cron等では既定値")
    ap.add_argument("--settings", help="settings.json のパス")
    args = ap.parse_args()

    settings = transcribe_engines.load_settings(args.settings)
    engine = transcribe_engines.choose_engine(settings, args.label)
    label = engine.get("label")

    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    print(f"\n{'='*60}\nPLAUD→文字起こし→Notion開始: {now}\nエンジン: {label} (type={engine.get('type')})\n{'='*60}")
    if not PLAUD_ACCOUNTS or not NOTION_TOKEN:
        print("❌ トークンが設定されていません"); return
    print(f"PLAUDアカウント: {len(PLAUD_ACCOUNTS)}件 ({', '.join(a['name'] for a in PLAUD_ACCOUNTS)})")

    print("\n[1] Notionの既存ページを取得中...")
    notion_pages = fetch_notion_pages()
    registered_ids = get_registered_ids(notion_pages)
    print(f"    ページ: {len(notion_pages)}件 / 登録済みID: {len(registered_ids)}件")

    print("\n[2] PLAUDからファイル一覧を取得中...")
    plaud_files = get_plaud_files()
    print(f"    取得: {len(plaud_files)}件")
    if not plaud_files:
        print("❌ PLAUDからファイルを取得できませんでした"); return

    print("\n[3] 未登録ファイルを抽出中...")
    new_files = [f for f in plaud_files if f["id"] not in registered_ids]
    new_files.sort(key=lambda x: x.get("start_time", 0))
    print(f"    未登録: {len(new_files)}件")
    if not new_files:
        print("\n✅ 新規ファイルなし。index.json を更新して終了します。")
        print("\n[4] index.json を更新中...")
        build_index(notion_pages)
        return

    print(f"\n[4] {len(new_files)}件を処理中...")
    with tempfile.TemporaryDirectory() as tmpdir:
        for i, f in enumerate(new_files, 1):
            file_id = f["id"]
            filename = f.get("fullname", f"{file_id}.ogg")
            name = f.get("filename", file_id)
            account = find_account(f.get("_account"))
            origin = f" <{account['name']}>" if account else ""
            print(f"\n  [{i}/{len(new_files)}] {name}{origin}")

            temp_url = get_download_url(file_id, account)
            if not temp_url:
                print(f"  ❌ URL取得失敗。スキップ"); continue

            audio_path = Path(tmpdir) / filename
            print(f"  → ダウンロード中... ({filename})")
            if not download_audio(temp_url, str(audio_path)):
                print(f"  ❌ ダウンロード失敗。スキップ"); continue
            print(f"  ✅ {audio_path.stat().st_size/1024/1024:.1f} MB")

            transcript = transcribe(str(audio_path), label=label, settings=settings)
            if not transcript:
                print(f"  ❌ 文字起こし失敗。スキップ"); continue
            print(f"  ✅ 文字起こし完了 ({len(transcript)}文字)")

            page_id = create_notion_page(f, transcript)
            if page_id:
                print(f"  ✅ Notion登録完了")

    print("\n[5] 最新のNotionページから index.json を更新中...")
    latest_notion_pages = fetch_notion_pages()
    build_index(latest_notion_pages)

    print(f"\n{'='*60}\n✅ 全処理完了: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S JST')}\n{'='*60}\n")

if __name__ == "__main__":
    main()