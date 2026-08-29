#!/usr/bin/env python3
import importlib.util, json, os, requests, shlex, shutil, subprocess, sys, tempfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

ENV_FILE = Path.home() / ".plaud_notion_sync.env"
def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())
load_env()

PLAUD_TOKEN  = os.environ.get("PLAUD_TOKEN", "")
PLAUD_DOMAIN = os.environ.get("PLAUD_DOMAIN", "https://api-apne1.plaud.ai")
PLAUD_WS_ID  = "ws_clQPe6Vll0"
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
NOTION_DB_ID = os.environ.get("NOTION_DS_ID", "28b0e7a535dc805697c6d4b9f8032d18")
WHISPER_MODEL = "mlx-community/whisper-medium-mlx"
JST = timezone(timedelta(hours=9))
_MLX_WHISPER_CMD = None
_RUNTIME_PATH_PREPARED = False
SCRIPT_DIR = Path(__file__).resolve().parent
DEBUG_LOG_DIR = SCRIPT_DIR / "logs" / "transcribe_debug"

# ★ ローカルの tools リポジトリのクローン先に合わせて変更してください
MINUTES_INDEX_PATH = Path(os.environ.get(
    "MINUTES_INDEX_PATH",
    str(Path.home() / "repos" / "tools" / "data" / "minutes" / "index.json")
))
GIT_AUTO_PUSH = os.environ.get("MINUTES_GIT_PUSH", "0") == "1"

PLAUD_HEADERS = {
    "Authorization": PLAUD_TOKEN,
    "Origin": "https://web.plaud.ai",
    "Referer": "https://web.plaud.ai/",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/148.0.0.0 Safari/537.36"
}

def ms_to_hms(ms):
    s = int(ms / 1000)
    h, r = divmod(s, 3600)
    m, sec = divmod(r, 60)
    if h > 0: return f"{h}時間 {m}分 {sec}秒"
    elif m > 0: return f"{m}分 {sec}秒"
    else: return f"{sec}秒"

def get_plaud_files():
    all_files = []
    page = 1
    while True:
        url = f"{PLAUD_DOMAIN}/file/simple/web?pageSize=50&pageNum={page}&workspaceId={PLAUD_WS_ID}"
        resp = requests.get(url, headers=PLAUD_HEADERS, timeout=60)
        if resp.status_code != 200: break
        data = resp.json()
        files = data.get("data_file_list", [])
        total = data.get("data_file_total", 0)
        if not files: break
        all_files.extend(files)
        if len(all_files) >= total or len(files) < 50: break
        page += 1
    return all_files

def get_download_url(file_id):
    resp = requests.get(f"{PLAUD_DOMAIN}/file/temp-url/{file_id}", headers=PLAUD_HEADERS, timeout=60)
    if resp.status_code == 200:
        data = resp.json()
        return data.get("temp_url") or data.get("temp_url_opus")
    return None

def download_audio(temp_url, dest_path):
    resp = requests.get(temp_url, timeout=300, stream=True)
    if resp.status_code == 200:
        with open(dest_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    return False

def resolve_mlx_whisper_cmd():
    global _MLX_WHISPER_CMD
    if _MLX_WHISPER_CMD is not None:
        return _MLX_WHISPER_CMD

    cli_path = shutil.which("mlx_whisper")
    if cli_path:
        _MLX_WHISPER_CMD = [cli_path]
        return _MLX_WHISPER_CMD

    # PATHにユーザーサイトのbinが入っていない環境向けに既知パスも確認する
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}"
    known_candidates = [
        Path.home() / "Library" / "Python" / py_ver / "bin" / "mlx_whisper",
        Path.home() / "Library" / "Python" / "3.9" / "bin" / "mlx_whisper",
        Path("/opt/homebrew/bin/mlx_whisper"),
        Path("/usr/local/bin/mlx_whisper"),
    ]
    for candidate in known_candidates:
        if candidate.exists() and os.access(candidate, os.X_OK):
            _MLX_WHISPER_CMD = [str(candidate)]
            return _MLX_WHISPER_CMD

    if importlib.util.find_spec("mlx_whisper") is not None:
        _MLX_WHISPER_CMD = [sys.executable, "-m", "mlx_whisper"]
        return _MLX_WHISPER_CMD

    _MLX_WHISPER_CMD = []
    return _MLX_WHISPER_CMD

def ensure_runtime_path():
    global _RUNTIME_PATH_PREPARED
    if _RUNTIME_PATH_PREPARED:
        return

    current = [p for p in os.environ.get("PATH", "").split(":") if p]
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}"
    candidates = [
        "/opt/homebrew/bin",
        "/usr/local/bin",
        str(Path.home() / "Library" / "Python" / py_ver / "bin"),
        str(Path.home() / "Library" / "Python" / "3.9" / "bin"),
    ]
    for c in candidates:
        if c not in current and Path(c).exists():
            current.append(c)
    os.environ["PATH"] = ":".join(current)
    _RUNTIME_PATH_PREPARED = True

def resolve_ffmpeg_cmd():
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg
    for candidate in [Path("/opt/homebrew/bin/ffmpeg"), Path("/usr/local/bin/ffmpeg")]:
        if candidate.exists() and os.access(candidate, os.X_OK):
            return str(candidate)
    return None

def write_transcribe_debug_log(audio_path, cmd, *, result=None, error=None, note=""):
    DEBUG_LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(JST).strftime("%Y%m%d_%H%M%S")
    audio_stem = Path(audio_path).stem
    safe_stem = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in audio_stem)[:64]
    log_path = DEBUG_LOG_DIR / f"{ts}_{safe_stem}.log"

    lines = [
        f"timestamp_jst: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S %z')}",
        f"audio_path: {audio_path}",
        f"cwd: {Path(audio_path).parent}",
        f"python: {sys.executable}",
        f"path: {os.environ.get('PATH', '')}",
        f"which_ffmpeg: {shutil.which('ffmpeg') or '(not found)'}",
        f"command: {shlex.join(cmd)}",
    ]
    if note:
        lines.append(f"note: {note}")
    if error is not None:
        lines.extend([
            "error_type:",
            type(error).__name__,
            "error_message:",
            str(error),
        ])
    if result is not None:
        lines.extend([
            f"returncode: {result.returncode}",
            "stdout:",
            result.stdout or "",
            "stderr:",
            result.stderr or "",
        ])

    txt_files = sorted([str(p.name) for p in Path(audio_path).parent.glob("*.txt")])
    lines.extend([
        "txt_files_in_output_dir:",
        "\n".join(txt_files) if txt_files else "(none)",
    ])

    log_path.write_text("\n".join(lines), encoding="utf-8")
    return log_path

def transcribe(audio_path):
    audio_path = Path(audio_path)
    out_dir = audio_path.parent
    ensure_runtime_path()

    ffmpeg_cmd = resolve_ffmpeg_cmd()
    if not ffmpeg_cmd:
        print("  ❌ ffmpeg が見つかりません。")
        print("     例: brew install ffmpeg")
        log_path = write_transcribe_debug_log(audio_path, ["ffmpeg"], note="ffmpeg_not_found")
        print(f"  📝 デバッグログ: {log_path}")
        return None

    cmd_prefix = resolve_mlx_whisper_cmd()
    if not cmd_prefix:
        print("  ❌ mlx_whisper が見つかりません。")
        print("     例: pip install mlx-whisper")
        log_path = write_transcribe_debug_log(audio_path, ["mlx_whisper"], note="mlx_whisper_not_found")
        print(f"  📝 デバッグログ: {log_path}")
        return None

    cmd = [
        *cmd_prefix, str(audio_path),
        "--model", WHISPER_MODEL,
        "--output-format", "txt",
        "--output-dir", str(out_dir),
        "--language", "ja",
        "--condition-on-previous-text", "False"
    ]
    print(f"  文字起こし実行中...")
    try:
        env = os.environ.copy()
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800, cwd=str(out_dir), env=env)
    except FileNotFoundError:
        print("  ❌ mlx_whisper コマンドを起動できませんでした。")
        print("     例: pip install mlx-whisper")
        log_path = write_transcribe_debug_log(audio_path, cmd, note="mlx_whisper_exec_not_found")
        print(f"  📝 デバッグログ: {log_path}")
        return None
    except subprocess.TimeoutExpired as e:
        print("  ❌ 文字起こしがタイムアウトしました。")
        log_path = write_transcribe_debug_log(audio_path, cmd, error=e, note="timeout")
        print(f"  📝 デバッグログ: {log_path}")
        return None
    except Exception as e:
        print(f"  ❌ 文字起こし実行で例外: {type(e).__name__}")
        log_path = write_transcribe_debug_log(audio_path, cmd, error=e, note="unexpected_exception")
        print(f"  📝 デバッグログ: {log_path}")
        return None
    if result.returncode != 0:
        print(f"  Whisperエラー: {result.stderr[:200]}")
        log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="non_zero_exit")
        print(f"  📝 デバッグログ: {log_path}")
        return None
    if "No such file or directory: 'ffmpeg'" in (result.stdout + result.stderr):
        print("  ❌ ffmpeg が見つからず音声読み込みに失敗しました。")
        log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="ffmpeg_missing_inside_mlx_whisper")
        print(f"  📝 デバッグログ: {log_path}")
        return None
    txt_path = out_dir / (audio_path.stem + ".txt")
    if txt_path.exists():
        transcript = txt_path.read_text(encoding="utf-8").strip()
        if transcript:
            return transcript
        log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="empty_transcript_primary_txt")
        print(f"  ⚠️ 文字起こし結果が空です。デバッグログ: {log_path}")
        return None
    # ディレクトリ内のtxtを探す
    txts = list(out_dir.glob("*.txt"))
    if txts:
        transcript = txts[0].read_text(encoding="utf-8").strip()
        if transcript:
            return transcript
        log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="empty_transcript_fallback_txt")
        print(f"  ⚠️ 文字起こし結果が空です。デバッグログ: {log_path}")
        return None
    log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="txt_not_found")
    print(f"  ⚠️ 文字起こし結果ファイルが見つかりません。デバッグログ: {log_path}")
    return None

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
        unchanged = old and all(old.get(k) == e[k] for k in ("title", "date", "duration", "status", "notionPageId"))
        e["updatedAt"] = (old.get("updatedAt") or now_iso) if unchanged else now_iso

    entries.sort(key=lambda e: e["date"], reverse=True)

    MINUTES_INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = MINUTES_INDEX_PATH.with_name(MINUTES_INDEX_PATH.name + ".tmp")
    tmp.write_text(json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(MINUTES_INDEX_PATH)
    return entries

def push_minutes_index():
    repo = str(MINUTES_INDEX_PATH.parent)
    try:
        subprocess.run(["git", "-C", repo, "add", str(MINUTES_INDEX_PATH)],
                       check=True, capture_output=True, text=True)
        if subprocess.run(["git", "-C", repo, "diff", "--cached", "--quiet"],
                          capture_output=True, text=True).returncode == 0:
            print("    差分なしのためcommitをスキップ"); return
        msg = f"chore: update minutes index ({datetime.now(JST).strftime('%Y-%m-%d %H:%M')})"
        subprocess.run(["git", "-C", repo, "commit", "-m", msg], check=True, capture_output=True, text=True)
        subprocess.run(["git", "-C", repo, "push"], check=True, capture_output=True, text=True, timeout=180)
        print("    ✅ git push 完了")
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
    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    print(f"\n{'='*60}\nPLAUD→文字起こし→Notion開始: {now}\n{'='*60}")
    if not PLAUD_TOKEN or not NOTION_TOKEN:
        print("❌ トークンが設定されていません"); return

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
    new_files = [f for f in plaud_files if f["id"] not in