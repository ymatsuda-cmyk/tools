#!/usr/bin/env python3
import importlib.util, json, os, re, requests, shlex, shutil, subprocess, sys, tempfile
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

# ── 文字起こし設定 ───────────────────────────────────────────
# medium(769M) → large-v3-turbo(809M) へ。サイズはほぼ同じで精度は大幅に上、
# かつ MLX の GPU 経路では medium より速い。
WHISPER_MODEL = os.environ.get(
    "WHISPER_MODEL", "mlx-community/whisper-large-v3-turbo"
)

# 幻聴(ハルシネーション)抑止＋精度向上のためのデコード設定
DECODE_OPTIONS = {
    "language": "ja",
    # 直前の出力を次のプロンプトに使わない。幻聴の自己増殖ループを断つ最重要設定
    "condition_on_previous_text": False,
    # 尤度の低い区間を温度を上げて再デコードするフォールバック
    "temperature": (0.0, 0.2, 0.4, 0.6, 0.8, 1.0),
    "compression_ratio_threshold": 2.4,
    "logprob_threshold": -1.0,
    "no_speech_threshold": 0.6,
}

# 用語集ファイル（無ければ自動生成される）
GLOSSARY_PATH = Path(os.environ.get(
    "PLAUD_GLOSSARY_PATH",
    str(Path.home() / ".plaud_glossary.json")
))
GLOSSARY_TEMPLATE = {
    "_comment": "hints: initial_promptに渡す正しい表記。replacements: 誤変換→正表記の一括置換",
    "hints": ["ライターム", "JBA", "PoC", "ロボットアーム"],
    "replacements": {
        "ライタイム": "ライターム",
        "ライタームー": "ライターム"
    }
}

JST = timezone(timedelta(hours=9))
_MLX_WHISPER_CMD = None
_MLX_WHISPER_MODULE = None
_RUNTIME_PATH_PREPARED = False
_GLOSSARY_CACHE = None
SCRIPT_DIR = Path(__file__).resolve().parent
DEBUG_LOG_DIR = SCRIPT_DIR / "logs" / "transcribe_debug"

# ★ ローカルの tools リポジトリのクローン先に合わせて変更してください
MINUTES_INDEX_PATH = Path(os.environ.get(
    "MINUTES_INDEX_PATH",
    str(Path.home() / "tools" / "data" / "minutes" / "index.json")
))
GIT_AUTO_PUSH = False
#GIT_AUTO_PUSH = os.environ.get("MINUTES_GIT_PUSH", "0") == "1"

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

# ── 用語集 ───────────────────────────────────────────────────
def load_glossary():
    global _GLOSSARY_CACHE
    if _GLOSSARY_CACHE is not None:
        return _GLOSSARY_CACHE
    if not GLOSSARY_PATH.exists():
        try:
            GLOSSARY_PATH.parent.mkdir(parents=True, exist_ok=True)
            GLOSSARY_PATH.write_text(
                json.dumps(GLOSSARY_TEMPLATE, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8")
            print(f"    ℹ️ 用語集を作成しました: {GLOSSARY_PATH}")
        except OSError:
            pass
        _GLOSSARY_CACHE = dict(GLOSSARY_TEMPLATE)
        return _GLOSSARY_CACHE
    try:
        data = json.loads(GLOSSARY_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        print(f"    ⚠️ 用語集を読めませんでした({type(e).__name__})。既定値を使用します")
        data = dict(GLOSSARY_TEMPLATE)
    data.setdefault("hints", [])
    data.setdefault("replacements", {})
    _GLOSSARY_CACHE = data
    return _GLOSSARY_CACHE

def build_initial_prompt():
    """initial_prompt は先頭ウィンドウへのヒント。効果が薄れるので入れ過ぎない。"""
    hints = [h for h in load_glossary().get("hints", []) if h]
    if not hints:
        return None
    prompt = "、".join(hints[:30])
    return f"以下は日本語の会議音声です。次の固有名詞が登場します: {prompt}。"

def apply_replacements(text):
    """毎回同じ誤り方をする語は機械的に置換する。AI判断より確実で速い。"""
    repl = load_glossary().get("replacements", {})
    if not text or not repl:
        return text, 0
    count = 0
    # 長い語から置換して部分一致による取りこぼしを防ぐ
    for wrong in sorted(repl, key=len, reverse=True):
        right = repl[wrong]
        if not wrong or wrong == right:
            continue
        n = text.count(wrong)
        if n:
            text = text.replace(wrong, right)
            count += n
    return text, count

def detect_repetition(text, min_len=8, threshold=6):
    """同一フレーズの連続を検出。幻聴ループの早期警告。"""
    if not text:
        return None
    m = re.search(r"(.{%d,40}?)\1{%d,}" % (min_len, threshold), text)
    return m.group(1) if m else None

# ── PLAUD ────────────────────────────────────────────────────
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

# ── 実行環境の解決 ───────────────────────────────────────────
def resolve_mlx_whisper_module():
    """Python API が使えるなら最優先。全デコードオプションを確実に渡せる。"""
    global _MLX_WHISPER_MODULE
    if _MLX_WHISPER_MODULE is not None:
        return _MLX_WHISPER_MODULE or None
    try:
        import mlx_whisper  # noqa: F401
        _MLX_WHISPER_MODULE = mlx_whisper
    except Exception:
        _MLX_WHISPER_MODULE = False
    return _MLX_WHISPER_MODULE or None

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
        f"model: {WHISPER_MODEL}",
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

# ── 文字起こし ───────────────────────────────────────────────
def _transcribe_via_module(audio_path):
    """Python API 経路。全デコードオプションを確実に反映できる。"""
    mlx_whisper = resolve_mlx_whisper_module()
    if mlx_whisper is None:
        return None, "module_unavailable"
    opts = dict(DECODE_OPTIONS)
    initial_prompt = build_initial_prompt()
    if initial_prompt:
        opts["initial_prompt"] = initial_prompt
    try:
        result = mlx_whisper.transcribe(
            str(audio_path),
            path_or_hf_repo=WHISPER_MODEL,
            **opts,
        )
    except Exception as e:
        log_path = write_transcribe_debug_log(
            audio_path, ["python:mlx_whisper.transcribe"],
            error=e, note="module_transcribe_failed")
        print(f"  ⚠️ Python API での文字起こしに失敗({type(e).__name__})。CLIを試します")
        print(f"  📝 デバッグログ: {log_path}")
        return None, "module_failed"
    return (result.get("text") or "").strip(), None

def _transcribe_via_cli(audio_path):
    """CLI 経路。出力txtは音声ファイル名に対応するものだけを読む。"""
    audio_path = Path(audio_path)
    # 出力先を専用ディレクトリに分離する。
    # 旧実装は tmpdir 直下に出力し、失敗時に out_dir.glob("*.txt")[0] で
    # 「前のファイルのtxt」を拾って別の会議の内容を登録する事故があり得た。
    out_dir = audio_path.parent / f"_stt_{audio_path.stem}"
    out_dir.mkdir(parents=True, exist_ok=True)

    cmd_prefix = resolve_mlx_whisper_cmd()
    if not cmd_prefix:
        print("  ❌ mlx_whisper が見つかりません。")
        print("     例: pip install mlx-whisper")
        log_path = write_transcribe_debug_log(audio_path, ["mlx_whisper"], note="mlx_whisper_not_found")
        print(f"  📝 デバッグログ: {log_path}")
        return None

    temps = ",".join(str(t) for t in DECODE_OPTIONS["temperature"])
    cmd = [
        *cmd_prefix, str(audio_path),
        "--model", WHISPER_MODEL,
        "--output-format", "txt",
        "--output-dir", str(out_dir),
        "--language", "ja",
        "--condition-on-previous-text", "False",
        "--temperature", temps,
        "--compression-ratio-threshold", str(DECODE_OPTIONS["compression_ratio_threshold"]),
        "--logprob-threshold", str(DECODE_OPTIONS["logprob_threshold"]),
        "--no-speech-threshold", str(DECODE_OPTIONS["no_speech_threshold"]),
    ]
    initial_prompt = build_initial_prompt()
    if initial_prompt:
        cmd += ["--initial-prompt", initial_prompt]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                timeout=1800, cwd=str(out_dir), env=os.environ.copy())
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
    if not txt_path.exists():
        # 専用ディレクトリなので、ここにあるtxtは必ずこの音声由来
        candidates = list(out_dir.glob("*.txt"))
        if not candidates:
            log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="txt_not_found")
            print(f"  ⚠️ 文字起こし結果ファイルが見つかりません。デバッグログ: {log_path}")
            return None
        txt_path = candidates[0]

    transcript = txt_path.read_text(encoding="utf-8").strip()
    if not transcript:
        log_path = write_transcribe_debug_log(audio_path, cmd, result=result, note="empty_transcript")
        print(f"  ⚠️ 文字起こし結果が空です。デバッグログ: {log_path}")
        return None
    return transcript

def transcribe(audio_path):
    audio_path = Path(audio_path)
    ensure_runtime_path()

    if not resolve_ffmpeg_cmd():
        print("  ❌ ffmpeg が見つかりません。")
        print("     例: brew install ffmpeg")
        log_path = write_transcribe_debug_log(audio_path, ["ffmpeg"], note="ffmpeg_not_found")
        print(f"  📝 デバッグログ: {log_path}")
        return None

    print(f"  文字起こし実行中... (model={WHISPER_MODEL})")
    transcript, _ = _transcribe_via_module(audio_path)
    if not transcript:
        transcript = _transcribe_via_cli(audio_path)
    if not transcript:
        return None

    # 幻聴ループの検出（無音区間で定型句を連呼する Whisper 特有の失敗）
    looped = detect_repetition(transcript)
    if looped:
        print(f"  ⚠️ 同一フレーズの連続を検出: 「{looped[:30]}」")
        print("     録音冒頭の無音や極端に小さい音量が原因のことが多いです")

    transcript, n = apply_replacements(transcript)
    if n:
        print(f"  ✏️ 用語辞書で {n}箇所を補正")
    return transcript

# ── Notion ───────────────────────────────────────────────────
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
        "permissions": _multi_select_tags(props, "権限"),
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
        unchanged = old and all(old.get(k) == e[k] for k in ("title", "date", "duration", "status", "tags", "permissions", "notionPageId"))
        e["updatedAt"] = (old.get("updatedAt") or now_iso) if unchanged else now_iso

    entries.sort(key=lambda e: e["date"], reverse=True)

    MINUTES_INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = MINUTES_INDEX_PATH.with_name(MINUTES_INDEX_PATH.name + ".tmp")
    tmp.write_text(json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(MINUTES_INDEX_PATH)
    return entries

def push_minutes_index():
    repo = str(Path.home() / "tools")
    target_file = "data/minutes/index.json"
    branch = "main"

    try:
        subprocess.run(["git", "-C", repo, "add", target_file],
                       check=True, capture_output=True, text=True)

        if subprocess.run(["git", "-C", repo, "diff", "--cached", "--quiet"],
                          capture_output=True, text=True).returncode == 0:
            print("    差分なしのためcommitをスキップ")
            return

        msg = (f"chore: update minutes index "
               f"({datetime.now(JST).strftime('%Y-%m-%d %H:%M')})")
        subprocess.run(["git", "-C", repo, "commit", "-m", msg],
                       check=True, capture_output=True, text=True)
        subprocess.run(["git", "-C", repo, "push", "origin", branch],
                       check=True, capture_output=True, text=True, timeout=180)
        print("    ✅ git push 完了")

    except subprocess.CalledProcessError as e:
        print(f"    ⚠️ git操作に失敗: {(e.stderr or '')[:500]}")
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

    load_glossary()

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
    for i, f in enumerate(new_files, 1):
        file_id = f["id"]
        filename = f.get("fullname", f"{file_id}.ogg")
        name = f.get("filename", file_id)
        print(f"\n  [{i}/{len(new_files)}] {name}")

        temp_url = get_download_url(file_id)
        if not temp_url:
            print(f"  ❌ URL取得失敗。スキップ"); continue

        # ファイルごとに独立した一時ディレクトリを使う。
        # 共有tmpdirだと前のファイルのtxtを拾う事故が起き得るため。
        with tempfile.TemporaryDirectory() as tmpdir:
            audio_path = Path(tmpdir) / filename
            print(f"  → ダウンロード中... ({filename})")
            if not download_audio(temp_url, str(audio_path)):
                print(f"  ❌ ダウンロード失敗。スキップ"); continue
            print(f"  ✅ {audio_path.stat().st_size/1024/1024:.1f} MB")

            transcript = transcribe(str(audio_path))

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
