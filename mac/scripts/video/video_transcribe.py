#!/usr/bin/env python3
"""Notionの動画DBを対象に、URL先の動画から文字起こしを行いページ本文へ書き込む。

対象の指定方法:
  --status 空欄 再取得   状態が空欄／再取得のページをまとめて処理（既定）
  --page-id <ID>        特定の1ページだけを処理

処理の流れ:
  1. Notionから対象ページを取得
  2. URLプロパティから動画IDを解決
  3. 字幕API（youtube-transcript-api）で文字起こしを取得
     取れない場合は yt-dlp で音声を落とし、mlx-whisper で文字起こし（既定でON）
  4. 既存の本文ブロックをアーカイブしてから新しい文字起こしを追記
  5. 動画タイトル・サムネイル・原文文字数・状態を更新

環境変数:
  NOTION_TOKEN     Notion Integration Token（必須）
  VIDEO_ENV_FILE   環境変数を読み込むファイルのパス（既定: ~/.video_notion_sync.env）
  VIDEO_DB_ID      対象データベースID（既定: 📚動画DB）

Whisperフォールバックを使うには:
  pip install mlx-whisper
  brew install ffmpeg yt-dlp   （未導入の場合）

使い方:
    python3 video_transcribe_notion.py --dry-run
    python3 video_transcribe_notion.py
    python3 video_transcribe_notion.py --limit 1
    python3 video_transcribe_notion.py --no-whisper       # Whisperフォールバックを使わない
    python3 video_transcribe_notion.py --page-id 3d00e7a535dc81e1b57bfc93f65019d2
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

JST = timezone(timedelta(hours=9))
SCRIPT_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------------- 環境変数

ENV_FILE = Path(os.environ.get("VIDEO_ENV_FILE", str(Path.home() / ".video_notion_sync.env")))


def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env()

NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
# 📚動画DB
VIDEO_DB_ID = os.environ.get("VIDEO_DB_ID", "3630e7a535dc8154ac62d41f7611540f")

STATUS_EMPTY = "空欄"
WHISPER_MODEL_DEFAULT = "mlx-community/whisper-large-v2-mlx"
WHISPER_LANGUAGE_DEFAULT = "ja"

# ---------------------------------------------------------------- Notion

NOTION_API = "https://api.notion.com/v1"


def notion_headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }


def get_valid_status_options():
    """DBの「状態」プロパティに定義済みの選択肢一覧を取得する。"""
    resp = requests.get(f"{NOTION_API}/databases/{VIDEO_DB_ID}",
                        headers=notion_headers(), timeout=60)
    if resp.status_code != 200:
        print(f"    ⚠️ DBスキーマ取得失敗: {resp.status_code} {resp.text[:200]}")
        return None
    prop = resp.json().get("properties", {}).get("状態", {})
    options = prop.get("select", {}).get("options", [])
    return {o["name"] for o in options}


def build_status_filter(statuses):
    """状態フィルタを組み立てる。'空欄' は is_empty に変換する。

    DBに存在しない選択肢が指定された場合は除外し、警告を出す
    （Notion APIは未定義の選択肢を equals に渡すと 400 を返すため）。
    """
    valid_options = get_valid_status_options()
    conds = []
    for s in statuses:
        if s == STATUS_EMPTY:
            conds.append({"property": "状態", "select": {"is_empty": True}})
            continue
        if valid_options is not None and s not in valid_options:
            print(f"    ⚠️ 状態「{s}」はこのDBに存在しないため無視します"
                  f"（利用可能: {', '.join(sorted(valid_options))}）")
            continue
        conds.append({"property": "状態", "select": {"equals": s}})
    if not conds:
        return None
    if len(conds) == 1:
        return conds[0]
    return {"or": conds}


def query_pages_by_status(statuses):
    filter_ = build_status_filter(statuses)
    if filter_ is None:
        print("    ⚠️ 有効な状態フィルタが無いため検索を中止します。")
        return []
    pages, has_more, cursor = [], True, None
    while has_more:
        payload = {"page_size": 100, "filter": filter_}
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(f"{NOTION_API}/databases/{VIDEO_DB_ID}/query",
                             headers=notion_headers(), json=payload, timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ Notion検索失敗: {resp.status_code} {resp.text[:300]}")
            break
        data = resp.json()
        pages.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return pages


def fetch_page(page_id):
    resp = requests.get(f"{NOTION_API}/pages/{page_id}", headers=notion_headers(), timeout=60)
    if resp.status_code != 200:
        print(f"  ❌ ページ取得失敗: {resp.status_code} {resp.text[:300]}")
        return None
    return resp.json()


def page_title(page):
    items = page.get("properties", {}).get("動画タイトル", {}).get("title", [])
    return items[0].get("plain_text", "") if items else "(無題)"


def page_url_value(page):
    """URLプロパティ（url型）を読む。空ならタイトルにURLが入っているケースを拾う。"""
    props = page.get("properties", {})
    # REST APIでは "userDefined:URL" ではなく "URL" がキー
    url_val = (props.get("URL") or {}).get("url") or ""
    if url_val:
        return url_val.strip()
    title = page_title(page)
    if title.startswith("http"):
        return title.strip()
    return ""


def get_all_children(block_id):
    children, has_more, cursor = [], True, None
    while has_more:
        url = f"{NOTION_API}/blocks/{block_id}/children?page_size=100"
        if cursor:
            url += f"&start_cursor={cursor}"
        resp = requests.get(url, headers=notion_headers(), timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ ブロック取得失敗: {resp.status_code} {resp.text[:200]}")
            break
        data = resp.json()
        children.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return children


def archive_transcript_children(page_id, keep_video_block=True):
    """既存の本文をアーカイブする。動画埋め込みブロックは残す。"""
    children = get_all_children(page_id)
    ok = total = 0
    for block in children:
        if keep_video_block and block.get("type") in ("video", "embed", "bookmark"):
            continue
        total += 1
        resp = requests.patch(f"{NOTION_API}/blocks/{block['id']}",
                              headers=notion_headers(), json={"archived": True}, timeout=60)
        if resp.status_code == 200:
            ok += 1
        else:
            print(f"    ⚠️ ブロック削除失敗: {resp.status_code} {resp.text[:200]}")
    return ok, total


def text_to_blocks(text):
    """2000文字ごとに分割した段落ブロックへ変換する。"""
    blocks = []
    for para in [p.strip() for p in text.split("\n") if p.strip()]:
        while para:
            chunk, para = para[:2000], para[2000:]
            blocks.append({"object": "block", "type": "paragraph",
                           "paragraph": {"rich_text": [{"type": "text", "text": {"content": chunk}}]}})
    return blocks or [{"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": [{"type": "text",
                                                    "text": {"content": "（文字起こし結果なし）"}}]}}]


def append_transcript(page_id, source_label, transcript_text, engine_label):
    ts = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    children = [
        {"object": "block", "type": "heading_2", "heading_2": {"rich_text": [
            {"type": "text", "text": {"content": f"🎬 {source_label}"}}]}},
        {"object": "block", "type": "callout", "callout": {
            "icon": {"type": "emoji", "emoji": "📝"},
            "rich_text": [{"type": "text", "text": {
                "content": f"文字起こし: {ts}（取得方法: {engine_label}）"}}]}},
    ]
    children.extend(text_to_blocks(transcript_text))

    batch, remaining = children[:100], children[100:]
    resp = requests.patch(f"{NOTION_API}/blocks/{page_id}/children",
                          headers=notion_headers(), json={"children": batch}, timeout=60)
    if resp.status_code != 200:
        print(f"    ❌ 本文追記失敗: {resp.status_code} {resp.text[:300]}")
        return False
    while remaining:
        batch, remaining = remaining[:100], remaining[100:]
        r = requests.patch(f"{NOTION_API}/blocks/{page_id}/children",
                           headers=notion_headers(), json={"children": batch}, timeout=60)
        if r.status_code != 200:
            print(f"    ⚠️ 追記の途中で失敗: {r.status_code} {r.text[:200]}")
            return False
    return True


def update_page_props(page_id, *, title=None, url=None, thumbnail=None,
                      char_count=None, status=None):
    props = {}
    if title:
        props["動画タイトル"] = {"title": [{"text": {"content": title[:2000]}}]}
    if url:
        props["URL"] = {"url": url}
    if thumbnail:
        props["サムネイル"] = {"url": thumbnail}
    if char_count is not None:
        props["原文文字数"] = {"number": char_count}
    if status:
        props["状態"] = {"select": {"name": status}}
    if not props:
        return True
    resp = requests.patch(f"{NOTION_API}/pages/{page_id}",
                          headers=notion_headers(), json={"properties": props}, timeout=60)
    if resp.status_code != 200:
        print(f"    ⚠️ プロパティ更新失敗: {resp.status_code} {resp.text[:300]}")
        return False
    return True


# ---------------------------------------------------------------- URL解決

YT_PATTERNS = [
    re.compile(r"(?:youtube\.com/watch\?(?:.*&)?v=)([A-Za-z0-9_-]{11})"),
    re.compile(r"(?:youtu\.be/)([A-Za-z0-9_-]{11})"),
    re.compile(r"(?:youtube\.com/(?:embed|shorts|live)/)([A-Za-z0-9_-]{11})"),
]


def video_id_from_url(url):
    for pat in YT_PATTERNS:
        m = pat.search(url or "")
        if m:
            return m.group(1)
    return None


def resolve_redirect(url):
    """share.google などの短縮URLを展開する。"""
    try:
        resp = requests.head(url, allow_redirects=True, timeout=30,
                             headers={"User-Agent": "Mozilla/5.0"})
        return resp.url or url
    except requests.RequestException:
        return url


# ---------------------------------------------------------------- メタ情報

def fetch_metadata(url):
    """yt-dlp --dump-json でタイトルとサムネイルだけ取得する（音声DLはしない）。"""
    try:
        result = subprocess.run(
            ["yt-dlp", "--dump-json", "--no-warnings", "--skip-download", url],
            capture_output=True, text=True, timeout=120)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return {}
    if result.returncode != 0 or not result.stdout.strip():
        return {}
    try:
        info = json.loads(result.stdout.splitlines()[0])
    except json.JSONDecodeError:
        return {}
    return {
        "title": info.get("title") or "",
        "thumbnail": info.get("thumbnail") or "",
        "duration": info.get("duration"),
    }


def fetch_oembed(video_id):
    """yt-dlpが失敗したとき用の軽量なタイトル取得（oEmbed）。"""
    try:
        resp = requests.get(
            "https://www.youtube.com/oembed",
            params={"url": f"https://www.youtube.com/watch?v={video_id}", "format": "json"},
            timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    except requests.RequestException:
        return {}
    if resp.status_code != 200:
        return {}
    try:
        info = resp.json()
    except ValueError:
        return {}
    return {
        "title": info.get("title") or "",
        "thumbnail": info.get("thumbnail_url") or "",
    }


def resolve_meta(video_id, canonical_url):
    """タイトルとサムネイルを確実に埋める。

    yt-dlp → oEmbed → サムネイルURLの直接組み立て、の順で補完する。
    """
    meta = fetch_metadata(canonical_url)
    if not meta.get("title") or not meta.get("thumbnail"):
        fallback = fetch_oembed(video_id)
        meta.setdefault("title", "")
        meta.setdefault("thumbnail", "")
        if not meta["title"]:
            meta["title"] = fallback.get("title", "")
        if not meta["thumbnail"]:
            meta["thumbnail"] = fallback.get("thumbnail", "")
    if not meta.get("thumbnail"):
        meta["thumbnail"] = f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg"
    return meta


def clean_title(title):
    """ブラウザ由来の「 - YouTube」サフィックスを落とす。"""
    title = (title or "").strip()
    for suffix in (" - YouTube", " – YouTube", " | YouTube"):
        if title.endswith(suffix):
            title = title[: -len(suffix)].strip()
    return title


# ---------------------------------------------------------------- 文字起こし

def format_timecode(seconds: float) -> str:
    """秒を [12:34] / [1:02:03] の表記にする"""
    total = int(seconds)
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


def build_timestamped_lines(fetched, window: int = 30, max_chars: int = 900):
    """
    セグメントをまとめて "[12:34] 本文" の行にする。

    1セグメントは2〜5秒しかないので、そのまま1行ずつ書くと本文が数百ブロックに
    膨らんで、Notionへの書き込み回数も表示も破綻する。約30秒ぶんを1行にまとめる。
    max_chars は Notion の rich_text 1件あたり2000字上限に対する余裕分。
    """
    lines = []
    start = None
    buf = []

    def flush():
        if buf:
            lines.append(f"[{format_timecode(start)}] {''.join(buf).strip()}")

    for seg in fetched:
        # 属性アクセスであることに注意（seg['text'] ではなく seg.text）
        text = getattr(seg, "text", "") or ""
        text = text.replace("\n", " ").strip()
        if not text:
            continue
        seg_start = getattr(seg, "start", 0) or 0
        if start is None:
            start = seg_start
        over_window = seg_start - start >= window
        over_chars = sum(len(b) for b in buf) + len(text) > max_chars
        if buf and (over_window or over_chars):
            flush()
            start = seg_start
            buf = []
        buf.append(text + " ")

    flush()
    return lines


def fetch_captions(video_id, languages=("ja", "ja-JP", "en")):
    """youtube-transcript-api で字幕を取得し、[12:34] 形式の時刻付きテキストにする。

    注意: 新しいAPIではインスタンス化して fetch() を呼ぶ。
    テキスト・開始秒は t.text / t.start（属性アクセス）で取り出す。
    """
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError:
        print("    ⚠️ youtube-transcript-api が未導入です（pip install youtube-transcript-api）")
        return None

    ytt = YouTubeTranscriptApi()
    for lang in languages:
        try:
            fetched = ytt.fetch(video_id, languages=[lang])
        except Exception:
            continue
        lines = build_timestamped_lines(fetched)
        if lines:
            return "\n".join(lines), f"字幕({lang})"
    return None


class _WhisperSeg:
    """build_timestamped_lines がそのまま使えるよう .text / .start を持たせたラッパー"""
    __slots__ = ("text", "start")

    def __init__(self, text, start):
        self.text = text
        self.start = start


def resolve_ffmpeg_cmd():
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg
    for candidate in (Path("/opt/homebrew/bin/ffmpeg"), Path("/usr/local/bin/ffmpeg")):
        if candidate.exists() and os.access(candidate, os.X_OK):
            return str(candidate)
    return None


def download_audio_for_whisper(canonical_url, dest_dir):
    """yt-dlpで音声のみをダウンロードする（変換はmlx-whisper内部のffmpegに任せる）。"""
    out_tmpl = str(Path(dest_dir) / "audio.%(ext)s")
    cmd = ["yt-dlp", "-f", "bestaudio/best", "--no-warnings", "-o", out_tmpl, canonical_url]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)
    except FileNotFoundError:
        print("    ❌ yt-dlp が見つかりません（pip install yt-dlp）")
        return None
    except subprocess.TimeoutExpired:
        print("    ❌ 音声ダウンロードがタイムアウトしました")
        return None
    if result.returncode != 0:
        print(f"    ❌ 音声ダウンロード失敗: {(result.stderr or '')[:300]}")
        return None
    audio_files = sorted(Path(dest_dir).glob("audio.*"))
    return audio_files[0] if audio_files else None


def transcribe_with_whisper(audio_path, model, language):
    """mlx-whisperでローカル文字起こしし、[12:34]付きの行に整形する。"""
    try:
        import mlx_whisper
    except ImportError:
        print("    ❌ mlx-whisper が未導入です（pip install mlx-whisper）")
        return None
    if not resolve_ffmpeg_cmd():
        print("    ❌ ffmpeg が見つかりません（brew install ffmpeg）")
        return None

    try:
        result = mlx_whisper.transcribe(
            str(audio_path),
            path_or_hf_repo=model,
            language=language,
            condition_on_previous_text=False,
            verbose=False,
        )
    except Exception as e:
        print(f"    ❌ Whisper文字起こし失敗: {type(e).__name__}: {e}")
        return None

    segments = result.get("segments") or []
    if segments:
        segs = [_WhisperSeg((s.get("text") or "").strip(), s.get("start") or 0.0)
                for s in segments if (s.get("text") or "").strip()]
    else:
        text = (result.get("text") or "").strip()
        segs = [_WhisperSeg(text, 0.0)] if text else []

    if not segs:
        return None
    lines = build_timestamped_lines(segs)
    return lines or None


def transcribe_video_via_whisper(canonical_url, *, model, language):
    """字幕が無い動画向け：音声DL→mlx-whisperで文字起こし。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        print("    → 音声をダウンロード中（yt-dlp）...")
        audio_path = download_audio_for_whisper(canonical_url, tmpdir)
        if not audio_path:
            return None
        size_mb = audio_path.stat().st_size / 1024 / 1024
        print(f"    ✅ 音声取得 ({size_mb:.1f} MB)")
        print(f"    → Whisperで文字起こし中...（model={model}）")
        lines = transcribe_with_whisper(audio_path, model, language)
        if not lines:
            return None
        model_short = model.split("/")[-1]
        return "\n".join(lines), f"Whisper({model_short})"


# ---------------------------------------------------------------- 本処理

def process_page(page, *, set_status_name, is_retry, whisper_enabled,
                 whisper_model, whisper_language):
    title = page_title(page)
    page_id = page.get("id")
    raw_url = page_url_value(page)

    print(f"  対象: {title[:60]}")
    print(f"        page_id={page_id}")

    if not raw_url:
        print("  ❌ URLが空のためスキップ")
        return False

    url = raw_url
    video_id = video_id_from_url(url)
    if not video_id:
        url = resolve_redirect(raw_url)
        video_id = video_id_from_url(url)
    if not video_id:
        print(f"  ⏭️  YouTube動画ではないためスキップ: {url[:80]}")
        return False

    canonical_url = f"https://www.youtube.com/watch?v={video_id}"

    meta = resolve_meta(video_id, canonical_url)
    resolved_title = clean_title(meta.get("title"))
    if resolved_title:
        print(f"        タイトル: {resolved_title[:60]}")
    else:
        print("        ⚠️ タイトルを取得できませんでした")

    got = fetch_captions(video_id)
    engine_label = None
    transcript = None
    if got:
        transcript, engine_label = got

    if not transcript and whisper_enabled:
        print("    字幕なし → Whisperにフォールバック")
        got = transcribe_video_via_whisper(canonical_url, model=whisper_model,
                                           language=whisper_language)
        if got:
            transcript, engine_label = got

    if not transcript:
        print("  ❌ 文字起こしを取得できませんでした。スキップ")
        return False
    print(f"  ✅ 文字起こし取得 ({len(transcript)}文字 / {engine_label})")

    if is_retry:
        print("    → 既存の本文を削除中...")
        ok, total = archive_transcript_children(page_id)
        print(f"    ✅ {ok}/{total} ブロックを削除")

    source_label = resolved_title or title or canonical_url
    if not append_transcript(page_id, source_label, transcript, engine_label):
        return False
    print("  ✅ 本文に追記")

    # タイトルは取得できたら常に上書きする（「 - YouTube」付きの旧タイトルも整う）
    new_title = resolved_title or None
    update_page_props(
        page_id,
        title=new_title,
        url=canonical_url if raw_url != canonical_url else None,
        thumbnail=meta.get("thumbnail") or None,
        char_count=len(transcript),
        status=set_status_name,
    )
    if new_title:
        print(f"  ✅ 動画タイトルを更新: {new_title[:60]}")
    if set_status_name:
        print(f"  ✅ 状態を「{set_status_name}」に更新")
    return True


def main():
    ap = argparse.ArgumentParser(description="Notion動画DBの文字起こし")
    ap.add_argument("--page-id", help="対象のNotionページID（--statusより優先）")
    ap.add_argument("--status", nargs="+", default=[STATUS_EMPTY, "再取得"],
                    help=f"対象とする状態（既定: {STATUS_EMPTY} 再取得）")
    ap.add_argument("--set-status", default="完了",
                    help="処理完了後に設定する状態（既定: 完了。空文字で更新しない）")
    ap.add_argument("--limit", type=int, help="処理する件数の上限")
    ap.add_argument("--no-whisper", action="store_true",
                    help="字幕が無いときWhisperにフォールバックしない")
    ap.add_argument("--whisper-model", default=WHISPER_MODEL_DEFAULT,
                    help=f"mlx-whisperのモデル（既定: {WHISPER_MODEL_DEFAULT}）")
    ap.add_argument("--whisper-language", default=WHISPER_LANGUAGE_DEFAULT,
                    help=f"Whisperの言語コード（既定: {WHISPER_LANGUAGE_DEFAULT}）")
    ap.add_argument("--dry-run", action="store_true", help="対象一覧を表示するだけ")
    args = ap.parse_args()

    if not NOTION_TOKEN:
        print("❌ NOTION_TOKEN が設定されていません")
        return

    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    print(f"\n{'='*60}\n動画DB→文字起こし開始: {now}\n{'='*60}")

    if args.page_id:
        page = fetch_page(args.page_id)
        targets = [page] if page else []
        is_retry_default = True
    else:
        targets = query_pages_by_status(args.status)
        is_retry_default = False

    if not targets:
        print("対象ページが見つかりませんでした。")
        return

    if args.limit:
        targets = targets[:args.limit]

    print(f"\n対象: {len(targets)}件")
    for p in targets:
        st = (p.get("properties", {}).get("状態", {}).get("select") or {}).get("name") or "(空欄)"
        print(f"  - [{st}] {page_title(p)[:60]}")

    if args.dry_run:
        print("\n--dry-run のため処理は行いません。")
        return

    set_status_name = args.set_status or None
    ok_count = 0
    for i, page in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}]")
        st = (page.get("properties", {}).get("状態", {}).get("select") or {}).get("name") or ""
        is_retry = is_retry_default or st == "再取得"
        try:
            if process_page(page, set_status_name=set_status_name, is_retry=is_retry,
                            whisper_enabled=not args.no_whisper,
                            whisper_model=args.whisper_model,
                            whisper_language=args.whisper_language):
                ok_count += 1
        except Exception as e:
            print(f"  ❌ 例外: {type(e).__name__}: {e}")

    print(f"\n{'='*60}\n完了: {ok_count}/{len(targets)}件  "
          f"{datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S JST')}\n{'='*60}\n")


if __name__ == "__main__":
    main()