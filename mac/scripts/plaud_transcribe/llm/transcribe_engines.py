#!/usr/bin/env python3
"""文字起こしエンジン切り替え層

settings.json の "label" でエンジンを選ぶ。
  type = "whisper"   … mlx-whisper のプレーンテキストをそのまま返す
  type = "qwen3-asr" … Qwen3-ASR（mlx-qwen3-asr）で区間を取得したうえで
                       ①声紋登録 ②話者自動認識 ③フィラー除去
                       ④表記統一 ⑤フォーマット を行う

出力例（qwen3-asr）:
    [ 0.00] 松田: 本日はお集まりいただきありがとうございます
"""
import importlib.util
import json
import os
import shlex
import shutil
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import llm_refine          # noqa: E402  フィラー除去・表記統一（ルールベース）
import speaker_diarize     # noqa: E402

JST = timezone(timedelta(hours=9))
DEBUG_LOG_DIR = SCRIPT_DIR / "logs" / "transcribe_debug"
DEFAULT_SETTINGS_PATH = Path(os.environ.get("TRANSCRIBE_SETTINGS", str(SCRIPT_DIR / "settings.json")))

_MLX_WHISPER_CMD = None
_RUNTIME_PATH_PREPARED = False
_SETTINGS_CACHE = None


# ------------------------------------------------------------------- 設定読込
def load_settings(path=None):
    global _SETTINGS_CACHE
    if path is None and _SETTINGS_CACHE is not None:
        return _SETTINGS_CACHE
    p = Path(path or DEFAULT_SETTINGS_PATH).expanduser()
    if not p.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {p}")
    settings = json.loads(p.read_text(encoding="utf-8"))
    settings["_path"] = str(p)
    if path is None:
        _SETTINGS_CACHE = settings
    return settings


def get_engine(settings, label=None):
    label = label or os.environ.get("TRANSCRIBE_LABEL") or settings.get("defaultLabel")
    engines = settings.get("engines", [])
    if not engines:
        raise ValueError("settings.json に engines がありません")
    if not label:
        return engines[0]
    for e in engines:
        if e.get("label") == label or e.get("id") == label:
            return e
    labels = ", ".join(e.get("label", "?") for e in engines)
    raise ValueError(f"label '{label}' が見つかりません。利用可能: {labels}")


def load_dictionary(settings):
    rel = settings.get("normalizeDictPath")
    if not rel:
        return {}
    p = Path(rel).expanduser()
    if not p.is_absolute():
        p = Path(settings.get("_path", str(SCRIPT_DIR))).parent / p
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


# ------------------------------------------------------------- whisper実行基盤
def resolve_mlx_whisper_cmd():
    global _MLX_WHISPER_CMD
    if _MLX_WHISPER_CMD is not None:
        return _MLX_WHISPER_CMD

    cli_path = shutil.which("mlx_whisper")
    if cli_path:
        _MLX_WHISPER_CMD = [cli_path]
        return _MLX_WHISPER_CMD

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
    for candidate in (Path("/opt/homebrew/bin/ffmpeg"), Path("/usr/local/bin/ffmpeg")):
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
        lines.extend(["error_type:", type(error).__name__, "error_message:", str(error)])
    if result is not None:
        lines.extend([f"returncode: {result.returncode}", "stdout:", result.stdout or "",
                      "stderr:", result.stderr or ""])
    files = sorted(p.name for p in Path(audio_path).parent.glob("*.*")
                   if p.suffix in (".txt", ".json"))
    lines.extend(["output_files_in_dir:", "\n".join(files) if files else "(none)"])
    log_path.write_text("\n".join(lines), encoding="utf-8")
    return log_path


def _run_whisper(audio_path, model, language, output_format):
    """mlx_whisperを実行し (成功フラグ, 出力ファイルパス or None) を返す"""
    audio_path = Path(audio_path)
    out_dir = audio_path.parent
    ensure_runtime_path()

    if not resolve_ffmpeg_cmd():
        print("  ❌ ffmpeg が見つかりません。（例: brew install ffmpeg）")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, ['ffmpeg'], note='ffmpeg_not_found')}")
        return None

    cmd_prefix = resolve_mlx_whisper_cmd()
    if not cmd_prefix:
        print("  ❌ mlx_whisper が見つかりません。（例: pip install mlx-whisper）")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, ['mlx_whisper'], note='mlx_whisper_not_found')}")
        return None

    cmd = [
        *cmd_prefix, str(audio_path),
        "--model", model,
        "--output-format", output_format,
        "--output-dir", str(out_dir),
        "--language", language,
        "--condition-on-previous-text", "False",
    ]
    print("  文字起こし実行中...")
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800,
                                cwd=str(out_dir), env=os.environ.copy())
    except FileNotFoundError as e:
        print("  ❌ mlx_whisper コマンドを起動できませんでした。")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, error=e, note='mlx_whisper_exec_not_found')}")
        return None
    except subprocess.TimeoutExpired as e:
        print("  ❌ 文字起こしがタイムアウトしました。")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, error=e, note='timeout')}")
        return None
    except Exception as e:  # noqa: BLE001
        print(f"  ❌ 文字起こし実行で例外: {type(e).__name__}")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, error=e, note='unexpected_exception')}")
        return None

    if result.returncode != 0:
        print(f"  Whisperエラー: {result.stderr[:200]}")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, result=result, note='non_zero_exit')}")
        return None
    if "No such file or directory: 'ffmpeg'" in (result.stdout + result.stderr):
        print("  ❌ ffmpeg が見つからず音声読み込みに失敗しました。")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, result=result, note='ffmpeg_missing_inside_mlx_whisper')}")
        return None

    ext = "." + output_format
    target = out_dir / (audio_path.stem + ext)
    if target.exists():
        return target
    candidates = list(out_dir.glob("*" + ext))
    if candidates:
        return candidates[0]
    print(f"  ⚠️ 文字起こし結果ファイル({ext})が見つかりません。")
    print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, result=result, note='output_not_found')}")
    return None


def whisper_text(audio_path, model, language="ja"):
    path = _run_whisper(audio_path, model, language, "txt")
    if not path:
        return None
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        print("  ⚠️ 文字起こし結果が空です。")
        return None
    return text


# --------------------------------------------------------------- フォーマット
def format_transcript(segments, unknown="話者不明"):
    lines = []
    for s in segments:
        text = (s.get("text") or "").strip()
        if not text:
            continue
        lines.append(f"[{float(s.get('start', 0.0)):5.2f}] {s.get('speaker') or unknown}: {text}")
    return "\n".join(lines)


# ------------------------------------------------------------------- メインAPI
def transcribe(audio_path, label=None, settings=None, verbose=True):
    """labelに応じて文字起こしを行い、文字列を返す（失敗時はNone）"""
    settings = settings or load_settings()
    engine = get_engine(settings, label)
    etype = engine.get("type")

    if verbose:
        print(f"  エンジン: {engine.get('label')} (type={etype})")

    # whisper：プレーンテキストをそのまま返す
    if etype == "whisper":
        return whisper_text(audio_path,
                            engine.get("model", "mlx-community/whisper-medium-mlx"),
                            engine.get("language", "ja"))

    if etype != "qwen3-asr":
        raise ValueError(f"未対応のtype '{etype}'（label={engine.get('label')}）。"
                         "whisper か qwen3-asr を指定してください")

    # --- qwen3-asr ---------------------------------------------------------
    import qwen3_asr  # noqa: PLC0415  遅延importでMac以外の環境でも読み込める
    if not qwen3_asr.available():
        print("  ❌ mlx-qwen3-asr が未導入です（pip install \"mlx-qwen3-asr[aligner]\"）")
        return None

    context = qwen3_asr.build_context(load_dictionary(settings), engine.get("contextTerms"))
    if context and verbose:
        print(f"  ドメイン語彙: {context[:60]}{'...' if len(context) > 60 else ''}")

    segments, diarized = qwen3_asr.transcribe_segments(audio_path, engine, context=context, verbose=verbose)
    if not segments:
        return None
    if verbose:
        print(f"  ✅ 区間取得: {len(segments)}件{'（話者分離済み）' if diarized else ''}")

    dia = engine.get("diarization", {})
    prefix = dia.get("namePrefix", "担当者")
    db_path = settings.get("speakerDbPath", "~/.plaud_speakers.json")

    # ① 声紋登録 / ② 話者自動認識
    if dia.get("enabled", True) and speaker_diarize.available():
        try:
            if diarized:
                # pyannoteで分離済み → 匿名ラベルに声紋DBの名前を割り当てるだけ
                segments, mapping = speaker_diarize.name_labeled_speakers(
                    audio_path, segments, db_path=db_path,
                    match_threshold=dia.get("matchThreshold", 0.72),
                    min_segment_sec=dia.get("minSegmentSec", 0.7),
                    name_prefix=prefix,
                )
            else:
                # pyannote不在 → 自前クラスタリングで話者を分ける
                segments, mapping = speaker_diarize.assign_speakers(
                    audio_path, segments, db_path=db_path,
                    cluster_threshold=dia.get("clusterThreshold", 0.62),
                    match_threshold=dia.get("matchThreshold", 0.72),
                    min_segment_sec=dia.get("minSegmentSec", 0.7),
                    name_prefix=prefix,
                )
            if verbose:
                print(f"  ✅ 話者判定: {', '.join(sorted(set(mapping.values())))}")
        except Exception as e:  # noqa: BLE001
            if verbose:
                print(f"  ⚠️ 声紋判定に失敗: {type(e).__name__}: {e}")
    elif verbose:
        print("  ⚠️ resemblyzer未導入のため声紋の名前付けをスキップします（pip install resemblyzer）")

    # ③ フィラー除去 / ④ 表記統一（ルールベース）
    segments = llm_refine.refine_segments(
        segments,
        fillers=settings.get("fillers", []),
        dictionary=load_dictionary(settings),
        verbose=verbose,
    )

    # ⑤ フォーマット
    return format_transcript(segments) or None


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="音声ファイルを文字起こしする")
    ap.add_argument("audio")
    ap.add_argument("--label")
    ap.add_argument("--settings")
    args = ap.parse_args()
    st = load_settings(args.settings)
    out = transcribe(args.audio, label=args.label, settings=st)
    print(out or "(失敗)")
