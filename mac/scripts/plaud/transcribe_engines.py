#!/usr/bin/env python3
"""文字起こしエンジン切り替え層

settings.json の "label" でエンジンを選ぶ。

  type = "whisper"   … mlx-whisper で文字起こしし、
                       ③フィラー除去 ④表記統一 を適用して返す
  type = "qwen3-asr" … Qwen3-ASR（mlx-qwen3-asr）で区間を取得したうえで
                       ①声紋登録 ②話者自動認識 ③フィラー除去
                       ④表記統一 ⑤フォーマット を行う

出力例（qwen3-asr）:
    [ 0.00] 松田: 本日はお集まりいただきありがとうございます
"""
import importlib.util
import json
import os
import re
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

# settings.json に decode が無い場合の既定値。
# condition_on_previous_text=False は無音区間の幻聴が自己増殖するのを断つ最重要設定。
DEFAULT_DECODE = {
    "condition_on_previous_text": False,
    "temperature": [0.0, 0.2, 0.4, 0.6, 0.8, 1.0],
    "compression_ratio_threshold": 2.4,
    "logprob_threshold": -1.0,
    "no_speech_threshold": 0.6,
}
DEFAULT_WHISPER_MODEL = "mlx-community/whisper-large-v3-turbo"

_MLX_WHISPER_CMD = None
_MLX_WHISPER_MODULE = None
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


def choose_engine(settings, label=None, interactive=True):
    """labelが未指定の場合、対話端末なら選択メニューを表示する。

    cron等の非対話実行（標準入力がttyでない）では、従来通り
    settings.json の defaultLabel / 環境変数 TRANSCRIBE_LABEL にフォールバックする。
    """
    if label or os.environ.get("TRANSCRIBE_LABEL"):
        return get_engine(settings, label)

    engines = settings.get("engines", [])
    if not engines:
        raise ValueError("settings.json に engines がありません")

    if not (interactive and sys.stdin.isatty()):
        return get_engine(settings, None)

    default_label = settings.get("defaultLabel")
    default_idx = next((i for i, e in enumerate(engines, 1)
                       if e.get("label") == default_label), 1)

    print("文字起こしエンジンを選択してください:")
    for i, e in enumerate(engines, 1):
        mark = "（既定）" if e.get("label") == default_label else ""
        print(f"  {i}. {e.get('label')}  (type={e.get('type')}){mark}")
    raw = input(f"番号またはlabelを入力 [{default_idx}]: ").strip()
    if not raw:
        return engines[default_idx - 1]
    if raw.isdigit() and 1 <= int(raw) <= len(engines):
        return engines[int(raw) - 1]
    return get_engine(settings, raw)


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


def build_initial_prompt(dictionary, extra_terms=None, limit=30):
    """normalize_dict の「正しい表記」側を initial_prompt のヒントに使う。

    initial_prompt は先頭ウィンドウへのヒントで効果が限定的なため、
    詰め込みすぎず主要な語に絞る。確実な補正は後段の表記統一で行う。
    """
    terms = []
    for value in dictionary.values():
        if isinstance(value, str) and value and value not in terms:
            terms.append(value)
    for t in (extra_terms or []):
        if t and t not in terms:
            terms.append(t)
    if not terms:
        return None
    return f"以下は日本語の会議音声です。次の固有名詞が登場します: {'、'.join(terms[:limit])}。"


def detect_repetition(text, min_len=8, threshold=6):
    """同一フレーズの連続を検出する（幻聴ループの早期警告）"""
    if not text:
        return None
    m = re.search(r"(.{%d,40}?)\1{%d,}" % (min_len, threshold), text)
    return m.group(1) if m else None


# ------------------------------------------------------------- whisper実行基盤
def resolve_mlx_whisper_module():
    """Python API が使えるなら最優先。全デコードオプションを確実に渡せる。"""
    global _MLX_WHISPER_MODULE
    if _MLX_WHISPER_MODULE is not None:
        return _MLX_WHISPER_MODULE or None
    try:
        import mlx_whisper  # noqa: PLC0415
        _MLX_WHISPER_MODULE = mlx_whisper
    except Exception:  # noqa: BLE001
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


def _decode_options(engine):
    opts = dict(DEFAULT_DECODE)
    opts.update(engine.get("decode") or {})
    return opts


def _whisper_via_module(audio_path, model, language, decode, initial_prompt):
    """Python API 経路。decode設定を確実に反映できる。"""
    mlx_whisper = resolve_mlx_whisper_module()
    if mlx_whisper is None:
        return None

    kwargs = {"language": language}
    kwargs["condition_on_previous_text"] = bool(decode.get("condition_on_previous_text", False))
    temps = decode.get("temperature")
    if isinstance(temps, (list, tuple)) and temps:
        kwargs["temperature"] = tuple(float(t) for t in temps)
    elif isinstance(temps, (int, float)):
        kwargs["temperature"] = float(temps)
    for key in ("compression_ratio_threshold", "logprob_threshold",
                "no_speech_threshold", "hallucination_silence_threshold"):
        if decode.get(key) is not None:
            kwargs[key] = float(decode[key])
    if initial_prompt:
        kwargs["initial_prompt"] = initial_prompt

    try:
        result = mlx_whisper.transcribe(str(audio_path), path_or_hf_repo=model, **kwargs)
    except Exception as e:  # noqa: BLE001
        log = write_transcribe_debug_log(audio_path, ["python:mlx_whisper.transcribe"],
                                         error=e, note="module_transcribe_failed")
        print(f"  ⚠️ Python API での文字起こしに失敗({type(e).__name__})。CLIを試します")
        print(f"  📝 デバッグログ: {log}")
        return None
    return (result.get("text") or "").strip() or None


def _run_whisper(audio_path, model, language, output_format, decode, initial_prompt):
    """mlx_whisper(CLI)を実行し 出力ファイルパス or None を返す"""
    audio_path = Path(audio_path)
    # 出力先を音声ごとの専用ディレクトリに分離する。
    # 旧実装は音声と同じディレクトリに出力し、見つからない場合に
    # glob("*.txt")[0] で「別の音声の結果」を拾う事故があり得た。
    out_dir = audio_path.parent / f"_stt_{audio_path.stem}"
    out_dir.mkdir(parents=True, exist_ok=True)
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
        "--condition-on-previous-text",
        "True" if decode.get("condition_on_previous_text") else "False",
    ]
    # mlx_whisper の CLI は --temperature に単一のfloatしか受け付けない
    # （温度フォールバックのタプルを渡せるのは Python API のみ）。
    # ここでは先頭値だけを渡し、代わりに幻聴抑止のしきい値で補う。
    temps = decode.get("temperature")
    if isinstance(temps, (list, tuple)) and temps:
        cmd += ["--temperature", str(float(temps[0]))]
    elif isinstance(temps, (int, float)):
        cmd += ["--temperature", str(float(temps))]
    if decode.get("hallucination_silence_threshold") is not None:
        cmd += ["--hallucination-silence-threshold",
                str(decode["hallucination_silence_threshold"])]
    if decode.get("compression_ratio_threshold") is not None:
        cmd += ["--compression-ratio-threshold", str(decode["compression_ratio_threshold"])]
    if decode.get("logprob_threshold") is not None:
        cmd += ["--logprob-threshold", str(decode["logprob_threshold"])]
    if decode.get("no_speech_threshold") is not None:
        cmd += ["--no-speech-threshold", str(decode["no_speech_threshold"])]
    if initial_prompt:
        cmd += ["--initial-prompt", initial_prompt]

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
    # 専用ディレクトリなので、ここにある出力は必ずこの音声由来
    candidates = list(out_dir.glob("*" + ext))
    if candidates:
        return candidates[0]

    print(f"  ⚠️ 文字起こし結果ファイル({ext})が見つかりません。")
    print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, cmd, result=result, note='output_not_found')}")
    return None


def whisper_text(audio_path, model, language="ja", decode=None, initial_prompt=None):
    decode = decode or dict(DEFAULT_DECODE)
    # PATH整備は Python API 経路より前に必ず行う。
    # mlx_whisper は内部で ffmpeg をサブプロセス起動するため、
    # PATHが /usr/bin:/bin のままだと FileNotFoundError: 'ffmpeg' になる。
    ensure_runtime_path()
    if not resolve_ffmpeg_cmd():
        print("  ❌ ffmpeg が見つかりません。（例: brew install ffmpeg）")
        print(f"  📝 デバッグログ: {write_transcribe_debug_log(audio_path, ['ffmpeg'], note='ffmpeg_not_found')}")
        return None

    text = _whisper_via_module(audio_path, model, language, decode, initial_prompt)
    if text:
        return text
    path = _run_whisper(audio_path, model, language, "txt", decode, initial_prompt)
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


def refine_plain_text(text, settings, verbose=True):
    """whisper の素のテキストにも フィラー除去 / 表記統一 を適用する。

    旧実装では qwen3-asr のときだけ llm_refine を通しており、
    whisper では normalize_dict.json が一度も適用されていなかった。
    """
    if not text:
        return text
    segments = [{"start": 0.0, "text": line}
                for line in text.splitlines() if line.strip()]
    if not segments:
        segments = [{"start": 0.0, "text": text}]
    refined = llm_refine.refine_segments(
        segments,
        fillers=settings.get("fillers", []),
        dictionary=load_dictionary(settings),
        verbose=verbose,
    )
    return "\n".join((s.get("text") or "").strip()
                     for s in refined if (s.get("text") or "").strip())


# ------------------------------------------------------------------- メインAPI
def transcribe(audio_path, label=None, settings=None, verbose=True):
    """labelに応じて文字起こしを行い、文字列を返す（失敗時はNone）"""
    settings = settings or load_settings()
    engine = get_engine(settings, label)
    etype = engine.get("type")
    if verbose:
        print(f"  エンジン: {engine.get('label')} (type={etype})")

    dictionary = load_dictionary(settings)

    # whisper：文字起こし後に フィラー除去・表記統一 を適用して返す
    if etype == "whisper":
        model = engine.get("model", DEFAULT_WHISPER_MODEL)
        decode = _decode_options(engine)
        initial_prompt = None
        if engine.get("initialPromptFromDict", True):
            initial_prompt = build_initial_prompt(dictionary, engine.get("contextTerms"))
            if initial_prompt and verbose:
                print(f"  ドメイン語彙: {initial_prompt[:60]}...")

        text = whisper_text(audio_path, model, engine.get("language", "ja"),
                            decode=decode, initial_prompt=initial_prompt)
        if not text:
            return None

        looped = detect_repetition(text)
        if looped and verbose:
            print(f"  ⚠️ 同一フレーズの連続を検出: 「{looped[:30]}」")
            print("     録音冒頭の無音や極端に小さい音量が原因のことが多いです")

        return refine_plain_text(text, settings, verbose=verbose) or None

    if etype != "qwen3-asr":
        raise ValueError(f"未対応のtype '{etype}'（label={engine.get('label')}）。"
                         "whisper か qwen3-asr を指定してください")

    # --- qwen3-asr ---------------------------------------------------------
    import qwen3_asr  # noqa: PLC0415  遅延importでMac以外の環境でも読み込める
    if not qwen3_asr.available():
        print("  ❌ mlx-qwen3-asr が未導入です（pip install \"mlx-qwen3-asr[aligner]\"）")
        return None

    context = qwen3_asr.build_context(dictionary, engine.get("contextTerms"))
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
        dictionary=dictionary,
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
