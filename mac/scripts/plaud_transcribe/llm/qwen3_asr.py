#!/usr/bin/env python3
"""Qwen3-ASR (mlx-qwen3-asr) を使った文字起こし

Apple Silicon 上でローカル実行する。返り値は
    [{"start": float, "end": float, "text": str, "speaker": str|None}, ...]
という共通の形なので、transcribe_engines.py 側の声紋照合・整形処理はそのまま使える。

前提:
    pip install "mlx-qwen3-asr[aligner]"        # 日本語のタイムスタンプ整合用
    pip install "mlx-qwen3-asr[diarize]"        # 話者分離を使う場合
    export PYANNOTE_AUTH_TOKEN=hf_...           # 話者分離を使う場合
"""
import importlib.util
import re

_SESSIONS = {}


def available() -> bool:
    return importlib.util.find_spec("mlx_qwen3_asr") is not None


def diarize_available() -> bool:
    return importlib.util.find_spec("pyannote") is not None


def get_session(model):
    """Sessionはモデルを保持し続けるので、複数ファイル処理で再ロードが起きない"""
    if model not in _SESSIONS:
        from mlx_qwen3_asr import Session
        _SESSIONS[model] = Session(model=model)
    return _SESSIONS[model]


def build_context(dictionary, extra_terms=None, max_chars=400):
    """表記統一辞書の「正しい表記」をドメイン語彙としてASRに与える

    Qwen3-ASR の context は空白区切りの用語列で、認識段階で
    固有名詞・専門用語の誤変換を減らせる（＝表記揺れの発生源を先に潰す）。
    """
    terms = []
    for v in (dictionary or {}).values():
        if v and v not in terms:
            terms.append(v)
    for t in (extra_terms or []):
        if t and t not in terms:
            terms.append(t)
    ctx = ""
    for t in terms:
        cand = (ctx + " " + t).strip()
        if len(cand) > max_chars:
            break
        ctx = cand
    return ctx


def _group_words(words, gap_sec=0.7, max_chars=60):
    """単語レベルのタイムスタンプを発話単位にまとめる（話者分離を使わない場合）"""
    segments, cur = [], None
    for w in words:
        text = (w.get("text") or "").strip()
        if not text:
            continue
        start, end = float(w.get("start", 0.0)), float(w.get("end", 0.0))
        if cur is None:
            cur = {"start": start, "end": end, "text": text}
            continue
        gap = start - cur["end"]
        if gap > gap_sec or len(cur["text"]) >= max_chars or re.search(r"[。！？!?]$", cur["text"]):
            segments.append(cur)
            cur = {"start": start, "end": end, "text": text}
        else:
            cur["text"] += text
            cur["end"] = end
    if cur:
        segments.append(cur)
    return segments


def transcribe_segments(audio_path, cfg, context="", verbose=True):
    """(segments, diarized) を返す。diarized=True なら speaker に SPEAKER_00 等が入る"""
    model = cfg.get("model", "Qwen/Qwen3-ASR-0.6B")
    session = get_session(model)

    kwargs = {
        "language": cfg.get("language", "Japanese"),
        "return_timestamps": True,
    }
    if context:
        kwargs["context"] = context

    want_diarize = bool(cfg.get("diarize", True))
    if want_diarize and not diarize_available():
        if verbose:
            print("  ⚠️ pyannoteが未導入のため話者分離をスキップします"
                  "（pip install \"mlx-qwen3-asr[diarize]\"）")
        want_diarize = False
    if want_diarize:
        kwargs["diarize"] = True
        if cfg.get("numSpeakers"):
            kwargs["diarization_num_speakers"] = int(cfg["numSpeakers"])
        else:
            if cfg.get("minSpeakers"):
                kwargs["diarization_min_speakers"] = int(cfg["minSpeakers"])
            if cfg.get("maxSpeakers"):
                kwargs["diarization_max_speakers"] = int(cfg["maxSpeakers"])

    result = None
    try:
        result = session.transcribe(str(audio_path), **kwargs)
    except Exception as e:  # noqa: BLE001
        if want_diarize:
            if verbose:
                print(f"  ⚠️ 話者分離付きの実行に失敗: {type(e).__name__}: {e}")
                print("     話者分離なしで再試行します")
            for k in ("diarize", "diarization_num_speakers",
                      "diarization_min_speakers", "diarization_max_speakers"):
                kwargs.pop(k, None)
            want_diarize = False
            try:
                result = session.transcribe(str(audio_path), **kwargs)
            except Exception as e2:  # noqa: BLE001
                print(f"  ❌ Qwen3-ASRの実行に失敗: {type(e2).__name__}: {e2}")
                return None, False
        else:
            print(f"  ❌ Qwen3-ASRの実行に失敗: {type(e).__name__}: {e}")
            return None, False

    if getattr(result, "truncated", False) and verbose:
        print("  ⚠️ デコードが打ち切られたチャンクがあります（truncated=True）")

    spk = getattr(result, "speaker_segments", None)
    if want_diarize and spk:
        segments = [{"start": float(s.get("start", 0.0)),
                     "end": float(s.get("end", 0.0)),
                     "text": (s.get("text") or "").strip(),
                     "speaker": s.get("speaker")}
                    for s in spk if (s.get("text") or "").strip()]
        if segments:
            return segments, True

    words = getattr(result, "segments", None)
    if words:
        return _group_words(words,
                            gap_sec=cfg.get("gapSec", 0.7),
                            max_chars=cfg.get("maxChars", 60)), False

    text = (getattr(result, "text", "") or "").strip()
    if text:
        return [{"start": 0.0, "end": 0.0, "text": text}], False
    return None, False
