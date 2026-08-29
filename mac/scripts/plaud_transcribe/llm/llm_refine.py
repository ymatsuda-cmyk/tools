#!/usr/bin/env python3
"""文字起こしのルールベース整形

qwen3-asr エンジンの後処理として、フィラー除去・表記統一（表記揺れ辞書の適用）を行う。
ローカルLLMは使用しない（qwen3-asr / whisper の2エンジン構成）。
"""
import re


def rule_clean(text, fillers, dictionary):
    """必ず効く決定論的なクリーニング"""
    out = text.strip()
    for f in sorted(fillers or [], key=len, reverse=True):
        out = out.replace(f + "、", "").replace(f, "")
    for src, dst in sorted((dictionary or {}).items(), key=lambda kv: len(kv[0]), reverse=True):
        out = re.sub(re.escape(src), dst, out, flags=re.IGNORECASE)
    out = re.sub(r"[、,]{2,}", "、", out)
    out = re.sub(r"\s{2,}", " ", out).strip("、 　")
    return out.strip()


def refine_segments(segments, fillers=None, dictionary=None, verbose=True):
    """セグメントの text をルールベースで整形して返す（speaker/start/end は変更しない）"""
    cleaned = []
    empty_after_clean = 0
    for seg in segments:
        s = dict(seg)
        cleaned_text = rule_clean(s.get("text", ""), fillers, dictionary)
        if s.get("text") and not cleaned_text:
            empty_after_clean += 1
        s["text"] = cleaned_text
        cleaned.append(s)
    if verbose and empty_after_clean:
        print(f"  ℹ️ フィラーのみの発話を{empty_after_clean}件除去しました")
    return cleaned
