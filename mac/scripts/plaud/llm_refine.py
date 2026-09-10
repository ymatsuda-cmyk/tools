#!/usr/bin/env python3
"""文字起こしのルールベース整形

whisper / qwen3-asr 双方の後処理として、フィラー除去・表記統一（表記揺れ辞書の適用）
を行う。ローカルLLMは使用しない。

設計上の要点:
  - 辞書は「1回のスキャンで一括置換」する。ルールを順番に適用すると、
    ある置換の結果が別のルールの入力になって連鎖する事故が起きるため。
  - 置換文字列は正規表現として解釈しない（"C:\\group" 等で例外になるのを防ぐ）。
  - ASCII の見出し語は単語境界付きで照合する。境界なしだと
    "applaud" → "apPLAUD"、"notion.so" → "Notion.so" のような誤置換が起きる。
"""
import re

# 半角英数のみで構成された見出し語は単語境界を付けて照合する
_ASCII_KEY = re.compile(r"^[0-9A-Za-z._-]+$")
# 日本語の見出し語の直前・直後に英字が来た場合は照合しない、といった配慮は不要
_URL_RE = re.compile(r"https?://\S+|[\w.+-]+@[\w-]+\.[\w.]+")


def _compile_dictionary(dictionary):
    """辞書を1本の正規表現にまとめる（長い見出し語を優先してマッチさせる）"""
    items = [(str(k), str(v)) for k, v in (dictionary or {}).items() if k]
    if not items:
        return None, {}
    items.sort(key=lambda kv: len(kv[0]), reverse=True)

    parts, mapping = [], {}
    for i, (src, dst) in enumerate(items):
        name = f"t{i}"
        mapping[name] = dst
        body = re.escape(src)
        if _ASCII_KEY.match(src):
            # 英数の見出し語は前後が英数でないときだけ一致させる
            body = rf"(?<![0-9A-Za-z]){body}(?![0-9A-Za-z])"
        parts.append(f"(?P<{name}>{body})")
    # ASCII見出し語の大文字小文字を吸収する。日本語には影響しない
    return re.compile("|".join(parts), re.IGNORECASE), mapping


def _compile_fillers(fillers):
    """フィラーを1本の正規表現にまとめる。直後の読点も一緒に落とす"""
    items = sorted({f for f in (fillers or []) if f}, key=len, reverse=True)
    if not items:
        return None
    body = "|".join(re.escape(f) for f in items)
    return re.compile(rf"(?:{body})[、,]?")


def _protect_urls(text):
    """URL・メールアドレスを退避する（辞書置換で壊さないため）"""
    stash = []

    def _hold(m):
        stash.append(m.group(0))
        return f"\x00{len(stash) - 1}\x00"

    return _URL_RE.sub(_hold, text), stash


def _restore_urls(text, stash):
    for i, original in enumerate(stash):
        text = text.replace(f"\x00{i}\x00", original)
    return text


def rule_clean(text, fillers, dictionary, *, protect_urls=True,
               _filler_re=None, _dict_re=None, _dict_map=None):
    """必ず効く決定論的なクリーニング

    _filler_re / _dict_re / _dict_map は refine_segments からの内部最適化用。
    単体で呼ぶ場合は fillers / dictionary だけ渡せばよい。
    """
    out = (text or "").strip()
    if not out:
        return ""

    stash = []
    if protect_urls:
        out, stash = _protect_urls(out)

    filler_re = _filler_re if _filler_re is not None else _compile_fillers(fillers)
    if filler_re:
        out = filler_re.sub("", out)

    if _dict_re is not None:
        dict_re, dict_map = _dict_re, (_dict_map or {})
    else:
        dict_re, dict_map = _compile_dictionary(dictionary)
    if dict_re:
        # 置換値を正規表現として解釈させないため関数で返す
        out = dict_re.sub(lambda m: dict_map.get(m.lastgroup, m.group(0)), out)

    if protect_urls:
        out = _restore_urls(out, stash)

    out = re.sub(r"[、,]{2,}", "、", out)
    out = re.sub(r"^[、,\s　]+", "", out)
    out = re.sub(r"\s{2,}", " ", out)
    return out.strip("、 　\t\n")


def refine_segments(segments, fillers=None, dictionary=None, verbose=True):
    """セグメントの text をルールベースで整形して返す（speaker/start/end は変更しない）"""
    # 正規表現は全セグメントで使い回す（長い議事録での再コンパイルを避ける）
    filler_re = _compile_fillers(fillers)
    dict_re, dict_map = _compile_dictionary(dictionary)

    cleaned = []
    empty_after_clean = 0
    replaced = 0

    for seg in segments:
        s = dict(seg)
        original = s.get("text", "")
        cleaned_text = rule_clean(
            original, fillers, dictionary,
            _filler_re=filler_re, _dict_re=dict_re, _dict_map=dict_map,
        )
        if original and not cleaned_text:
            empty_after_clean += 1
        if dict_re and original:
            replaced += len(dict_re.findall(original))
        s["text"] = cleaned_text
        cleaned.append(s)

    if verbose:
        if replaced:
            print(f"  ✏️ 表記統一を{replaced}箇所に適用しました")
        if empty_after_clean:
            print(f"  ℹ️ フィラーのみの発話を{empty_after_clean}件除去しました")
    return cleaned


if __name__ == "__main__":
    import json
    import sys
    from pathlib import Path

    dic_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("normalize_dict.json")
    dic = json.loads(dic_path.read_text(encoding="utf-8")) if dic_path.exists() else {}
    fillers = ["えー", "えーと", "ええと", "えっと", "あー", "あのー", "そのー",
               "まあその", "なんていうか", "はいはい", "うーん"]
    for line in sys.stdin:
        print(rule_clean(line.rstrip("\n"), fillers, dic))
