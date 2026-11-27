"""m2legacy.py — 今の minutes（Notion）の形式 → MinutesStock の形式

  マーカー   文字列に <m1>…</m1>（色1〜3）が埋め込まれている → 文章は素の文字、マーカーは n.json の marks に位置で持つ
  ToDo経過   [{text, done, logs:[{date, text}]}] の JSON → n.json の logs（ToDo の key ごと）
  ToDo       並び順 → id（'i0', 'i1' …）を付ける。完了は n.json の done
位置は JavaScript の文字列と同じ数え方（UTF-16）にそろえる（絵文字などが前にあっても、ずれないように）。
"""
import datetime as dt
import json
import re
import secrets
import time

TAG_RE = re.compile(r"<m([123])>([\s\S]*?)</m\1>")


def u16(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def new_mark_id() -> str:
    return "k" + "".join(secrets.choice("abcdefghijklmnopqrstuvwxyz0123456789") for _ in range(8))


def strip_marks(raw: str):
    """'<m1>abc</m1>def' → ('abcdef', [(開始, 終了, 色)])。開始・終了は UTF-16 の位置"""
    raw = raw or ""
    out, spans, pos, plain_len = [], [], 0, 0
    for m in TAG_RE.finditer(raw):
        gap = raw[pos:m.start()]
        out.append(gap)
        plain_len += u16(gap)
        seg = m.group(2)
        spans.append((plain_len, plain_len + u16(seg), int(m.group(1)), seg))
        out.append(seg)
        plain_len += u16(seg)
        pos = m.end()
    out.append(raw[pos:])
    return "".join(out), spans


def collect(loc: str, raw: str, marks: list, at_ms: int) -> str:
    """raw のタグを外した文字を返し、マーカーを marks に足す"""
    plain, spans = strip_marks(raw)
    for s, e, c, q in spans:
        if q.strip():
            marks.append({"id": new_mark_id(), "loc": loc, "s": s, "e": e, "c": c, "q": q, "at": at_ms})
    return plain


def has_tags(raw: str) -> bool:
    return bool(TAG_RE.search(raw or ""))


def parse_logs(logs_json):
    """ToDo経過の JSON → [{text(タグつき), done, logs:[{d, x}]}]"""
    if not logs_json:
        return []
    try:
        arr = json.loads(logs_json) if isinstance(logs_json, str) else logs_json
    except ValueError:
        return []
    out = []
    for e in arr if isinstance(arr, list) else []:
        logs = []
        for l in e.get("logs") or []:
            d = str(l.get("date") or "")
            try:
                ms = int(dt.datetime.fromisoformat(d.replace("Z", "+00:00")).timestamp() * 1000)
            except ValueError:
                ms = int(time.time() * 1000)
            x = str(l.get("text") or "").strip()
            if x:
                logs.append({"d": ms, "x": x})
        out.append({"text": str(e.get("text") or ""), "done": bool(e.get("done")), "logs": logs})
    return out


def build_new(raw: dict, logs_json, memo: str, model: str, at_ms: int, done_by_index=None):
    """raw = {card, decisions:[str], todos:[(str, done)], topics:[str], agenda:[{topic, points:[str|{text,t}], outcome}]}
    → (s, n)。n は {done, logs, marks, memo?}"""
    marks = []
    card = collect("summary", raw.get("card", ""), marks, at_ms)
    s = {
        "cardSummary": card, "summary": card,
        "decisions": [{"text": collect(f"decision:{i}", t, marks, at_ms), "t": ""} for i, t in enumerate(raw.get("decisions") or [])],
        "agenda": [], "topics": [], "todos": [],
        "model": model or "（今の minutes）", "at": at_ms,
    }
    for i, a in enumerate(raw.get("agenda") or []):
        pts = []
        for j, p in enumerate(a.get("points") or []):
            text, t = (p.get("text", ""), str(p.get("t", ""))) if isinstance(p, dict) else (str(p), "")
            pts.append({"text": collect(f"agenda:{i}:point:{j}", text, marks, at_ms), "t": t})
        outcome = str(a.get("outcome") or "")
        s["agenda"].append({
            "topic": collect(f"agenda:{i}:topic", str(a.get("topic") or ""), marks, at_ms) or "（議題）",
            "t": str(a.get("t") or ""),
            "outcome": collect(f"agenda:{i}:outcome", outcome, marks, at_ms) if outcome else "",
            "points": pts,
        })
    for i, t in enumerate(raw.get("topics") or []):
        s["topics"].append(collect(f"topic:{i}", t, marks, at_ms))

    # ToDo：並び順の id、完了、経過（本文が同じものに付ける。付けられなかった経過は、ToDo ごと残す）
    done, logs = {}, {}
    entries = parse_logs(logs_json)
    used = set()
    for i, (text_raw, is_done) in enumerate(raw.get("todos") or []):
        key = f"i{i}"
        text = collect(f"todo:{key}", text_raw, marks, at_ms)
        s["todos"].append({"id": key, "text": text, "owner": "", "due": ""})
        if is_done:
            done[key] = True
        for k, e in enumerate(entries):
            if k not in used and e["text"] == text_raw and e["logs"]:
                logs[key] = e["logs"]
                used.add(k)
                break
    for k, e in enumerate(entries):
        if k in used or not e["logs"]:
            continue
        key = f"o{k}"
        s["todos"].append({"id": key, "text": strip_marks(e["text"])[0], "owner": "", "due": "", "orphan": True})
        logs[key] = e["logs"]
        if e["done"]:
            done[key] = True
    n = {"done": done, "logs": logs, "marks": marks}
    if memo:
        n["memo"] = memo
    return s, n


def needs_upgrade(s: dict, n: dict) -> bool:
    """前の版の取り込み（タグが文字に混ざったまま／経過が todoLogs のまま／ToDo に id が無い）か"""
    if (n or {}).get("todoLogs"):
        return True
    texts = [s.get("cardSummary", ""), s.get("summary", "")]
    texts += [d.get("text", "") for d in s.get("decisions", [])] + list(s.get("topics", [])) + [t.get("text", "") for t in s.get("todos", [])]
    for a in s.get("agenda", []):
        texts += [a.get("topic", ""), a.get("outcome", "")] + [p.get("text", "") for p in a.get("points", [])]
    if any(has_tags(t) for t in texts):
        return True
    return any(not t.get("id") for t in s.get("todos", [])) and bool((n or {}).get("done"))


def upgrade_pair(s: dict, n: dict):
    """前の版の取り込み結果を、新しい形に直す → (s, n)"""
    n = dict(n or {})
    raw = {
        "card": s.get("summary") or s.get("cardSummary") or "",
        "decisions": [d.get("text", "") for d in s.get("decisions", [])],
        "todos": [(t.get("text", ""), False) for t in s.get("todos", [])],
        "topics": list(s.get("topics", [])),
        "agenda": [{"topic": a.get("topic", ""), "outcome": a.get("outcome", ""), "points": [dict(p) for p in a.get("points", [])]} for a in s.get("agenda", [])],
    }
    # 前の版は、完了を「並び順の番号」で持っていた
    old_done = {str(k): v for k, v in (n.get("done") or {}).items()}
    raw["todos"] = [(text, bool(old_done.get(str(i)))) for i, (text, _) in enumerate(raw["todos"])]
    new_s, new_n = build_new(raw, n.get("todoLogs"), n.get("memo", ""), s.get("model", ""), s.get("at") or int(time.time() * 1000))
    new_s["cardSummary"] = strip_marks(s.get("cardSummary", ""))[0] or new_s["cardSummary"]
    for k in ("edited",):
        if k in s:
            new_s[k] = s[k]
    # 新しい形で付いていた完了・経過・マーカーは残す
    for key, v in (n.get("done") or {}).items():
        if not str(key).isdigit() and key not in new_n["done"]:
            new_n["done"][key] = v
    for key, v in (n.get("logs") or {}).items():
        new_n["logs"].setdefault(key, v)
    new_n["marks"] = (n.get("marks") or []) + new_n["marks"]
    return new_s, new_n
