#!/usr/bin/env python3
"""notion_import.py — Notion 議事録DB（PLAUD Note との受け口）から取り込み、暗号化して GitHub に入れる

Notion 議事録DBは、PLAUD Note からの登録口として今までどおり使う（plaud-notion-sync / export・文字起こし）。
Mac がここから取り込み、新しいアプリ用に暗号化して GitHub private に格納する。

  取り込む条件  状態が「文字起こし」か「要約」で、取込ID が空のページ
    文字起こし → 文字起こしだけ取り込む（アプリでは「要約待ち」）
    要約       → 文字起こし＋要約＋ToDoの完了・経過＋メモも取り込む（アプリでは「要約済み」）
  閲覧グループ  「権限」の値（jba など）をそのまま role として使う。管理者は常に閲覧できる
  取り込んだあと 取込ID を書き、状態を「取込済み」にする。
               NOTION_CLEAR_AFTER_IMPORT=1 なら、本文（文字起こし）と要約・メモの欄を消す
               （ミーティング名・日時・会議時間・権限・カテゴリーなど管理用の項目は残す）

  使い方
    python3 notion_import.py --dry-run          何件取り込めるかだけ見る（何も書かない）
    python3 notion_import.py --limit 20         20件だけ取り込む（過去分の移行を少しずつ試すとき）
    python3 notion_import.py                    全部取り込む（minutes_worker.py からも毎回呼ばれる）
"""
import argparse
import datetime as dt
import json
import re
import time
from zoneinfo import ZoneInfo

import requests

from m2common import ENV, GitHub, log, load_keys
from m2crypto import new_dek, seal_json, wrap_dek

JST = ZoneInfo("Asia/Tokyo")
ADMIN_ROLE = "xYz"

# 議事録DBのプロパティ名（今の minutes の GAS と同じ）
P_TITLE, P_STATUS, P_DATE, P_DUR = "ミーティング名", "状態", "日時", "会議時間"
P_PERM, P_TAGS = "権限", "カテゴリー"
P_SUMMARY, P_DECISIONS, P_TODOS, P_TODO_LOGS, P_TOPICS = "要約", "決定事項", "ToDo", "ToDo経過", "論点"
P_AGENDA, P_MODEL, P_GENERATED, P_MEMO, P_MINDMAP = "議事", "要約モデル", "要約日時", "メモ", "マインドマップ"
P_IMPORT_ID = "取込ID"
ST_TRANSCRIBED, ST_SUMMARIZED, ST_IMPORTED = "文字起こし", "要約", "取込済み"
CLEAR_PROPS = [P_SUMMARY, P_DECISIONS, P_TODOS, P_TODO_LOGS, P_TOPICS, P_AGENDA, P_MEMO, P_MINDMAP]


class NotionDB:
    def __init__(self, token, db):
        self.db = db
        self.s = requests.Session()
        self.s.headers.update({"Authorization": "Bearer " + token, "Notion-Version": "2022-06-28", "Content-Type": "application/json"})

    def req(self, method, path, **kw):
        for i in range(6):
            r = self.s.request(method, "https://api.notion.com/v1/" + path, timeout=60, **kw)
            if r.status_code in (429, 502, 503):
                time.sleep(1 + i * 2)
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"Notion {r.status_code}: {r.text[:300]}")
            time.sleep(0.34)  # 1秒3回まで
            return r.json()
        raise RuntimeError("Notion が混み合っています")

    def ensure_import_prop(self):
        db = self.req("GET", f"databases/{self.db}")
        if P_IMPORT_ID not in db["properties"]:
            self.req("PATCH", f"databases/{self.db}", json={"properties": {P_IMPORT_ID: {"rich_text": {}}}})
            log(f"議事録DBに「{P_IMPORT_ID}」を追加しました")

    def pending(self):
        body = {"page_size": 100, "filter": {"and": [
            {"or": [{"property": P_STATUS, "select": {"equals": ST_TRANSCRIBED}},
                    {"property": P_STATUS, "select": {"equals": ST_SUMMARIZED}}]},
            {"property": P_IMPORT_ID, "rich_text": {"is_empty": True}},
        ]}, "sorts": [{"property": P_DATE, "direction": "ascending"}]}
        cursor = None
        while True:
            if cursor:
                body["start_cursor"] = cursor
            j = self.req("POST", f"databases/{self.db}/query", json=body)
            for p in j["results"]:
                if not p.get("archived") and not p.get("in_trash"):
                    yield p
            if not j.get("has_more"):
                return
            cursor = j["next_cursor"]

    def blocks(self, page_id):
        out, cursor = [], None
        while True:
            q = f"blocks/{page_id}/children?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
            j = self.req("GET", q)
            out += j["results"]
            if not j.get("has_more"):
                return out
            cursor = j["next_cursor"]


# ---------- Notion の値を読む ----------

def plain(rt):
    return "".join(x.get("plain_text", "") for x in (rt or []))


def prop_text(props, name):
    p = props.get(name) or {}
    t = p.get("type")
    if t in ("rich_text", "title"):
        return plain(p.get(t)).strip()
    return ""


def prop_multi(props, name):
    p = props.get(name) or {}
    return [o["name"] for o in (p.get("multi_select") or [])]


def prop_select(props, name):
    p = props.get(name) or {}
    return (p.get("select") or {}).get("name", "")


def prop_date(props, name):
    p = props.get(name) or {}
    return (p.get("date") or {}).get("start") or ""


def to_local(iso):
    """Notion の日時 → アプリの形（JST の 'YYYY-MM-DDTHH:MM'）"""
    if not iso:
        return ""
    if len(iso) == 10:
        return iso + "T00:00"
    d = dt.datetime.fromisoformat(iso.replace("Z", "+00:00"))
    if d.tzinfo is None:
        d = d.replace(tzinfo=JST)
    return d.astimezone(JST).strftime("%Y-%m-%dT%H:%M")


def parse_dur(text):
    """「1時間 12分 36秒」→ 秒"""
    sec = 0
    for n, unit in re.findall(r"(\d+)\s*(時間|分|秒)", text or ""):
        sec += int(n) * {"時間": 3600, "分": 60, "秒": 1}[unit]
    return sec


def parse_t(s):
    m = re.match(r"^(?:(\d+):)?(\d{1,2}):(\d{2})$", s.strip())
    if not m:
        return None
    return int(m.group(1) or 0) * 3600 + int(m.group(2)) * 60 + int(m.group(3))


def segments_from_text(text):
    """行頭の [00:01:12] / 00:01:12 / 1:12 を時刻として読む（web/src/summary.js と同じ）"""
    segs = []
    for line in text.replace("\r\n", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^\[?((?:\d+:)?\d{1,2}:\d{2})\]?\s*(.*)$", line)
        if m and m.group(2):
            segs.append({"t": parse_t(m.group(1)), "s": m.group(2)})
        else:
            segs.append({"t": None, "s": line})
    return segs


def lines(text):
    return [x for x in (text or "").split("\n") if x.strip()]


def old_summary(props):
    """今の minutes の要約欄 → アプリの形。ToDo の完了は notes 側に分ける"""
    agenda = []
    raw = prop_text(props, P_AGENDA)
    if raw:
        try:
            for a in json.loads(raw) or []:
                agenda.append({
                    "topic": str(a.get("topic", "")).strip() or "（議題）", "t": str(a.get("t", "")),
                    "outcome": str(a.get("outcome", "")).strip(),
                    "points": [{"text": p, "t": ""} if isinstance(p, str) else {"text": str(p.get("text", "")), "t": str(p.get("t", ""))}
                               for p in a.get("points") or [] if p],
                })
        except ValueError:
            log("議事の JSON を読めませんでした（途中で切れている可能性）")
    todos, done = [], {}
    for i, line in enumerate(lines(prop_text(props, P_TODOS))):
        m = re.match(r"^\[([ xX])\]\s?(.*)$", line)
        text = m.group(2) if m else line
        if m and m.group(1).lower() == "x":
            done[str(len(todos))] = True
        todos.append({"text": text.strip(), "owner": "", "due": ""})
    card = prop_text(props, P_SUMMARY)
    gen = prop_date(props, P_GENERATED)
    s = {
        "cardSummary": card, "summary": card, "agenda": agenda,
        "decisions": [{"text": d, "t": ""} for d in lines(prop_text(props, P_DECISIONS))],
        "todos": todos, "topics": lines(prop_text(props, P_TOPICS)),
        "model": prop_text(props, P_MODEL) or "（今の minutes）",
        "at": int(dt.datetime.fromisoformat(gen.replace("Z", "+00:00")).timestamp() * 1000) if len(gen) > 10 else int(time.time() * 1000),
    }
    mind = prop_text(props, P_MINDMAP)
    if mind:
        s["mindmap"] = mind  # 第2段階で表示する
    notes = {"done": done}
    memo = prop_text(props, P_MEMO)
    if memo:
        notes["memo"] = memo
    logs = prop_text(props, P_TODO_LOGS)
    if logs:
        notes["todoLogs"] = logs  # 今の minutes の ToDo経過（JSON 文字列のまま保管）
    return s, notes


# ---------- 取り込み ----------

def new_id():
    import secrets
    a = "abcdefghijklmnopqrstuvwxyz0123456789"
    return "m_" + "".join(secrets.choice(a) for _ in range(12))


def import_pages(dry_run=False, limit=None, clear=None) -> int:
    token, db = ENV.get("NOTION_TOKEN"), ENV.get("NOTION_MINUTES_DB_ID")
    if not token or not db:
        return 0
    clear = ENV.get("NOTION_CLEAR_AFTER_IMPORT", "1") == "1" if clear is None else clear
    keys = load_keys()
    roles = keys["roles"]
    admin_gid = roles[ADMIN_ROLE]["gid"]
    n = NotionDB(token, db)
    if not dry_run:
        n.ensure_import_prop()
    gh = None if dry_run else GitHub()
    count = 0
    for page in n.pending():
        if limit and count >= limit:
            break
        props = page["properties"]
        title = prop_text(props, P_TITLE) or "（無題）"
        status = prop_select(props, P_STATUS)
        try:
            body_blocks = n.blocks(page["id"])
            text = "\n".join(t for t in (plain((b.get(b["type"]) or {}).get("rich_text")) for b in body_blocks) if t.strip())
            segs = segments_from_text(text)
            has_sum = status == ST_SUMMARIZED and (prop_date(props, P_GENERATED) or prop_text(props, P_SUMMARY))
            perms = prop_multi(props, P_PERM)
            unknown = [r for r in perms if r not in roles]
            gids = sorted({admin_gid, *[roles[r]["gid"] for r in perms if r in roles]})
            if dry_run:
                log(f"[確認] {title}（{to_local(prop_date(props, P_DATE))}）{'要約済み' if has_sum else '要約待ち'} ・ 文字起こし {len(segs)} 行 ・ 閲覧 {', '.join(perms) or '管理者のみ'}"
                    + (f" ・ 対応する role が無い権限: {', '.join(unknown)}" if unknown else ""))
                count += 1
                continue
            if not segs and not has_sum:
                log(f"{title}: 文字起こしが空なので見送ります")
                continue
            if unknown:
                log(f"{title}: 対応する role が無い権限は外しました（{', '.join(unknown)}）")

            eid = new_id()
            dek = new_dek()
            k = {g: wrap_dek(dek, next(r["pub"] for r in roles.values() if r["gid"] == g)) for g in gids}
            meta = {"title": title, "tags": prop_multi(props, P_TAGS), "notion": page["id"]}
            entry = {
                "id": eid, "date": to_local(prop_date(props, P_DATE)), "dur": parse_dur(prop_text(props, P_DUR)) or
                (max([x["t"] for x in segs if x["t"] is not None], default=0)),
                "st": "ok" if has_sum else "sw", "src": "plaud", "by": admin_gid,
                "upd": int(time.time() * 1000), "k": k, "m": seal_json(dek, meta),
            }
            gh.write(f"m/{eid}/t.json", json.dumps(seal_json(dek, {"segments": segs, "asr": "Notion から取り込み"})), f"t {eid}")
            if has_sum:
                s, notes = old_summary(props)
                gh.write(f"m/{eid}/s.json", json.dumps(seal_json(dek, s)), f"s {eid}")
                if notes.get("done") or notes.get("memo") or notes.get("todoLogs"):
                    gh.write(f"m/{eid}/n.json", json.dumps(seal_json(dek, notes)), f"n {eid}")
            gh.update_index(lambda lst: lst.append(entry))

            # Notion 側：取込ID と状態を書き、必要なら中身を消す
            upd = {P_IMPORT_ID: {"rich_text": [{"text": {"content": eid}}]}, P_STATUS: {"select": {"name": ST_IMPORTED}}}
            if clear:
                for name in CLEAR_PROPS:
                    if name in props and props[name].get("type") == "rich_text":
                        upd[name] = {"rich_text": []}
            n.req("PATCH", f"pages/{page['id']}", json={"properties": upd})
            if clear:
                for b in body_blocks:
                    n.req("DELETE", f"blocks/{b['id']}")
            log(f"{title}: 取り込みました（{eid}・{'要約済み' if has_sum else '要約待ち'}{'・Notion の本文を消去' if clear else ''}）")
            count += 1
        except Exception as ex:  # noqa: BLE001  1件の失敗で全体を止めない
            log(f"{title}: 取り込めませんでした {ex}")
    return count


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dry-run", action="store_true", help="何も書かずに、取り込める件数だけ見る")
    p.add_argument("--limit", type=int, default=None, help="取り込む件数の上限")
    p.add_argument("--keep", action="store_true", help="取り込んでも Notion の本文・要約を消さない")
    a = p.parse_args()
    if not ENV.get("NOTION_TOKEN") or not ENV.get("NOTION_MINUTES_DB_ID"):
        raise SystemExit(".env に NOTION_TOKEN と NOTION_MINUTES_DB_ID を書いてください")
    n = import_pages(dry_run=a.dry_run, limit=a.limit, clear=False if a.keep else None)
    print(f"{'確認' if a.dry_run else '取り込み'}：{n} 件")


if __name__ == "__main__":
    main()
