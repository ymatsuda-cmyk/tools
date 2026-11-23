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
    python3 notion_import.py --keep             取り込むが、Notion の本文・要約は残す（移行の確認用）
    python3 notion_import.py --clear-imported   --keep で残した取り込み済みページの本文・要約を、あとから消す
"""
import argparse
import datetime as dt
import fcntl
import sys
import json
import re
import time
from zoneinfo import ZoneInfo

import requests

from m2common import ENV, GitHub, log, load_keys
from m2crypto import jwk_to_priv, new_dek, open_json, seal_json, unwrap_dek, wrap_dek
from m2legacy import build_new, needs_upgrade, upgrade_pair

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

    _last = 0.0

    def req(self, method, path, **kw):
        for i in range(8):
            # Notion の上限（平均1秒3回）に合わせて、送り始めの間隔を 0.34 秒あける
            wait = 0.34 - (time.monotonic() - NotionDB._last)
            if wait > 0:
                time.sleep(wait)
            NotionDB._last = time.monotonic()
            r = self.s.request(method, "https://api.notion.com/v1/" + path, timeout=60, **kw)
            if r.status_code in (429, 502, 503, 504):
                time.sleep(float(r.headers.get("Retry-After") or (1 + i * 2)))
                continue
            if method == "DELETE" and r.status_code in (400, 404) and "archived" in r.text:
                return {}  # もう消えているブロック（前回の途中で消したもの）
            if r.status_code >= 400:
                raise RuntimeError(f"Notion {r.status_code}: {r.text[:300]}")
            return r.json()
        raise RuntimeError("Notion が混み合っています")

    def has_import_prop(self) -> bool:
        return P_IMPORT_ID in self.req("GET", f"databases/{self.db}")["properties"]

    def ensure_import_prop(self):
        if not self.has_import_prop():
            self.req("PATCH", f"databases/{self.db}", json={"properties": {P_IMPORT_ID: {"rich_text": {}}}})
            log(f"議事録DBに「{P_IMPORT_ID}」を追加しました")

    def pending(self, has_prop=True):
        conds = [{"or": [{"property": P_STATUS, "select": {"equals": ST_TRANSCRIBED}},
                         {"property": P_STATUS, "select": {"equals": ST_SUMMARIZED}}]}]
        if has_prop:  # 「取込ID」の列がまだ無いとき（初回の --dry-run）は、全部が未取り込み
            conds.append({"property": P_IMPORT_ID, "rich_text": {"is_empty": True}})
        body = {"page_size": 100, "filter": {"and": conds}, "sorts": [{"property": P_DATE, "direction": "ascending"}]}
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

    def imported(self):
        body = {"page_size": 100, "filter": {"and": [
            {"property": P_STATUS, "select": {"equals": ST_IMPORTED}},
            {"property": P_IMPORT_ID, "rich_text": {"is_not_empty": True}},
        ]}}
        cursor = None
        while True:
            if cursor:
                body["start_cursor"] = cursor
            j = self.req("POST", f"databases/{self.db}/query", json=body)
            yield from j["results"]
            if not j.get("has_more"):
                return
            cursor = j["next_cursor"]

    def clear_page(self, page, body_blocks=None, progress=None) -> bool:
        """本文と要約・メモの欄を消す。消すものが無ければ False"""
        props = page["properties"]
        upd = {name: {"rich_text": []} for name in CLEAR_PROPS
               if name in props and props[name].get("type") == "rich_text" and props[name].get("rich_text")}
        blocks = self.blocks(page["id"]) if body_blocks is None else body_blocks
        if upd:
            self.req("PATCH", f"pages/{page['id']}", json={"properties": upd})
        for i, b in enumerate(blocks, 1):
            self.req("DELETE", f"blocks/{b['id']}")
            if progress and (i % 10 == 0 or i == len(blocks)):
                progress(f"Notion の本文を消去中 {i}/{len(blocks)}")
        return bool(upd or blocks)

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
    """今の minutes の要約欄 → MinuteStock の形 (s, n)。マーカーは位置で持ち、ToDo の完了・経過は n に分ける"""
    agenda = []
    raw_agenda = prop_text(props, P_AGENDA)
    if raw_agenda:
        try:
            for a in json.loads(raw_agenda) or []:
                agenda.append({"topic": str(a.get("topic", "")).strip(), "t": str(a.get("t", "")), "outcome": str(a.get("outcome", "")).strip(),
                               "points": [p for p in a.get("points") or [] if p]})
        except ValueError:
            log("議事の JSON を読めませんでした（途中で切れている可能性）")
    todos = []
    for line in lines(prop_text(props, P_TODOS)):
        m = re.match(r"^\[([ xX])\]\s?(.*)$", line)
        todos.append(((m.group(2) if m else line).strip(), bool(m and m.group(1).lower() == "x")))
    gen = prop_date(props, P_GENERATED)
    at_ms = int(dt.datetime.fromisoformat(gen.replace("Z", "+00:00")).timestamp() * 1000) if len(gen) > 10 else int(time.time() * 1000)
    raw = {"card": prop_text(props, P_SUMMARY), "decisions": lines(prop_text(props, P_DECISIONS)), "todos": todos,
           "topics": lines(prop_text(props, P_TOPICS)), "agenda": agenda}
    s, n = build_new(raw, prop_text(props, P_TODO_LOGS), prop_text(props, P_MEMO), prop_text(props, P_MODEL), at_ms)
    mind = prop_text(props, P_MINDMAP)
    if mind:
        s["mindmap"] = mind  # 保管はするが、アプリでの表示はまだ
    return s, n


# ---------- 取り込み ----------

def new_id():
    import secrets
    a = "abcdefghijklmnopqrstuvwxyz0123456789"
    return "m_" + "".join(secrets.choice(a) for _ in range(12))


class Progress:
    """画面に進み具合を出す（ターミナルで実行したときだけ。監視からの実行ではログだけ）"""

    def __init__(self, total, label):
        self.total, self.label, self.i, self.t0 = total, label, 0, time.monotonic()
        self.tty = sys.stdout.isatty()

    def start(self, title):
        self.i += 1
        self.title = title[:30]
        self.show("読み込み中")

    def eta(self):
        if self.i <= 1:
            return ""
        per = (time.monotonic() - self.t0) / (self.i - 1)
        rest = int(per * (self.total - self.i + 1))
        return f" ・ 残り約{rest // 60}分{rest % 60:02d}秒" if rest >= 60 else f" ・ 残り約{rest}秒"

    def show(self, what):
        if self.tty:
            line = f"[{self.i}/{self.total}] {self.title} … {what}{self.eta()}"
            sys.stdout.write("\r\033[K" + line)
            sys.stdout.flush()

    def done(self, msg):
        if self.tty:
            sys.stdout.write("\r\033[K")
            sys.stdout.flush()
        log(f"[{self.i}/{self.total}] {msg}")


def already_in_github(gh, roles) -> dict:
    """GitHub に入っている PLAUD の会議 → {Notion のページID: 議事録ID}。
    GitHub に書いたあと、Notion に取込ID を書く前に止まった場合に、二重に取り込まないために使う"""
    admin = roles[ADMIN_ROLE]
    priv = jwk_to_priv(admin["priv"])
    out = {}
    for e in gh.get_json("index.json", []) or []:
        if e.get("src") != "plaud" or admin["gid"] not in (e.get("k") or {}):
            continue
        try:
            meta = open_json(unwrap_dek(e["k"][admin["gid"]], priv), e["m"])
            if meta.get("notion"):
                out[meta["notion"]] = e["id"]
        except Exception:  # noqa: BLE001
            pass
    return out


def import_pages(dry_run=False, limit=None, clear=None) -> int:
    token, db = ENV.get("NOTION_TOKEN"), ENV.get("NOTION_MINUTES_DB_ID")
    if not token or not db:
        return 0
    clear = ENV.get("NOTION_CLEAR_AFTER_IMPORT", "1") == "1" if clear is None else clear
    # 手で実行した移行と、5分おきの監視が同時に動いて二重に取り込まないようにする
    lock = open("/tmp/minutes2_import.lock", "w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        log("別の取り込みが動いているので、今回は見送ります")
        return 0
    keys = load_keys()
    roles = keys["roles"]
    admin_gid = roles[ADMIN_ROLE]["gid"]
    n = NotionDB(token, db)
    if dry_run:
        has_prop = n.has_import_prop()
    else:
        n.ensure_import_prop()
        has_prop = True
    gh = None if dry_run else GitHub()
    done_map = {} if dry_run else already_in_github(gh, roles)
    count = 0
    targets = list(n.pending(has_prop))
    if limit:
        targets = targets[:limit]
    if targets and sys.stdout.isatty():
        print(f"{'確認' if dry_run else '取り込み'}の対象：{len(targets)} 件")
    pg = Progress(len(targets), "取り込み")
    for page in targets:
        props = page["properties"]
        title = prop_text(props, P_TITLE) or "（無題）"
        status = prop_select(props, P_STATUS)
        pg.start(title)
        try:
            body_blocks = n.blocks(page["id"])
            text = "\n".join(t for t in (plain((b.get(b["type"]) or {}).get("rich_text")) for b in body_blocks) if t.strip())
            segs = segments_from_text(text)
            has_sum = status == ST_SUMMARIZED and (prop_date(props, P_GENERATED) or prop_text(props, P_SUMMARY))
            perms = prop_multi(props, P_PERM)
            unknown = [r for r in perms if r not in roles]
            gids = sorted({admin_gid, *[roles[r]["gid"] for r in perms if r in roles]})
            if dry_run:
                pg.done(f"[確認] {title}（{to_local(prop_date(props, P_DATE))}）{'要約済み' if has_sum else '要約待ち'} ・ 文字起こし {len(segs)} 行 ・ 閲覧 {', '.join(perms) or '管理者のみ'}"
                    + (f" ・ 対応する role が無い権限: {', '.join(unknown)}" if unknown else ""))
                count += 1
                continue
            if page["id"] in done_map:
                # 前回、GitHub に書いたところで止まっていた → 取込ID を書くところから続ける
                eid = done_map[page["id"]]
                n.req("PATCH", f"pages/{page['id']}", json={"properties": {
                    P_IMPORT_ID: {"rich_text": [{"text": {"content": eid}}]}, P_STATUS: {"select": {"name": ST_IMPORTED}}}})
                if clear:
                    n.clear_page(page, body_blocks, pg.show)
                pg.done(f"{title}: 前回の続きとして取込ID を書きました（{eid}）")
                count += 1
                continue
            if not segs and not has_sum:
                pg.done(f"{title}: 文字起こしが空なので見送ります")
                continue
            if unknown:
                log(f"{title}: 対応する role が無い権限は外しました（{', '.join(unknown)}）")

            pg.show("暗号化して GitHub に保存中")
            eid = new_id()
            dek = new_dek()
            k = {g: wrap_dek(dek, next(r["pub"] for r in roles.values() if r["gid"] == g)) for g in gids}
            meta = {"title": title, "tags": prop_multi(props, P_TAGS), "notion": page["id"]}
            s_new, notes = old_summary(props) if has_sum else (None, None)
            has_notes = bool(notes and (notes["done"] or notes["logs"] or notes["marks"] or notes.get("memo")))
            entry = {
                "id": eid, "date": to_local(prop_date(props, P_DATE)), "dur": parse_dur(prop_text(props, P_DUR)) or
                (max([x["t"] for x in segs if x["t"] is not None], default=0)),
                "st": "ok" if has_sum else "sw", "src": "plaud", "by": admin_gid,
                "upd": int(time.time() * 1000), "k": k, "m": seal_json(dek, meta),
            }
            if s_new is not None:
                entry["td"] = len(s_new["todos"])   # ToDo の件数（ToDo 一覧で、読まなくてよい会議を飛ばす）
            if has_notes:
                entry["nt"] = 1                     # ノート（メモ・マーカー・経過・完了）がある
            gh.write(f"m/{eid}/t.json", json.dumps(seal_json(dek, {"segments": segs, "asr": "Notion から取り込み"})), f"t {eid}")
            if s_new is not None:
                gh.write(f"m/{eid}/s.json", json.dumps(seal_json(dek, s_new)), f"s {eid}")
                if has_notes:
                    gh.write(f"m/{eid}/n.json", json.dumps(seal_json(dek, notes)), f"n {eid}")
            gh.update_index(lambda lst: lst.append(entry))

            # Notion 側：取込ID と状態を書き、必要なら中身を消す
            upd = {P_IMPORT_ID: {"rich_text": [{"text": {"content": eid}}]}, P_STATUS: {"select": {"name": ST_IMPORTED}}}
            n.req("PATCH", f"pages/{page['id']}", json={"properties": upd})
            if clear:
                n.clear_page(page, body_blocks, pg.show)
            pg.done(f"{title}: 取り込みました（{eid}・{'要約済み' if has_sum else '要約待ち'}{'・Notion の本文を消去' if clear else ''}）")
            count += 1
        except Exception as ex:  # noqa: BLE001  1件の失敗で全体を止めない
            pg.done(f"{title}: 取り込めませんでした {ex}")
    return count


def upgrade_imported() -> int:
    """前の版の notion_import.py で取り込んだ会議（マーカーのタグが文字に混ざっている、ToDo経過が生のまま）を、新しい形に直す"""
    keys = load_keys()
    admin = keys["roles"][ADMIN_ROLE]
    priv = jwk_to_priv(admin["priv"])
    gh = GitHub()
    targets = [e for e in (gh.get_json("index.json", []) or []) if e.get("src") == "plaud" and admin["gid"] in (e.get("k") or {})]
    if sys.stdout.isatty():
        print(f"確かめる取り込み済みの会議：{len(targets)} 件")
    pg = Progress(len(targets), "更新")
    changed = 0
    for e in targets:
        pg.start(e["id"])
        try:
            dek = unwrap_dek(e["k"][admin["gid"]], priv)
            sv = gh.get_json(f"m/{e['id']}/s.json")
            if sv is None:
                pg.done(f"{e['id']}: 要約が無いので飛ばします")
                continue
            nv = gh.get_json(f"m/{e['id']}/n.json")
            s_old, n_old = open_json(dek, sv), (open_json(dek, nv) if nv else {})
            if not needs_upgrade(s_old, n_old):
                pg.done(f"{e['id']}: 直す必要はありません")
                continue
            s_new, n_new = upgrade_pair(s_old, n_old)
            gh.write(f"m/{e['id']}/s.json", json.dumps(seal_json(dek, s_new)), f"s {e['id']}")
            gh.write(f"m/{e['id']}/n.json", json.dumps(seal_json(dek, n_new)), f"n {e['id']}")
            def fn(lst, eid=e["id"], td=len(s_new["todos"])):
                for x in lst:
                    if x["id"] == eid:
                        x["td"] = td
                        x["nt"] = 1
                        x["upd"] = int(time.time() * 1000)
            gh.update_index(fn)
            changed += 1
            pg.done(f"{e['id']}: 新しい形に直しました（マーカー {len(n_new['marks'])} 件・経過のある ToDo {len(n_new['logs'])} 件）")
        except Exception as ex:  # noqa: BLE001
            pg.done(f"{e['id']}: 直せませんでした {ex}")
    return changed


def clear_imported() -> int:
    """--keep で残した、取り込み済みページの本文・要約を消す"""
    n = NotionDB(ENV["NOTION_TOKEN"], ENV["NOTION_MINUTES_DB_ID"])
    pages = list(n.imported())
    if sys.stdout.isatty():
        print(f"確かめる取り込み済みページ：{len(pages)} 件")
    pg = Progress(len(pages), "消去")
    count = 0
    for page in pages:
        title = prop_text(page["properties"], P_TITLE) or "（無題）"
        pg.start(title)
        try:
            if n.clear_page(page, progress=pg.show):
                count += 1
                pg.done(f"{title}: Notion の本文・要約を消しました")
            else:
                pg.done(f"{title}: 消すものはありませんでした")
        except Exception as ex:  # noqa: BLE001
            pg.done(f"{title}: 消せませんでした {ex}")
    return count


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dry-run", action="store_true", help="何も書かずに、取り込める件数だけ見る")
    p.add_argument("--limit", type=int, default=None, help="取り込む件数の上限")
    p.add_argument("--keep", action="store_true", help="取り込んでも Notion の本文・要約を消さない")
    p.add_argument("--clear-imported", action="store_true", help="取り込み済みで、まだ本文が残っているページを消す")
    p.add_argument("--upgrade", action="store_true", help="前の版で取り込んだ会議を、マーカー・ToDo経過の新しい形に直す")
    a = p.parse_args()
    if not ENV.get("NOTION_TOKEN") or not ENV.get("NOTION_MINUTES_DB_ID"):
        raise SystemExit(".env に NOTION_TOKEN と NOTION_MINUTES_DB_ID を書いてください")
    try:
        if a.upgrade:
            print(f"更新：{upgrade_imported()} 件")
            return
        if a.clear_imported:
            print(f"消去：{clear_imported()} 件")
            return
        n = import_pages(dry_run=a.dry_run, limit=a.limit, clear=False if a.keep else None)
        print(f"{'確認' if a.dry_run else '取り込み'}：{n} 件")
    except KeyboardInterrupt:
        print("\n中断しました。もう一度実行すると続きから取り込みます。"
              "Notion の本文を消している途中で止めたページは、--clear-imported で消し直せます。")


if __name__ == "__main__":
    main()
