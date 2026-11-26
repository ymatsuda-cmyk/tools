#!/usr/bin/env python3
"""minutes_worker.py — Mac の定期監視（launchd で5分おきに起動）

1. 文字起こし待ち（tw）の音声を、Drive の受付フォルダ（Google ドライブ for desktop で同期）から読み、
   mlx_whisper で文字起こしして、暗号化して GitHub に保存 → 要約待ち（sw）
0. （NOTION_MINUTES_DB_ID があれば）Notion 議事録DB（PLAUD Note の受け口）から取り込む → notion_import.py
2. （NOTION_TOKEN / NOTION_WORK_DB_ID があれば）Claude.ai 用の受け渡し
   ・要約待ちを Notion の作業DBへ書き出す
   ・Claude.ai が書き込んだ要約（状態＝要約済み）を回収し、暗号化して GitHub に保存 → 要約済み（ok）
   ・回収した、またはアプリで要約済みになったページは Notion から消す（ゴミ箱へ）
3. worker.json に確認時刻を書く（アプリの変換管理に「最終確認 ○分前」と出る）

Mac は管理者の鍵を keys.json に持つので、すべての議事録を開ける。
"""
import fcntl
import json
import pathlib
import sys
import threading
import time
import traceback

import requests

from m2common import ENV, ENV_OVERRIDES, GitHub, HOME, log, load_keys
from m2crypto import jwk_to_priv, open_json, seal_json, unwrap_dek

ADMIN_ROLE = "xYz"
EVERY = int(ENV.get("MINUTES2_EVERY", "300"))
MODEL = ENV.get("WHISPER_MODEL", "mlx-community/whisper-large-v3-turbo")
DRIVE_DIR = pathlib.Path(ENV.get("DRIVE_DIR", "")).expanduser()
WAIT_SYNC_SEC = 3 * 3600  # Drive の同期をこれだけ待っても音声が来なければ失敗にする
STATE = HOME / "worker_state.json"


class Worker:
    def __init__(self):
        self.gh = GitHub()
        keys = load_keys()
        admin = keys["roles"][ADMIN_ROLE]
        self.admin_gid = admin["gid"]
        self.admin_priv = jwk_to_priv(admin["priv"])
        self.dek_cache = {}
        self.notion = None   # Claude.ai との受け渡しの状態（worker.json に載せて、アプリに見せる）。前回の値から始める
        try:
            self.notion = json.loads(STATE.read_text()).get("notion") if STATE.exists() else None
        except ValueError:
            self.notion = None
        self.stats = {"取り込み": 0, "文字起こし": 0, "Notionへ書き出し": 0, "要約を回収": 0}   # 1回の確認でやったこと（最後にログへ1行で残す）

    # ---------- 共通 ----------

    def dek(self, e) -> bytes:
        if e["id"] not in self.dek_cache:
            w = (e.get("k") or {}).get(self.admin_gid)
            if not w:
                raise RuntimeError("管理者の鍵がありません")
            self.dek_cache[e["id"]] = unwrap_dek(w, self.admin_priv)
        return self.dek_cache[e["id"]]

    def read_file(self, e, name):
        v = self.gh.get_json(f"m/{e['id']}/{name}.json")
        return None if v is None else open_json(self.dek(e), v)

    def write_file(self, e, name, obj):
        self.gh.write(f"m/{e['id']}/{name}.json", json.dumps(seal_json(self.dek(e), obj)), f"{name} {e['id']}")

    def patch(self, eid, **fields):
        def fn(lst):
            for x in lst:
                if x["id"] == eid:
                    for k, v in fields.items():
                        if v is None:
                            x.pop(k, None)
                        else:
                            x[k] = v
                    x["upd"] = int(time.time() * 1000)
                    return x
            return None
        return self.gh.update_index(fn)

    def heartbeat(self, busy=None, force=False):
        st = {}
        if STATE.exists():
            try:
                st = json.loads(STATE.read_text())
            except ValueError:
                st = {}
        now = int(time.time() * 1000)
        # 毎回コミットすると履歴が増えるので、状態が変わったときと 15 分ごとだけ書く
        if not force and st.get("busy") == busy and st.get("notion") == self.notion and now - st.get("at", 0) < 15 * 60 * 1000:
            return
        body = {"at": now, "every": EVERY, "busy": busy, "asr": MODEL.split("/")[-1]}
        if self.notion is not None:
            body["notion"] = self.notion   # {ok, err, ids: 作業DBに要約待ちで置いてある議事録ID}
        try:
            self.gh.write("worker.json", json.dumps(body), "worker")
            STATE.write_text(json.dumps({"at": now, "busy": busy, "notion": self.notion}))
        except Exception as ex:  # noqa: BLE001  確認時刻が書けなくても本処理は続ける
            log(f"worker.json を書けませんでした: {ex}")

    # ---------- 1. 文字起こし ----------

    def transcribe_waiting(self, index):
        waiting = sorted([e for e in index if e.get("st") == "tw" and e.get("src") != "text"], key=lambda e: e.get("upd", 0))
        for e in waiting:
            started = False
            try:
                meta = open_json(self.dek(e), e["m"])
                f = (meta.get("file") or {})
                path = DRIVE_DIR / f.get("name", "")
                if not f.get("name") or not path.is_file():
                    if time.time() * 1000 - e.get("upd", 0) > WAIT_SYNC_SEC * 1000:
                        self.patch(e["id"], st="err", err="音声が Mac に届きません（Drive の同期を確かめてください）")
                        log(f"{e['id']}: 音声が見つからないまま時間切れ")
                    else:
                        log(f"{e['id']}: Drive の同期待ち（{path.name}）")
                    continue
                if not self._synced(path):
                    log(f"{e['id']}: 同期の途中なので次回に回します")
                    continue
                log(f"{e['id']}: 文字起こし開始 {path.name}")
                self.patch(e["id"], st="tr")
                started = True
                self.heartbeat(busy=e["id"], force=True)
                stop = self._beat_while(e["id"])   # 長い録音でも、アプリが「応答なし」と誤らないよう、数分おきに知らせる
                try:
                    segments, dur = transcribe(path)
                finally:
                    stop()   # 終わりを待ってから、あとの書き込みへ（同じブランチへの同時の書き込みを避ける）
                if not segments:
                    raise RuntimeError("音声から文字を起こせませんでした（無音の可能性があります）")
                self.write_file(e, "t", {"segments": segments, "asr": MODEL})
                self.patch(e["id"], st="sw", dur=int(dur or e.get("dur") or 0), err=None)
                self.stats["文字起こし"] += 1
                log(f"{e['id']}: 文字起こし完了 {len(segments)} 区切り / {int(dur)} 秒")
            except Exception as ex:  # noqa: BLE001
                log(f"{e['id']}: 失敗 {ex}\n{traceback.format_exc()}")
                try:
                    self.patch(e["id"], st="err", err=str(ex)[:200])
                except Exception:  # noqa: BLE001
                    pass
            finally:
                if started:
                    self.heartbeat(busy=None, force=True)

    def _beat_while(self, busy, every=240):
        """文字起こしの間、every 秒おきに「文字起こし中」の確認を書く。返す関数を呼ぶと止まり、最後の1回が終わるのを待つ"""
        done = threading.Event()

        def loop():
            while not done.wait(every):
                try:
                    self.heartbeat(busy=busy, force=True)
                except Exception as ex:  # noqa: BLE001  知らせが書けなくても、文字起こしは続ける
                    log(f"文字起こし中の確認を書けませんでした: {ex}")

        t = threading.Thread(target=loop, name="minutes2-beat", daemon=True)
        t.start()

        def stop():
            done.set()
            t.join(timeout=60)
        return stop

    @staticmethod
    def _synced(path: pathlib.Path) -> bool:
        """サイズが 5 秒変わらなければ同期済みとみなす"""
        a = path.stat().st_size
        time.sleep(5)
        return a > 0 and a == path.stat().st_size

    # ---------- 2. Claude.ai との受け渡し（Notion） ----------

    def notion_bridge(self, index):
        token, db = ENV.get("NOTION_TOKEN"), ENV.get("NOTION_WORK_DB_ID")
        if not token or not db:
            if any(e.get("st") == "sw" for e in index):
                log("要約待ちがありますが、.env の NOTION_TOKEN / NOTION_WORK_DB_ID が空のため、Notion の作業DBへ書き出していません")
            self.notion = {"ok": False, "err": "Mac の .env に NOTION_TOKEN / NOTION_WORK_DB_ID がありません", "ids": []}
            return
        try:
            self._notion_bridge(Notion(token, db), index)
        except Exception as ex:  # noqa: BLE001  受け渡しが失敗しても、文字起こしなどは続ける
            msg = str(ex)
            if "404" in msg:
                msg = ("作業DBが見つかりません（NOTION_WORK_DB_ID がデータベースのID〈URLの32桁〉か、"
                       "インテグレーションをDBに接続したかを確かめてください）: " + msg)
            log(f"Notion の作業DBとの受け渡しに失敗しました: {msg}")
            self.notion = {"ok": False, "err": msg[:300], "ids": []}

    def _notion_bridge(self, n, index):
        pages = n.pages()  # {議事録ID: page}
        by_id = {e["id"]: e for e in index}

        # 回収：Claude.ai が要約済みにしたもの
        for eid, page in list(pages.items()):
            if n.status(page) != "要約済み":
                continue
            e = by_id.get(eid)
            if not e or e.get("st") not in ("sw", "ok"):
                n.trash(page["id"]); pages.pop(eid)
                continue
            try:
                raw = n.summary_json(page["id"])
                s = normalize(raw)
                self.write_file(e, "s", s)
                self.patch(eid, st="ok", err=None, td=len(s["todos"]))
                e["st"] = "ok"  # 手元の一覧も更新して、下の書き出しで二重に出さない
                n.trash(page["id"]); pages.pop(eid)
                self.stats["要約を回収"] += 1
                log(f"{eid}: Claude.ai の要約を回収")
            except Exception as ex:  # noqa: BLE001
                log(f"{eid}: 要約を読み取れません {ex}")
                n.set_status(page["id"], "要約待ち", note=f"要約JSONを読み取れませんでした: {str(ex)[:120]}")

        # 片付け：もう要約待ちではないもの
        for eid, page in list(pages.items()):
            e = by_id.get(eid)
            if not e or e.get("st") != "sw":
                n.trash(page["id"]); pages.pop(eid)

        # 書き出し：要約待ちで、まだ Notion に無いもの
        for e in index:
            if e.get("st") != "sw" or e["id"] in pages:
                continue
            try:
                meta = open_json(self.dek(e), e["m"])
                t = self.read_file(e, "t") or {}
                n.create(e["id"], meta.get("title") or "（無題）", e.get("date", ""), e.get("dur", 0), t.get("segments") or [])
                pages[e["id"]] = {"properties": {"状態": {"select": {"name": "要約待ち"}}}}
                self.stats["Notionへ書き出し"] += 1
                log(f"{e['id']}: Notion の作業DBへ書き出し")
            except Exception as ex:  # noqa: BLE001
                log(f"{e['id']}: Notion へ書き出せません {ex}")
        self.notion = {"ok": True, "err": "", "ids": sorted(eid for eid, p in pages.items() if n.status(p) == "要約待ち")}

    # ---------- 実行 ----------

    def run(self):
        try:
            self.gh.check()
        except Exception as ex:  # noqa: BLE001  リポジトリに入れないまま続けても、何もできない（何もしない回に見えてしまう）
            log(f"GitHub につながりません。この回は何もしません: {ex}")
            return
        self.heartbeat()
        try:
            from notion_import import import_pages
            self.stats["取り込み"] = import_pages() or 0
        except Exception as ex:  # noqa: BLE001  取り込みが失敗しても文字起こしは続ける
            log(f"Notion からの取り込みに失敗しました: {ex}")
        index = self.gh.get_json("index.json", []) or []
        self.transcribe_waiting(index)
        index = self.gh.get_json("index.json", []) or []
        self.notion_bridge(index)
        self.heartbeat()
        # 何も無かった回も、確認したことが分かるように1行残す（動いているのにログが空、を無くす）
        log("確認しました：" + "・".join(f"{k} {v}件" for k, v in self.stats.items()))


def transcribe(path: pathlib.Path):
    """→ ([{t: 秒, s: 文}], 長さ秒)。ASR を替えるときはここだけ直す"""
    import mlx_whisper  # Mac の Apple Silicon でだけ使える
    r = mlx_whisper.transcribe(str(path), path_or_hf_repo=MODEL, language="ja",
                               condition_on_previous_text=False, verbose=None)
    segs, dur = [], 0.0
    for s in r.get("segments", []):
        text = (s.get("text") or "").strip()
        dur = max(dur, float(s.get("end") or 0))
        if text:
            segs.append({"t": round(float(s.get("start") or 0), 1), "s": text})
    return segs, dur


def fmt_t(sec):
    if sec is None:
        return ""
    s = int(sec)
    h, m, r = s // 3600, (s % 3600) // 60, s % 60
    return f"{h}:{m:02d}:{r:02d}" if h else f"{m}:{r:02d}"


def normalize(raw: dict) -> dict:
    """Claude.ai の要約 JSON を、アプリの形にそろえる（web/src/summary.js の normalizeSummary と同じ考え方）"""
    if not isinstance(raw, dict):
        raise ValueError("JSON がオブジェクトではありません")

    def item(v):
        if isinstance(v, str):
            return {"text": v.strip(), "t": ""}
        return {"text": str((v or {}).get("text", "")).strip(), "t": str((v or {}).get("t", "")).strip()}

    agenda = []
    for a in raw.get("agenda") or []:
        agenda.append({
            "topic": str(a.get("topic", "")).strip() or "（議題）",
            "t": str(a.get("t", "")).strip(),
            "outcome": str(a.get("outcome", "")).strip(),
            "points": [p for p in (item(x) for x in a.get("points") or []) if p["text"]],
        })
    todos = []
    for t in raw.get("todos") or []:
        if isinstance(t, str):
            todos.append({"text": t.strip(), "owner": "", "due": ""})
        else:
            todos.append({"text": str(t.get("text", "")).strip(), "owner": str(t.get("owner", "")).strip(), "due": str(t.get("due", "")).strip()})
    todos = [t for t in todos if t["text"]]
    for i, t in enumerate(todos):
        t["id"] = f"i{i}"   # 完了・経過を結び付ける識別子（アプリの todoKey と同じ規則）
    s = {
        "cardSummary": str(raw.get("cardSummary", "")).strip(),
        "summary": str(raw.get("summary", "") or raw.get("cardSummary", "")).strip(),
        "agenda": agenda,
        "decisions": [d for d in (item(x) for x in raw.get("decisions") or []) if d["text"]],
        "todos": todos,
        "topics": [str(x).strip() for x in raw.get("topics") or [] if str(x).strip()],
        "model": str(raw.get("model") or "Claude.ai"),
        "at": int(time.time() * 1000),
    }
    if not s["agenda"] and not s["summary"]:
        raise ValueError("要約が空です")
    return s


class Notion:
    """Claude.ai との受け渡し用の作業DB。プロパティ：名前（タイトル）／ID（テキスト）／状態（セレクト：要約待ち・要約済み）／メモ（テキスト）"""

    def __init__(self, token, db):
        self.db = db
        self.s = requests.Session()
        self.s.headers.update({"Authorization": "Bearer " + token, "Notion-Version": "2022-06-28", "Content-Type": "application/json"})

    def _req(self, method, path, **kw):
        for i in range(4):
            r = self.s.request(method, "https://api.notion.com/v1/" + path, timeout=60, **kw)
            if r.status_code == 429:
                time.sleep(1 + i * 2)
                continue
            if r.status_code >= 400:   # Notion の説明（「インテグレーションと共有されていません」など）をそのまま出す
                try:
                    j = r.json()
                except ValueError:
                    j = {}
                raise RuntimeError(f"{r.status_code} {j.get('code', '')}: {j.get('message') or r.text[:200]}")
            return r.json()
        raise RuntimeError("429 rate_limited: Notion への呼び出しが多すぎます")

    def pages(self):
        out, cursor = {}, None
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            j = self._req("POST", f"databases/{self.db}/query", json=body)
            for p in j["results"]:
                rt = p["properties"].get("ID", {}).get("rich_text", [])
                eid = "".join(x.get("plain_text", "") for x in rt).strip()
                if eid:
                    out[eid] = p
            if not j.get("has_more"):
                return out
            cursor = j["next_cursor"]

    @staticmethod
    def status(page):
        sel = page["properties"].get("状態", {}).get("select")
        return sel["name"] if sel else ""

    def set_status(self, page_id, name, note=""):
        props = {"状態": {"select": {"name": name}}}
        if note:
            props["メモ"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
        self._req("PATCH", f"pages/{page_id}", json={"properties": props})

    def trash(self, page_id):
        self._req("PATCH", f"pages/{page_id}", json={"archived": True})

    def create(self, eid, title, date, dur, segments):
        def para(text):
            return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": [{"type": "text", "text": {"content": text}}]}}
        blocks = [para(f"会議日時: {date.replace('T', ' ')} / 長さ: {fmt_t(dur)}"), para("文字起こし（行頭の [m:ss] は発言の時刻）:")]
        buf = ""
        for s in segments:
            line = (f"[{fmt_t(s.get('t'))}] " if s.get("t") is not None else "") + s.get("s", "")
            if len(buf) + len(line) + 1 > 1800:
                blocks.append(para(buf)); buf = ""
            buf = (buf + "\n" + line) if buf else line[:1800]
        if buf:
            blocks.append(para(buf))
        props = {
            "名前": {"title": [{"text": {"content": title[:200]}}]},
            "ID": {"rich_text": [{"text": {"content": eid}}]},
            "状態": {"select": {"name": "要約待ち"}},
        }
        page = self._req("POST", "pages", json={"parent": {"database_id": self.db}, "properties": props, "children": blocks[:100]})
        for i in range(100, len(blocks), 100):
            self._req("PATCH", f"blocks/{page['id']}/children", json={"children": blocks[i:i + 100]})

    def summary_json(self, page_id):
        """ページ内の json コードブロックを（分割されていてもつなげて）読む"""
        parts, cursor = [], None
        while True:
            q = f"blocks/{page_id}/children?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
            j = self._req("GET", q)
            for b in j["results"]:
                if b["type"] == "code":
                    parts.append("".join(x.get("plain_text", "") for x in b["code"]["rich_text"]))
            if not j.get("has_more"):
                break
            cursor = j["next_cursor"]
        if not parts:
            raise ValueError("要約JSON のコードブロックがありません")
        text = "".join(parts).strip()
        a, b = text.find("{"), text.rfind("}")
        return json.loads(text[a:b + 1])


def main():
    HOME.mkdir(parents=True, exist_ok=True)
    for k in ENV_OVERRIDES:
        log(f"注意：環境変数 {k} が、.env の {k} と違う値で、.env より優先されています。意図していなければ unset {k} してください")
    lock = open("/tmp/minutes2_worker.lock", "w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        log("前回の処理がまだ動いているので、今回は見送ります")
        return
    if not DRIVE_DIR.is_dir():
        log(f"DRIVE_DIR が見つかりません: {DRIVE_DIR}")
    try:
        Worker().run()
    except Exception as ex:  # noqa: BLE001
        log(f"処理を中断しました: {ex}\n{traceback.format_exc()}")
        sys.exit(1)


if __name__ == "__main__":
    main()
