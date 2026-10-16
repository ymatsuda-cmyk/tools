#!/usr/bin/env python3
"""情報漏洩事例フィード（cases-feed.json）の更新スクリプト。

  python fetch_cases.py --feed ../cases-feed.json --sources sources.json [--dry-run]

・標準ライブラリだけで動きます（pip install は不要）。
・sources.json に書いた情報源（RSS/Atom、または記事一覧ページ）から新着を集め、
  すでに cases-feed.json にある記事（URLで判定）は追加しません。
・保存するのは「タイトル・公表日・元記事URL・原因タイプ（自動判定）」だけです。記事の本文は保存しません。
・企業名・件数・原因タイプはタイトルからの簡易判定です。誤りはExcelの「漏洩事例」シートで直してください。
"""
import argparse, datetime, email.utils, hashlib, html as htmllib, json, os, re, sys, time
import urllib.parse, urllib.request
from html.parser import HTMLParser
from xml.etree import ElementTree as ET

UA = "Mozilla/5.0 (compatible; leak-cases-feed/1.0)"
JST = datetime.timezone(datetime.timedelta(hours=9))

# ---------- 取得 ----------
def http_get(url, timeout=30, retries=2, pause=1.0):
    last = None
    for i in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                raw = r.read()
                charset = None
                try:
                    charset = r.headers.get_content_charset()
                except Exception:
                    pass
            if not charset:
                m = re.search(rb'<meta[^>]+charset=["\']?([A-Za-z0-9_\-]+)', raw[:4096], re.I) or re.search(rb'encoding=["\']([A-Za-z0-9_\-]+)', raw[:200], re.I)
                charset = m.group(1).decode("ascii") if m else "utf-8"
            try:
                return raw.decode(charset, errors="replace")
            except LookupError:
                return raw.decode("utf-8", errors="replace")
        except Exception as e:  # noqa
            last = e
            time.sleep(pause * (i + 1))
    raise RuntimeError(f"取得できませんでした: {url} ({last})")

# ---------- 日付 ----------
def to_jst_date(dt):
    if dt.tzinfo is None:
        return dt.date().isoformat()
    return dt.astimezone(JST).date().isoformat()

def parse_date_text(s):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return to_jst_date(email.utils.parsedate_to_datetime(s))
    except Exception:
        pass
    try:
        return to_jst_date(datetime.datetime.fromisoformat(s.replace("Z", "+00:00")))
    except Exception:
        pass
    m = re.search(r"(20\d{2})\s*[年/\.\-]\s*(\d{1,2})\s*[月/\.\-]\s*(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
        except ValueError:
            return None
    return None

# ---------- RSS / Atom ----------
def _local(tag):
    return tag.split("}")[-1] if isinstance(tag, str) else ""

def parse_feed(xml_text):
    root = ET.fromstring(xml_text.lstrip("\ufeff"))
    out = []
    for el in root.iter():
        if _local(el.tag) not in ("item", "entry"):
            continue
        d = {"title": "", "link": "", "date": None}
        for ch in el:
            n = _local(ch.tag)
            txt = (ch.text or "").strip()
            if n == "title":
                d["title"] = re.sub(r"\s+", " ", txt)
            elif n == "link":
                d["link"] = ch.attrib.get("href") or txt or d["link"]
            elif n in ("pubDate", "date", "published", "updated") and not d["date"]:
                d["date"] = parse_date_text(txt)
        if d["title"] and d["link"]:
            out.append(d)
    return out

# ---------- 記事一覧ページ ----------
class _Links(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links, self._href, self._buf = [], None, []
    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self._href = dict(attrs).get("href")
            self._buf = []
    def handle_data(self, data):
        if self._href is not None:
            self._buf.append(data)
    def handle_endtag(self, tag):
        if tag == "a" and self._href is not None:
            self.links.append((self._href, re.sub(r"\s+", " ", "".join(self._buf)).strip()))
            self._href = None

def parse_link_list(html_text, base_url, link_regex):
    p = _Links()
    p.feed(html_text)
    rx = re.compile(link_regex)
    seen, out = set(), []
    for href, text in p.links:
        url = urllib.parse.urljoin(base_url, href).split("#")[0]
        if not rx.search(url) or url in seen:
            continue
        seen.add(url)
        out.append({"link": url, "title": text})
    return out

def article_meta(html_text):
    """記事ページから タイトル と 公表日 を取り出す。"""
    title = None
    m = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']+)', html_text, re.I) or \
        re.search(r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:title', html_text, re.I)
    if m:
        title = m.group(1)
    if not title:
        m = re.search(r"<title[^>]*>(.*?)</title>", html_text, re.S | re.I)
        title = m.group(1) if m else ""
    title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", title)).strip()
    title = re.split(r"\s*[|｜]\s*サイバーセキュリティ", title)[0].strip()
    date = None
    for rx in (r'datePublished"\s*:\s*"([^"]+)"',
               r'<meta[^>]+property=["\']article:published_time["\'][^>]+content=["\']([^"\']+)',
               r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']article:published_time',
               r'<time[^>]+datetime=["\']([^"\']+)'):
        m = re.search(rx, html_text, re.I)
        if m:
            date = parse_date_text(m.group(1))
            if date:
                break
    if not date:
        body = re.sub(r"<script.*?</script>|<style.*?</style>", "", html_text, flags=re.S | re.I)
        date = parse_date_text(re.sub(r"<[^>]+>", " ", body)[:6000])
    return {"title": title, "date": date}

# ---------- タイトルからの簡易判定 ----------
RELEVANT = re.compile(r"漏えい|漏洩|流出|不正アクセス|不正ログイン|誤送信|誤送付|紛失|盗難|ランサム|サイバー攻撃|個人情報|情報流出|閲覧可能|乗っ取り|なりすまし|持ち出し|持出し|内部不正|不正取得|情報窃取|改ざん")
CAT_RULES = [
    ("ransom", r"ランサム|身代金|暗号化|マルウェア|ウイルス|感染"),
    ("chain",  r"委託先|委託|再委託|外部サービス|クラウド|連携|SaaS|サプライチェーン|提供元|ベンダー|事業者のサーバ|基盤"),
    ("acct",   r"乗っ取り|不正ログイン|なりすまし|クレデンシャル|フィッシング|アカウント"),
    ("human",  r"誤送信|誤送付|誤って|誤公開|設定ミス|設定不備|設定の不備|公開状態|閲覧可能|CC|BCC|誤配|誤操作"),
    ("ext",    r"不正アクセス|脆弱性|侵入|サイバー攻撃|外部から"),
]
def classify(title):
    for cat, rx in CAT_RULES:
        if re.search(rx, title):
            return cat
    return "other"

_NUM = r"(?:\d[\d,]*\.?\d*億)?(?:\d[\d,]*\.?\d*万)?(?:\d[\d,]*)?"
_COUNT = re.compile(r"(約|最大約|最大|のべ|延べ|少なくとも)?\s*(" + _NUM + r")\s*(件|人|名|アカウント|口座|台|通|社)")
def _jp_number(txt):
    m = re.fullmatch(r"(?:(\d[\d,]*\.?\d*)億)?(?:(\d[\d,]*\.?\d*)万)?(\d[\d,]*)?", txt)
    if not m or not any(m.groups()):
        return 0
    f = lambda g: float(g.replace(",", "")) if g else 0.0
    return int(f(m.group(1)) * 1e8 + f(m.group(2)) * 1e4 + f(m.group(3)))
def extract_count(title):
    for m in _COUNT.finditer(title):
        n = _jp_number(m.group(2))
        if n:
            return n, re.sub(r"\s+", "", m.group(0))
    return 0, ""

_GENERIC = re.compile(r"\d+[万億件人]|某|同社|被害|不正|漏えい|漏洩|流出|攻撃|ランサム|委託先|感染|情報|個人|従業員|元社員|顧客|会員|閲覧")
def extract_company(title):
    """タイトルの先頭から企業・サービス名らしい部分を取り出す（当たらないときは空欄）。"""
    t = re.sub(r"^[【\[（(][^】\]）)]{1,12}[】\]）)]\s*", "", title.strip())
    t = re.sub(r"(株式会社|\(株\)|（株）)", "", t)
    m = re.match(r"^「([^」]{1,24})」", t)
    if m:
        return m.group(1)
    m = re.search(r"、|，|への|における|から|で|に|が|を|は|も|へ", t)
    head = t[: m.start()] if m else ""
    head = head.strip(" 　")
    if not (2 <= len(head) <= 30) or _GENERIC.search(head):
        return ""
    return head

# ---------- フィード本体 ----------
def norm_url(u):
    return re.sub(r"[#?].*$", "", (u or "").strip()).rstrip("/")

def make_id(url):
    m = re.search(r"/news/(\d+)", url)
    return "cs" + m.group(1) if m else "a" + hashlib.sha1(url.encode("utf-8")).hexdigest()[:10]

def new_item(src_name, url, title, date, today):
    n, nt = extract_count(title)
    d = date or today
    return {"id": make_id(url), "pub": d, "base": d, "baseType": "公表",
            "co": extract_company(title), "svc": "", "n": n, "nt": nt, "cat": classify(title),
            "imp": 0, "impWhy": "", "summary": title, "rel": "", "url": url,
            "src": src_name, "got": today}

def collect(source, log):
    name = source["name"]
    today = datetime.datetime.now(JST).date().isoformat()
    items = []
    if source["type"] == "rss":
        for e in parse_feed(http_get(source["url"])):
            items.append((e["link"], e["title"], e["date"]))
    elif source["type"] == "html-list":
        page = http_get(source["url"])
        cands = parse_link_list(page, source["url"], source["link_regex"])
        log(f"  一覧から記事候補 {len(cands)} 件")
        for c in cands[: int(source.get("max_articles", 40))]:
            if source.get("_skip") and norm_url(c["link"]) in source["_skip"]:
                continue
            try:
                time.sleep(float(source.get("delay_sec", 1.0)))
                meta = article_meta(http_get(c["link"]))
                items.append((c["link"], meta["title"] or c["title"], meta["date"]))
            except Exception as e:  # noqa
                log(f"  ! 記事を読めませんでした: {c['link']} ({e})")
    else:
        raise ValueError(f"未対応のtype: {source['type']}")
    out = []
    for link, title, date in items:
        if not title or not RELEVANT.search(title):
            continue
        out.append(new_item(name, link, title, date, today))
    return out


# ---------- 元記事から詳細を整理する ----------
SEC_KEYS = ["sec_date", "sec_summary", "sec_discovery", "sec_risk", "sec_cause", "sec_response", "sec_vendor"]
CATS = {"ext", "chain", "ransom", "acct", "human", "other"}
SYSTEM_PROMPT = (
    "あなたは情報セキュリティの調査担当です。与えられた記事本文だけを根拠に、情報漏洩事例を整理してください。"
    "記事に書かれていないことは推測せず、「記事に記載なし」と書きます。"
    "本文の文章をそのまま写さず、自分の言葉で簡潔にまとめます（引用する場合は短い語句だけ）。"
    "出力はJSONオブジェクトのみ（説明文やコードブロックは付けない）。"
)
USER_TEMPLATE = """次の記事を整理し、下記のキーを持つJSONで答えてください。
- co: 企業・団体名
- svc: 対象のサービス・システム名（不明なら空文字）
- title: 事例名（40字以内）
- sub: 運営会社や委託関係の補足（不明なら空文字）
- cat: ext(外部からの不正アクセス)/chain(委託先・クラウド経由)/ransom(ランサムウェア・マルウェア)/acct(メール・アカウント乗っ取り)/human(誤送信・設定ミス・人的)/other(改ざん・その他) のいずれか
- n: 漏洩した件数（数値。不明なら0）
- nt: 件数の表記（例：約660万件。不明なら空文字）
- base: 発生日（不明なら検知日、それも不明なら公表日）。YYYY-MM-DD
- baseType: base が 発生/検知/公表 のどれか
- facts: [["検知","2026年9月25日"],["漏洩","約660万件"]] の形式（最大4組）
- sec_date: 時系列（発生・検知・公表など）。各行を「・」で始める
- sec_summary: 概要（2〜4行。各行を「・」で始める）
- sec_discovery: 発覚の経緯
- sec_risk: リスク・影響範囲（漏洩した情報、想定される被害）
- sec_cause: 原因（判明している原因と、被害を広げた要因）
- sec_response: 対処（暫定対策と恒久対策に分けて書く）
- sec_vendor: システムを導入・保守する業者として取るべき対策を3〜5点（記事の要約ではなく、この事例から導く提案として書く）

記事のURL: {url}
記事のタイトル: {title}
記事の本文:
{body}
"""

def html_to_text(html_text, limit=14000):
    t = re.sub(r"(?is)<(script|style|noscript|nav|header|footer|aside|form|svg)[^>]*>.*?</\1>", "", html_text)
    m = re.search(r"(?is)<article[^>]*>(.*?)</article>", t) or re.search(r"(?is)<main[^>]*>(.*?)</main>", t)
    if m:
        t = m.group(1)
    t = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h[1-6]|tr|td)>", "\n", t)
    t = htmllib.unescape(re.sub(r"<[^>]+>", "", t))
    lines = [re.sub(r"[ \t\u3000]+", " ", l).strip() for l in t.split("\n")]
    return "\n".join(l for l in lines if l)[:limit]

def meta_description(html_text):
    for rx in (r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)',
               r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:description',
               r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)'):
        m = re.search(rx, html_text, re.I)
        if m:
            return re.sub(r"\s+", " ", htmllib.unescape(m.group(1))).strip()[:220]
    return ""

def call_claude(api_key, prompt, timeout=90):
    base = os.environ.get("ANTHROPIC_BASE_URL", "https://api.anthropic.com").rstrip("/")
    body = json.dumps({"model": os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5"), "max_tokens": 3000,
                       "system": SYSTEM_PROMPT, "messages": [{"role": "user", "content": prompt}]}).encode("utf-8")
    req = urllib.request.Request(base + "/v1/messages", data=body, method="POST",
                                 headers={"content-type": "application/json", "x-api-key": api_key,
                                          "anthropic-version": "2023-06-01"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.loads(r.read().decode("utf-8"))
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")

def parse_model_json(text):
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t)
    m = re.search(r"\{.*\}", t, re.S)
    return json.loads(m.group(0) if m else t)

def _str(v, n=3000):
    return v.strip()[:n] if isinstance(v, str) else ""

def apply_summary(it, d):
    """モデルの出力を検証して、事例の項目に反映する（不足・不正な値は無視）。"""
    if not _str(d.get("title"), 80) and not any(_str(d.get(k)) for k in SEC_KEYS):
        raise ValueError("整理結果が空でした")
    for k in ("co", "svc"):
        if not it.get(k) and _str(d.get(k), 60):
            it[k] = _str(d[k], 60)
    it["title"] = _str(d.get("title"), 80) or it.get("title", "")
    it["sub"] = _str(d.get("sub"), 120)
    if d.get("cat") in CATS:
        it["cat"] = d["cat"]
    try:
        n = int(d.get("n") or 0)
        if n > 0 and not it.get("n"):
            it["n"] = n
    except (TypeError, ValueError):
        pass
    if _str(d.get("nt"), 40) and not it.get("nt"):
        it["nt"] = _str(d["nt"], 40)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", _str(d.get("base"), 10) or ""):
        it["base"] = d["base"]
        if d.get("baseType") in ("発生", "検知", "公表"):
            it["baseType"] = d["baseType"]
    facts = []
    for f in (d.get("facts") or [])[:4]:
        if isinstance(f, (list, tuple)) and len(f) == 2 and _str(f[0], 20) and _str(f[1], 60):
            facts.append(_str(f[0], 20).replace("|", "/").replace("=", "＝") + "=" + _str(f[1], 60).replace("|", "/"))
    it["facts"] = "|".join(facts)
    for k in SEC_KEYS:
        it[k] = _str(d.get(k))
    it["sources"] = "元記事|" + it["url"]
    return it

def needs_enrich(it, have_key, today):
    if not it.get("url") or it.get("title"):
        return False
    res, at = it.get("enrichResult"), it.get("enrichedAt")
    if res == "ok":
        return False
    if res == "meta":
        return have_key            # キーが使えるようになったら整理し直す（紹介文だけが入っている状態）
    if any(it.get(k) for k in SEC_KEYS):
        return False               # すでに手元で詳細が入っているものには触れない
    if res == "fail":
        return at != today         # 失敗は翌日以降に再試行
    return True

def enrich(it, api_key, log, today):
    """元記事を読み、詳細（7項目）を整理して it に書き込む。結果を it['enrichResult'] に残す。"""
    it["enrichedAt"] = today
    try:
        page = http_get(it["url"])
    except Exception as e:  # noqa
        it["enrichResult"] = "fail"
        log(f"  ! 元記事を読めませんでした: {it['url']} ({e})")
        return False
    desc = meta_description(page)
    if api_key:
        try:
            body = html_to_text(page)
            if len(body) < 80:
                raise ValueError("記事本文を取り出せませんでした")
            out = call_claude(api_key, USER_TEMPLATE.format(url=it["url"], title=it.get("summary", ""), body=body))
            apply_summary(it, parse_model_json(out))
            it["enrichResult"] = "ok"
            it["method"] = "AI"
            return True
        except Exception as e:  # noqa
            log(f"  ! 整理に失敗: {it['url']} ({e})")
    if desc:                        # キーが無い・失敗したとき：記事の紹介文（短い）だけを概要にする
        it["sec_summary"] = desc
        it["sources"] = "元記事|" + it["url"]
        it["enrichResult"] = "meta"
        return True
    it["enrichResult"] = "fail"
    return False

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--feed", required=True)
    ap.add_argument("--sources", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--enrich-limit", type=int, default=8, help="1回の実行で元記事から詳細を整理する最大件数（0で無効）")
    a = ap.parse_args(argv)
    log = lambda s: print(s, flush=True)

    try:
        feed = json.load(open(a.feed, encoding="utf-8"))
    except FileNotFoundError:
        feed = {"version": 1, "items": []}
    feed.setdefault("items", [])
    known = {norm_url(i.get("url")) for i in feed["items"] if i.get("url")}
    ids = {i.get("id") for i in feed["items"]}

    sources = [s for s in json.load(open(a.sources, encoding="utf-8"))["sources"] if s.get("enabled", True)]
    ok_sources, added = [], []
    for s in sources:
        log(f"■ {s['name']}")
        s["_skip"] = known
        try:
            got = collect(s, log)
        except Exception as e:  # noqa
            log(f"  ! 失敗: {e}")
            continue
        ok_sources.append(s["name"])
        n = 0
        for it in got:
            if norm_url(it["url"]) in known or it["id"] in ids:
                continue
            known.add(norm_url(it["url"])); ids.add(it["id"]); added.append(it); n += 1
        log(f"  新規 {n} 件")
    if sources and not ok_sources:
        log("すべての情報源で失敗しました")
        return 1
    log(f"合計 新規 {len(added)} 件（フィード全体 {len(feed['items']) + len(added)} 件）")
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    today = datetime.datetime.now(JST).date().isoformat()
    pool = feed["items"] + added
    targets = sorted([i for i in pool if needs_enrich(i, bool(api_key), today)], key=lambda i: i.get("pub", ""), reverse=True)
    if a.dry_run:
        for it in added[:20]:
            log(f"  + {it['pub']} [{it['cat']}] {it['co'] or '-'} / {it['summary'][:50]}")
        log(f"元記事から詳細を整理できる事例: {len(targets)} 件（APIキー: {'あり' if api_key else 'なし'}）")
        return 0
    enriched = 0
    if a.enrich_limit > 0 and targets:
        log(f"■ 元記事から詳細を整理（{'Claude API' if api_key else '紹介文のみ（ANTHROPIC_API_KEY なし）'}、最大 {a.enrich_limit} 件 / 対象 {len(targets)} 件）")
        for it in targets[: a.enrich_limit]:
            time.sleep(1.0 if not os.environ.get("FETCH_NO_SLEEP") else 0)
            ok = enrich(it, api_key, log, today)
            enriched += 1 if ok else 0
            log(f"  {'○' if ok else '×'} {it.get('co') or it['url'][:40]} → {it.get('enrichResult')}")
    if added or enriched or targets[: max(a.enrich_limit, 0)]:
        feed["items"].extend(added)
        feed["generatedAt"] = datetime.datetime.now(JST).isoformat(timespec="seconds")
        names = list(dict.fromkeys((feed.get("sources") or []) + ok_sources))
        feed["sources"] = names
        feed["version"] = 1
        with open(a.feed, "w", encoding="utf-8") as f:
            json.dump(feed, f, ensure_ascii=False, indent=1)
            f.write("\n")
    return 0

if __name__ == "__main__":
    sys.exit(main())
