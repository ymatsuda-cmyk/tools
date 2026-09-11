#!/usr/bin/env python3
"""Notionのweb記事DBを対象に、URL先の本文を抜き出してページ本文へ書き込む。

対象の指定方法:
  --status 空欄 再取得   状態が空欄／再取得のページをまとめて処理（既定）
  --page-id <ID>        特定の1ページだけを処理

処理の流れ:
  1. Notionから対象ページを取得
  2. 本文が空のページだけを処理（再取得／--force のときは既存本文を捨てて入れ直す）
  3. URL先のHTMLを取得し、本文らしいブロックを抜き出す
  4. ページ送りのリンクがあれば1ページ目から順にたどって追記する
  5. og:image などからサムネイルを決め、タイトルとあわせてNotionへ反映
  6. 1件でも更新できたら build_clipstock_json.py を呼び、
     data/clipstock/index.json を作り直す

環境変数:
  NOTION_TOKEN     Notion Integration Token（必須）
  WEB_ENV_FILE     環境変数を読み込むファイルのパス（既定: ~/.video_notion_sync.env）
  WEB_DB_ID        対象データベースID（既定: web記事DB）

必要なパッケージ:
  pip install requests beautifulsoup4

使い方:
    python3 web_extract.py --dry-run
    python3 web_extract.py
    python3 web_extract.py --limit 1
    python3 web_extract.py --page-id 4130e7a535dc80000000000000000000
    python3 web_extract.py --page-id <ID> --force      # 本文があっても入れ直す
"""
import argparse
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests

JST = timezone(timedelta(hours=9))
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[2]

# ---------------------------------------------------------------- 環境変数

ENV_FILE = Path(os.environ.get("WEB_ENV_FILE", str(Path.home() / ".video_notion_sync.env")))


def load_env():
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env()

NOTION_API = "https://api.notion.com/v1"
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
WEB_DB_ID = os.environ.get("WEB_DB_ID", "4130e7a535dc83509c9a01cd6ac0a6a7")

# Notion側のカラム名。build_clipstock_json.py の PROP_* と一致させること
PROP_TITLE = "タイトル"
PROP_URL = "URL"
PROP_THUMB = "サムネイル"
PROP_STATUS = "状態"
PROP_RAW_COUNT = "原文文字数"

STATUS_EMPTY = "空欄"
STATUS_RETRY = "再取得"

CLIPSTOCK_BUILDER = SCRIPT_DIR.parent / "video" / "build_clipstock_json.py"
CLIPSTOCK_OUT_DIR = REPO_ROOT / "data" / "clipstock"

BROWSER_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "ja,en;q=0.8",
}
MAX_HTML_BYTES = 5 * 1024 * 1024
NOTION_TEXT_LIMIT = 2000
NOTION_URL_LIMIT = 2000

# ---------------------------------------------------------------- Notion


def notion_headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }


def get_valid_status_options():
    """DBの「状態」プロパティに定義済みの選択肢一覧を取得する。"""
    resp = requests.get(f"{NOTION_API}/databases/{WEB_DB_ID}",
                        headers=notion_headers(), timeout=60)
    if resp.status_code != 200:
        print(f"    ⚠️ DBスキーマ取得失敗: {resp.status_code} {resp.text[:200]}")
        return None
    prop = resp.json().get("properties", {}).get(PROP_STATUS, {})
    return {o["name"] for o in prop.get("select", {}).get("options", [])}


def build_status_filter(statuses):
    """状態フィルタを組み立てる。'空欄' は is_empty に変換する。

    DBに存在しない選択肢を equals に渡すとNotion APIが400を返すため、
    定義済みの選択肢だけを条件に残す。
    """
    valid_options = get_valid_status_options()
    conds = []
    for s in statuses:
        if s == STATUS_EMPTY:
            conds.append({"property": PROP_STATUS, "select": {"is_empty": True}})
            continue
        if valid_options is not None and s not in valid_options:
            print(f"    ⚠️ 状態「{s}」はこのDBに存在しないため無視します"
                  f"（利用可能: {', '.join(sorted(valid_options))}）")
            continue
        conds.append({"property": PROP_STATUS, "select": {"equals": s}})
    if not conds:
        return None
    return conds[0] if len(conds) == 1 else {"or": conds}


def query_pages_by_status(statuses):
    filter_ = build_status_filter(statuses)
    if filter_ is None:
        print("    ⚠️ 有効な状態フィルタが無いため検索を中止します。")
        return []
    pages, has_more, cursor = [], True, None
    while has_more:
        payload = {"page_size": 100, "filter": filter_}
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(f"{NOTION_API}/databases/{WEB_DB_ID}/query",
                             headers=notion_headers(), json=payload, timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ Notion検索失敗: {resp.status_code} {resp.text[:300]}")
            break
        data = resp.json()
        pages.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return pages


def fetch_page(page_id):
    resp = requests.get(f"{NOTION_API}/pages/{page_id}", headers=notion_headers(), timeout=60)
    if resp.status_code != 200:
        print(f"  ❌ ページ取得失敗: {resp.status_code} {resp.text[:300]}")
        return None
    return resp.json()


def page_title(page):
    """title型のプロパティを読む。DBごとに名前が違うので型でも探す。"""
    props = page.get("properties", {})
    items = (props.get(PROP_TITLE) or {}).get("title")
    if items is None:
        for value in props.values():
            if isinstance(value, dict) and value.get("type") == "title":
                items = value.get("title")
                break
    return "".join(p.get("plain_text", "") for p in (items or [])) or "(無題)"


def page_status(page):
    sel = (page.get("properties", {}).get(PROP_STATUS) or {}).get("select")
    return (sel or {}).get("name") or ""


def page_url_value(page):
    """URLプロパティ（url型）を読む。空ならタイトルにURLが入っているケースを拾う。"""
    url = ((page.get("properties", {}).get(PROP_URL)) or {}).get("url") or ""
    if url:
        return url.strip()
    title = page_title(page)
    return title.strip() if title.startswith("http") else ""


def get_all_children(block_id):
    children, has_more, cursor = [], True, None
    while has_more:
        url = f"{NOTION_API}/blocks/{block_id}/children?page_size=100"
        if cursor:
            url += f"&start_cursor={cursor}"
        resp = requests.get(url, headers=notion_headers(), timeout=60)
        if resp.status_code != 200:
            print(f"    ⚠️ ブロック取得失敗: {resp.status_code} {resp.text[:200]}")
            break
        data = resp.json()
        children.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        cursor = data.get("next_cursor")
    return children


# ブックマークや埋め込みは「本文」とみなさない（URLを貼っただけの状態）
NON_BODY_TYPES = ("bookmark", "embed", "video", "link_preview", "unsupported", "child_database")


def has_body(children):
    for block in children:
        btype = block.get("type")
        if btype in NON_BODY_TYPES:
            continue
        rich = (block.get(btype) or {}).get("rich_text")
        if rich is None or "".join(p.get("plain_text", "") for p in rich).strip():
            return True
    return False


def archive_body_children(page_id, children):
    """既存の本文をアーカイブする。ブックマークや埋め込みは残す。"""
    ok = total = 0
    for block in children:
        if block.get("type") in NON_BODY_TYPES:
            continue
        total += 1
        resp = requests.patch(f"{NOTION_API}/blocks/{block['id']}",
                              headers=notion_headers(), json={"archived": True}, timeout=60)
        if resp.status_code == 200:
            ok += 1
        else:
            print(f"    ⚠️ ブロック削除失敗: {resp.status_code} {resp.text[:200]}")
    return ok, total


def rich_text(content):
    return [{"type": "text", "text": {"content": content}}]


BLOCK_TYPE_BY_KIND = {
    "h2": "heading_2",
    "h3": "heading_3",
    "li": "bulleted_list_item",
    "quote": "quote",
    "p": "paragraph",
}


def blocks_from_sections(sections):
    """(種類, テキスト) の並びをNotionブロックへ変換する。長文は2000字ごとに割る。"""
    blocks = []
    for kind, text in sections:
        btype = BLOCK_TYPE_BY_KIND.get(kind, "paragraph")
        rest = text
        while rest:
            chunk, rest = rest[:NOTION_TEXT_LIMIT], rest[NOTION_TEXT_LIMIT:]
            blocks.append({"object": "block", "type": btype, btype: {"rich_text": rich_text(chunk)}})
            btype = "paragraph"  # 続きは見出しではなく本文として出す
    return blocks or [{"object": "block", "type": "paragraph",
                       "paragraph": {"rich_text": rich_text("（本文を抽出できませんでした）")}}]


def append_blocks(page_id, blocks):
    """100件ずつに分けて本文を追記する。"""
    remaining = blocks
    while remaining:
        batch, remaining = remaining[:100], remaining[100:]
        resp = requests.patch(f"{NOTION_API}/blocks/{page_id}/children",
                              headers=notion_headers(), json={"children": batch}, timeout=60)
        if resp.status_code != 200:
            print(f"    ❌ 本文追記失敗: {resp.status_code} {resp.text[:300]}")
            return False
    return True


def build_body_blocks(articles):
    """取得した各ページを、見出し＋取得元の注記つきで並べる。"""
    ts = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    blocks = []
    for i, art in enumerate(articles, 1):
        label = art["title"] or art["url"]
        if len(articles) > 1:
            label = f"{label}（{i}/{len(articles)}ページ目）"
        blocks.append({"object": "block", "type": "heading_2",
                       "heading_2": {"rich_text": rich_text(f"🌐 {label}"[:NOTION_TEXT_LIMIT])}})
        blocks.append({"object": "block", "type": "callout", "callout": {
            "icon": {"type": "emoji", "emoji": "🔗"},
            "rich_text": rich_text(f"取得: {ts}\n{art['url']}"[:NOTION_TEXT_LIMIT])}})
        blocks.extend(blocks_from_sections(art["sections"]))
    return blocks


def update_page_props(page_id, *, title=None, url=None, thumbnail=None,
                      char_count=None, status=None):
    props = {}
    if title:
        props[PROP_TITLE] = {"title": [{"text": {"content": title[:NOTION_TEXT_LIMIT]}}]}
    if url:
        props[PROP_URL] = {"url": url[:NOTION_URL_LIMIT]}
    if thumbnail:
        props[PROP_THUMB] = {"url": thumbnail[:NOTION_URL_LIMIT]}
    if char_count is not None:
        props[PROP_RAW_COUNT] = {"number": char_count}
    if status:
        props[PROP_STATUS] = {"select": {"name": status}}
    if not props:
        return True
    resp = requests.patch(f"{NOTION_API}/pages/{page_id}",
                          headers=notion_headers(), json={"properties": props}, timeout=60)
    if resp.status_code != 200:
        print(f"    ⚠️ プロパティ更新失敗: {resp.status_code} {resp.text[:300]}")
        return False
    return True


# ---------------------------------------------------------------- HTML取得

def fetch_html(url):
    """HTMLを取得して (本文, 最終URL) を返す。取れなければ (None, None)。"""
    if not re.match(r"^https?://", url or "", re.I):
        return None, None
    try:
        resp = requests.get(url, headers=BROWSER_HEADERS, timeout=60,
                            allow_redirects=True, stream=True)
    except requests.RequestException as e:
        print(f"    ⚠️ 取得失敗: {type(e).__name__}: {e}")
        return None, None
    with resp:
        if resp.status_code != 200:
            print(f"    ⚠️ HTTP {resp.status_code}: {url[:100]}")
            return None, None
        ctype = resp.headers.get("Content-Type", "")
        if ctype and "html" not in ctype.lower() and "xml" not in ctype.lower():
            print(f"    ⚠️ HTMLではないためスキップ: {ctype}")
            return None, None
        raw = b""
        for chunk in resp.iter_content(65536):
            raw += chunk
            if len(raw) >= MAX_HTML_BYTES:
                break
        enc = resp.encoding if "charset=" in ctype.lower() else None
        if not enc:
            m = re.search(rb'charset=["\']?([\w-]+)', raw[:4096], re.I)
            enc = m.group(1).decode("ascii", "ignore") if m else (resp.apparent_encoding or "utf-8")
        return raw.decode(enc, errors="replace"), resp.url or url


def make_soup(html):
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print("    ❌ beautifulsoup4 が未導入です（pip install beautifulsoup4）")
        return None
    return BeautifulSoup(html, "html.parser")


# ---------------------------------------------------------------- 本文の抽出

NOISE_TAGS = ("script", "style", "noscript", "template", "svg", "iframe", "form",
              "button", "nav", "header", "footer", "aside")
NOISE_WORDS = re.compile(
    r"(comment|share|social|sidebar|related|recommend|ranking|breadcrumb|pager|pagination|"
    r"advert|\bads?\b|banner|subscribe|newsletter|footer|header|global-?nav|menu)", re.I)
CONTENT_HINTS = ("article", "main", "[role=main]", ".entry-content", ".post-content",
                 ".article-body", ".articleBody", ".post-body", ".content-body",
                 "#content", ".content", "#main")
TEXT_TAGS = ("h1", "h2", "h3", "h4", "h5", "p", "li", "blockquote", "pre", "dd")
KIND_BY_TAG = {"h1": "h2", "h2": "h2", "h3": "h3", "h4": "h3", "h5": "h3",
               "li": "li", "dd": "li", "blockquote": "quote"}


def strip_noise(soup):
    for tag in soup.find_all(NOISE_TAGS):
        tag.decompose()
    for tag in soup.find_all(attrs={"class": NOISE_WORDS}):
        tag.decompose()
    for tag in soup.find_all(attrs={"id": NOISE_WORDS}):
        tag.decompose()


def text_score(el):
    return sum(len(t.get_text(strip=True)) for t in el.find_all(("p", "li")))


def pick_container(soup):
    """本文らしい要素を選ぶ。段落の文字数がいちばん多いものを採る。"""
    best, best_score = None, 0
    for sel in CONTENT_HINTS:
        for el in soup.select(sel):
            score = text_score(el)
            if score > best_score:
                best, best_score = el, score
    body = soup.body or soup
    return best if best is not None and best_score >= 200 else body


def extract_sections(container):
    """本文を (種類, テキスト) の並びで取り出す。入れ子の重複は親側を捨てる。"""
    sections, seen = [], set()
    for el in container.find_all(TEXT_TAGS):
        if el.find(TEXT_TAGS):  # 見出しや段落を内包する箱は中身側で拾う
            continue
        text = re.sub(r"[ \t\u3000]+", " ", el.get_text(" ", strip=True)).strip()
        if len(text) < 2:
            continue
        key = (el.name, text)
        if key in seen:
            continue
        seen.add(key)
        sections.append((KIND_BY_TAG.get(el.name, "p"), text))
    return sections


def meta_content(soup, *selectors):
    for sel in selectors:
        el = soup.select_one(sel)
        if el:
            value = (el.get("content") or el.get("href") or "").strip()
            if value:
                return value
    return ""


def clean_title(title, site_name):
    """「記事名 | サイト名」のような末尾のサイト名を落とす。"""
    t = re.sub(r"\s+", " ", title or "").strip()
    if site_name and t.endswith(site_name) and len(t) > len(site_name) + 3:
        t = t[: -len(site_name)].rstrip(" |-–—｜·»>/:").strip()
    return t


def extract_title(soup):
    site = meta_content(soup, 'meta[property="og:site_name"]')
    raw = meta_content(soup, 'meta[property="og:title"]', 'meta[name="twitter:title"]')
    if not raw:
        h1 = soup.find("h1")
        raw = h1.get_text(" ", strip=True) if h1 else ""
    if not raw and soup.title:
        raw = soup.title.get_text(" ", strip=True)
    return clean_title(raw, site)


def extract_thumbnail(soup, container, base_url):
    """タイトル画像を選ぶ。OGP優先、無ければ本文中でいちばん大きい画像。"""
    url = meta_content(soup, 'meta[property="og:image"]', 'meta[property="og:image:url"]',
                       'meta[name="twitter:image"]', 'meta[name="thumbnail"]',
                       'link[rel="image_src"]')
    if url:
        return absolute_image(url, base_url)

    best, best_area, first = None, 0, None
    for img in (container or soup).find_all("img"):
        src = (img.get("src") or img.get("data-src") or img.get("data-original") or "").strip()
        if not src or src.startswith("data:"):
            continue
        area = int_attr(img, "width") * int_attr(img, "height")
        first = first or src
        if area > best_area:
            best, best_area = src, area
    return absolute_image(best or first, base_url)


def int_attr(el, name):
    m = re.search(r"\d+", str(el.get(name) or ""))
    return int(m.group()) if m else 0


def absolute_image(src, base_url):
    if not src:
        return ""
    url = urljoin(base_url, src)
    return url if re.match(r"^https?://", url, re.I) else ""


# ---------------------------------------------------------------- ページ送り

PAGE_QUERY_KEYS = ("page", "p", "pg", "pagenum", "page_no")
PAGE_PATH_RE = re.compile(r"(?i)/(?:page|p)/(\d+)(/?)$")


def page_number_of(url):
    """URLからページ番号を読む。見つからなければ1。"""
    parts = urlsplit(url)
    for key, value in parse_qsl(parts.query):
        if key.lower() in PAGE_QUERY_KEYS and value.isdigit():
            return int(value)
    m = PAGE_PATH_RE.search(parts.path)
    return int(m.group(1)) if m else 1


def page_shape(url):
    """ページ番号を除いたURL。同じ連載のページかどうかの判定に使う。"""
    parts = urlsplit(url)
    query = [(k, v) for k, v in parse_qsl(parts.query)
             if not (k.lower() in PAGE_QUERY_KEYS and v.isdigit())]
    path = PAGE_PATH_RE.sub("/", parts.path)
    return urlunsplit((parts.scheme, parts.netloc, path.rstrip("/"), urlencode(query), ""))


def next_page_urls(current_url, soup):
    """同じ記事のページ送りリンクを、ページ番号の小さい順で返す。"""
    host = urlparse(current_url).netloc
    shape = page_shape(current_url)
    found = {}

    rel_next = soup.select_one('link[rel="next"], a[rel="next"]')
    if rel_next:
        href = (rel_next.get("href") or "").strip()
        if href:
            url = urljoin(current_url, href)
            found[page_number_of(url)] = url

    for a in soup.find_all("a", href=True):
        url = urljoin(current_url, a["href"].strip())
        if not re.match(r"^https?://", url, re.I) or urlparse(url).netloc != host:
            continue
        if page_shape(url) != shape:
            continue
        num = page_number_of(url)
        if num > 1:
            found.setdefault(num, url.split("#")[0])
    return [found[n] for n in sorted(found)]


def first_page_url(url):
    """連載の途中のURLを渡されたときに1ページ目へ戻す。"""
    if page_number_of(url) <= 1:
        return url
    parts = urlsplit(url)
    query = [(k, v) for k, v in parse_qsl(parts.query)
             if not (k.lower() in PAGE_QUERY_KEYS and v.isdigit())]
    path = PAGE_PATH_RE.sub("/", parts.path)
    return urlunsplit((parts.scheme, parts.netloc, path, urlencode(query), ""))


def collect_articles(start_url, *, max_pages, delay):
    """1ページ目から順にたどり、各ページの本文を集める。"""
    articles, seen = [], set()
    queue = [first_page_url(start_url)]
    if queue[0] != start_url:
        print(f"    → 1ページ目から取得します: {queue[0][:100]}")

    while queue and len(articles) < max_pages:
        url = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        if articles:
            time.sleep(delay)
        html, final_url = fetch_html(url)
        if not html:
            if not articles and url != start_url:
                queue.append(start_url)  # 1ページ目が無いサイトは元のURLで取り直す
            continue
        seen.add(final_url)
        soup = make_soup(html)
        if soup is None:
            return []

        title = extract_title(soup)
        strip_noise(soup)
        container = pick_container(soup)
        sections = extract_sections(container)
        if not sections:
            print(f"    ⚠️ 本文を抽出できませんでした: {final_url[:100]}")
            continue
        articles.append({
            "url": final_url,
            "title": title,
            "sections": sections,
            "thumbnail": extract_thumbnail(soup, container, final_url),
        })
        print(f"    ✅ {len(articles)}ページ目: {sum(len(t) for _, t in sections)}文字")

        for nxt in next_page_urls(final_url, soup):
            if nxt not in seen and nxt not in queue:
                queue.append(nxt)
    return articles


# ---------------------------------------------------------------- 本処理

def process_page(page, *, set_status_name, force, max_pages, delay):
    title = page_title(page)
    page_id = page.get("id")
    url = page_url_value(page)
    status = page_status(page)

    print(f"  対象: {title[:60]}")
    print(f"        page_id={page_id}")

    if not url:
        print("  ❌ URLが空のためスキップ")
        return False

    children = get_all_children(page_id)
    body_exists = has_body(children)
    refill = force or status == STATUS_RETRY
    if body_exists and not refill:
        print("  ⏭️  本文がすでにあるためスキップ（入れ直すなら --force か状態を再取得に）")
        return False

    articles = collect_articles(url, max_pages=max_pages, delay=delay)
    if not articles:
        print("  ❌ 本文を取得できませんでした。スキップ")
        return False

    total_chars = sum(len(t) for art in articles for _, t in art["sections"])
    print(f"  ✅ {len(articles)}ページ / 合計 {total_chars}文字")

    if body_exists:
        print("    → 既存の本文を削除中...")
        ok, total = archive_body_children(page_id, children)
        print(f"    ✅ {ok}/{total} ブロックを削除")

    if not append_blocks(page_id, build_body_blocks(articles)):
        return False
    print("  ✅ 本文に追記")

    new_title = articles[0]["title"] or None
    thumbnail = next((a["thumbnail"] for a in articles if a["thumbnail"]), None)
    canonical = articles[0]["url"]
    update_page_props(
        page_id,
        title=new_title,
        url=canonical if canonical != url else None,
        thumbnail=thumbnail,
        char_count=total_chars,
        status=set_status_name,
    )
    if new_title:
        print(f"  ✅ タイトルを更新: {new_title[:60]}")
    if thumbnail:
        print(f"  ✅ サムネイルを登録: {thumbnail[:80]}")
    if set_status_name:
        print(f"  ✅ 状態を「{set_status_name}」に更新")
    return True


def rebuild_clipstock_index():
    """一覧用の index.json を build_clipstock_json.py で作り直す。"""
    if not CLIPSTOCK_BUILDER.exists():
        print(f"⚠️ {CLIPSTOCK_BUILDER.name} が見つからないため index.json は更新しません")
        return
    print(f"\nindex.json を更新中: {CLIPSTOCK_OUT_DIR}")
    try:
        result = subprocess.run(
            [sys.executable, str(CLIPSTOCK_BUILDER), "--out", str(CLIPSTOCK_OUT_DIR)],
            capture_output=True, text=True, timeout=900)
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"⚠️ index.json の更新に失敗: {type(e).__name__}: {e}")
        return
    out = (result.stdout or "").strip()
    if out:
        print(out)
    if result.returncode != 0:
        print(f"⚠️ index.json の更新に失敗: {(result.stderr or '').strip()[-300:]}")
    else:
        print(f"✅ index.json を更新しました: {CLIPSTOCK_OUT_DIR / 'index.json'}")


def main():
    ap = argparse.ArgumentParser(description="Notion web記事DBの本文抽出")
    ap.add_argument("--page-id", help="対象のNotionページID（--statusより優先）")
    ap.add_argument("--status", nargs="+", default=[STATUS_EMPTY, STATUS_RETRY],
                    help=f"対象とする状態（既定: {STATUS_EMPTY} {STATUS_RETRY}）")
    ap.add_argument("--set-status", default="完了",
                    help="処理完了後に設定する状態（既定: 完了。空文字で更新しない）")
    ap.add_argument("--limit", type=int, help="処理する件数の上限")
    ap.add_argument("--force", action="store_true", help="本文があっても取り直して入れ替える")
    ap.add_argument("--max-pages", type=int, default=20, help="たどるページ送りの上限（既定: 20）")
    ap.add_argument("--delay", type=float, default=1.0, help="ページ取得の間隔・秒（既定: 1.0）")
    ap.add_argument("--no-rebuild", action="store_true", help="index.json を作り直さない")
    ap.add_argument("--dry-run", action="store_true", help="対象一覧を表示するだけ")
    args = ap.parse_args()

    if not NOTION_TOKEN:
        print("❌ NOTION_TOKEN が設定されていません")
        return 1

    now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S JST")
    print(f"\n{'='*60}\nweb記事DB→本文抽出開始: {now}\n{'='*60}")
    print(f"実行中のPython: {sys.executable} ({sys.version.split()[0]})")

    if args.page_id:
        page = fetch_page(args.page_id)
        targets = [page] if page else []
    else:
        targets = query_pages_by_status(args.status)

    if not targets:
        print("対象ページが見つかりませんでした。")
        return 0

    if args.limit:
        targets = targets[:args.limit]

    print(f"\n対象: {len(targets)}件")
    for p in targets:
        print(f"  - [{page_status(p) or '(空欄)'}] {page_title(p)[:60]}")

    if args.dry_run:
        print("\n--dry-run のため処理は行いません。")
        return 0

    ok_count = 0
    for i, page in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}]")
        try:
            if process_page(page, set_status_name=args.set_status or None,
                            force=args.force or bool(args.page_id),
                            max_pages=args.max_pages, delay=args.delay):
                ok_count += 1
        except Exception as e:  # noqa: BLE001
            print(f"  ❌ 例外: {type(e).__name__}: {e}")

    print(f"\n{'='*60}\n完了: {ok_count}/{len(targets)}件  "
          f"{datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S JST')}\n{'='*60}\n")

    if ok_count and not args.no_rebuild:
        rebuild_clipstock_index()
    return 0


if __name__ == "__main__":
    sys.exit(main())
