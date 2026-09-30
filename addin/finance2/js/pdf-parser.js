/* ============================================================
 * pdf-parser.js — 合計残高試算表PDFの読み取り（AI不使用）
 * ------------------------------------------------------------
 * pdf.js のテキストレイヤーから文字と座標を取り出し、
 *   y座標で行をまとめる → x座標で「コード／科目名／数値4列」に振り分ける
 * 列位置は各ページの見出し「繰越残高 借方 貸方 残高」の右端を基準にする。
 * 公開: window.TBParser.parse(arrayBuffer, fileName) → Promise<result>
 * ============================================================ */
(function (global) {
  "use strict";
  const PDFJS_VER = "3.11.174";
  const PDFJS_BASE = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/" + PDFJS_VER;
  const CMAP_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@" + PDFJS_VER + "/cmaps/";

  // 集計行とみなす科目名（コードがなく、名前が「計」で終わらないもの）
  const SUBTOTAL_NAMES = new Set([
    "純売上高", "売上原価", "売上総利益", "売上総損失", "販売費及び一般管理費",
    "営業利益", "営業損失", "営業外収益", "営業外費用", "経常利益", "経常損失",
    "特別利益", "特別損失", "税引前当期純利益", "税引前当期純損失",
    "当期純利益", "当期純損失", "製造原価", "当期製品製造原価"
  ]);
  const COL_NAMES = ["繰越残高", "借方", "貸方", "残高"];

  let loading = null;
  function loadPdfjs() {
    if (global.pdfjsLib) return Promise.resolve(global.pdfjsLib);
    if (!loading) {
      const base = global.FIN2_PDFJS_BASE || PDFJS_BASE;
      loading = new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = base + "/pdf.min.js";
        s.onload = () => {
          global.pdfjsLib.GlobalWorkerOptions.workerSrc = base + "/pdf.worker.min.js";
          resolve(global.pdfjsLib);
        };
        s.onerror = () => { loading = null; reject(new Error("PDF読み取りライブラリを読み込めませんでした。通信環境を確認してください。")); };
        document.head.appendChild(s);
      });
    }
    return loading;
  }

  const normName = (s) => String(s || "").replace(/[\s\u3000]/g, "");
  const NUM_RE = /^[-−△▲]?\(?[\d,]+\)?$/;
  function toNum(s) {
    let t = String(s).trim();
    let neg = false;
    if (/^[-−△▲]/.test(t)) { neg = true; t = t.slice(1); }
    if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
    const n = Number(t.replace(/,/g, ""));
    return neg ? -n : n;
  }

  // 空白を含むテキスト片を、文字数比でx座標を割り振りながら分割する
  function splitItem(it) {
    const parts = it.s.split(/[ \u3000]+/).filter(Boolean);
    if (parts.length <= 1) return [it];
    const total = it.s.length || 1, cw = it.w / total;
    const out = []; let pos = 0;
    for (const p of parts) {
      const idx = it.s.indexOf(p, pos);
      out.push({ s: p, x: it.x + idx * cw, x1: it.x + (idx + p.length) * cw, y: it.y });
      pos = idx + p.length;
    }
    return out;
  }

  function groupRows(items) {
    const sorted = items.slice().sort((a, b) => b.y - a.y);
    const rows = [];
    for (const it of sorted) {
      const last = rows[rows.length - 1];
      if (last && Math.abs(last.y - it.y) <= 2.5) last.items.push(it);
      else rows.push({ y: it.y, items: [it] });
    }
    rows.forEach(r => r.items.sort((a, b) => a.x - b.x));
    return rows;
  }

  function wareki(era, y) {
    const n = y === "元" ? 1 : Number(y);
    return { "令和": 2018, "平成": 1988 }[era] + n;
  }

  function detectPeriod(text) {
    const t = text.replace(/[\s\u3000]/g, "");
    const re = /自(令和|平成)?(\d{4}|\d{1,2}|元)年(\d{1,2})月(\d{1,2})日至(令和|平成)?(\d{4}|\d{1,2}|元)年(\d{1,2})月(\d{1,2})日/;
    const m = t.match(re);
    if (!m) return null;
    const fy = m[1] ? wareki(m[1], m[2]) : Number(m[2]);
    const ty = m[5] ? wareki(m[5], m[6]) : Number(m[6]);
    return {
      from: { y: fy, m: Number(m[3]), d: Number(m[4]) },
      to: { y: ty, m: Number(m[7]), d: Number(m[8]) }
    };
  }

  function sheetOf(text) {
    if (/貸借対照表/.test(text)) return "BS";
    if (/損益計算書/.test(text)) return "PL";
    if (/製造原価/.test(text)) return "MC";
    return null;
  }

  function classify(code, nameN) {
    if (code) return "明細";
    if (/(計|合計)$/.test(nameN) || SUBTOTAL_NAMES.has(nameN)) return "集計";
    return "明細";
  }

  async function parse(buffer, fileName) {
    const res = {
      fileName, period: null, company: "", taxMode: "", rows: [],
      errors: [], warnings: [], checks: {}
    };
    let lib;
    try { lib = await loadPdfjs(); } catch (e) { res.errors.push(e.message); return res; }
    let doc;
    try {
      doc = await lib.getDocument({ data: new Uint8Array(buffer), cMapUrl: global.FIN2_CMAP_URL || CMAP_URL, cMapPacked: true }).promise;
    } catch (e) {
      res.errors.push("PDFを開けませんでした（" + (e && e.message || e) + "）");
      return res;
    }
    const orderBy = { BS: 0, PL: 0, MC: 0 };
    const seen = new Set();
    let periodInfo = null;
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      let items = tc.items
        .filter(i => i.str && i.str.trim())
        .map(i => ({ s: i.str.trim(), x: i.transform[4], w: i.width, x1: i.transform[4] + i.width, y: i.transform[5] }));
      const pageText = items.map(i => i.s).join("");
      if (!pageText.trim()) continue;
      const sheet = sheetOf(pageText);
      if (!periodInfo) periodInfo = detectPeriod(pageText);
      if (!res.company) {
        const c = items.find(i => /(株式会社|有限会社|合同会社|合資会社|一般社団法人|医療法人)/.test(i.s));
        if (c) res.company = c.s.replace(/出力者.*$/, "").trim();
      }
      if (!res.taxMode) { const m = pageText.match(/【(税込|税抜)】/); if (m) res.taxMode = m[1]; }
      if (!sheet) { res.warnings.push(p + "ページ目は貸借対照表・損益計算書ではないため読み飛ばしました"); continue; }

      items = items.flatMap(splitItem);
      const rows = groupRows(items);
      // 見出し行を探す
      const hdr = rows.find(r => COL_NAMES.every(n => r.items.some(i => i.s === n)));
      if (!hdr) { res.warnings.push(p + "ページ目で列見出し（繰越残高・借方・貸方・残高）が見つかりませんでした"); continue; }
      const colR = COL_NAMES.map(n => hdr.items.find(i => i.s === n).x1);
      const numLeft = hdr.items.find(i => i.s === "繰越残高").x - 30;
      const nameHdr = hdr.items.find(i => i.s === "科目名");
      const codeRight = nameHdr ? nameHdr.x - 2 : 80;

      for (const r of rows) {
        if (r.y >= hdr.y - 1) continue;
        const nums = r.items.filter(i => i.x >= numLeft && NUM_RE.test(i.s));
        if (nums.length === 0) continue;
        const vals = [null, null, null, null];
        for (const n of nums) {
          let best = 0, bd = Infinity;
          colR.forEach((cx, k) => { const d = Math.abs(cx - n.x1); if (d < bd) { bd = d; best = k; } });
          vals[best] = toNum(n.s);
        }
        if (vals.some(v => v === null)) continue;
        const codeItem = r.items.find(i => i.x1 <= codeRight + 1 && /^[0-9A-Za-z-]+$/.test(i.s));
        const code = codeItem ? codeItem.s : "";
        const name = r.items
          .filter(i => i !== codeItem && i.x < numLeft && i.x1 > codeRight - 1)
          .map(i => i.s).join(" ").replace(/\s+/g, " ").trim();
        if (!name || /^出力者/.test(name)) continue;
        const nameN = normName(name);
        let key = sheet + ":" + (code || "#" + nameN);
        let dup = 2;
        while (seen.has(key)) key = sheet + ":" + (code || "#" + nameN) + "#" + dup++;
        seen.add(key);
        res.rows.push({
          sheet, order: ++orderBy[sheet], kind: classify(code, nameN), key, code, name,
          open: vals[0], dr: vals[1], cr: vals[2], bal: vals[3]
        });
      }
    }

    // ---- 期間 ----
    if (periodInfo) {
      const { from, to } = periodInfo;
      if (from.y !== to.y || from.m !== to.m) {
        res.errors.push(`期間が1か月ではありません（${from.y}/${from.m}/${from.d}〜${to.y}/${to.m}/${to.d}）。月次の試算表を出力してください。`);
      } else {
        res.period = to.y + "-" + String(to.m).padStart(2, "0");
      }
    }
    if (!res.rows.length) {
      res.errors.push("表の行を読み取れませんでした。文字情報のないPDF（スキャン画像など）は読み込めません。");
      return res;
    }
    if (!periodInfo) res.errors.push("期間（自〜至）を読み取れませんでした。");

    // ---- 検算 ----
    const bad = res.rows.filter(r => r.open + r.dr - r.cr !== r.bal && r.open - r.dr + r.cr !== r.bal);
    res.checks.formulaNG = bad.map(r => r.name);
    if (bad.length) res.warnings.push("繰越残高・借方・貸方から残高が計算できない行があります：" + bad.slice(0, 5).map(r => r.name).join("、"));
    const byName = (sheet, names) => res.rows.find(r => r.sheet === sheet && names.includes(normName(r.name)));
    const asset = byName("BS", ["資産合計", "資産の部合計"]);
    const le = byName("BS", ["負債・純資産合計", "負債純資産合計", "負債及び純資産合計", "負債・純資産の部合計"]);
    if (asset && le) {
      res.checks.balance = asset.bal === le.bal;
      res.checks.assetTotal = asset.bal;
      if (!res.checks.balance) res.errors.push(`資産合計（${asset.bal.toLocaleString()}）と負債・純資産合計（${le.bal.toLocaleString()}）が一致しません。`);
    } else {
      res.checks.balance = null;
      res.warnings.push("資産合計／負債・純資産合計の行が見つからないため、貸借の一致を確認できませんでした。");
    }
    if (!res.rows.some(r => r.sheet === "PL")) res.warnings.push("損益計算書のページがありません。");
    if (!res.rows.some(r => r.sheet === "BS")) res.warnings.push("貸借対照表のページがありません。");
    return res;
  }

  global.TBParser = { parse, loadPdfjs, normName, SUBTOTAL_NAMES };
})(window);
