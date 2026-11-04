/* ============================================================
 * 営業報告アドイン shinki.js（新規営業タブ）
 * ------------------------------------------------------------
 * 営業代行会社から毎月受領する「営業代行 月次報告書（.docx）」を
 * 読み取り、シート「新規営業」に月単位で蓄積・可視化する。
 *
 * シート「新規営業」（1行目=見出し、2行目以降=データ。3ブロック横並び）
 *  【月次実績】   A:対象月 B:委託先 C:委託先担当 D:架電数 E:アポ獲得
 *                 F:見込み G:折り返し待ち H:再架電予定 I:担当者不在
 *                 J:NG（不要） K:NG（時期尚早） L:不通 M:未対応・その他
 *                 N:翌月目標架電数
 *  【企業フォロー】P:対象月 Q:区分 R:会社名 S:担当者 T:接触日 U:状況メモ
 *                 V:連絡先 W:次回アクション日 X:次回アクション内容
 *                 Y:自社対応状況（手入力用。再取込でも同じ会社なら保持）
 *  【所感・傾向】 AA:対象月 AB:種別（NG傾向／所感／翌月計画） AC:見出し AD:内容
 *  （O列・Z列は区切りの空列）
 *
 * 同じ対象月を再取込した場合は、その月の行だけを置き換える（上書き）。
 * app.js のグローバル（esc / uiAlert / uiConfirm / groupedBarChart /
 * fiscalMonths / termOfDate / termLabel / toDate / toSerial / demoMode）を利用。
 * 名前衝突を避けるため IIFE で包み、HTML から呼ぶ関数だけ window に公開する。
 * ============================================================ */
(function () {
  "use strict";

  const SK_SHEET = "新規営業";
  const SK_MAX_ROWS = 2000;
  const JSZIP_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

  const MONTHLY_COLS = ["対象月", "委託先", "委託先担当", "架電数", "アポ獲得", "見込み", "折り返し待ち",
    "再架電予定", "担当者不在", "NG（不要）", "NG（時期尚早）", "不通", "未対応・その他", "翌月目標架電数"];
  const FOLLOW_COLS = ["対象月", "区分", "会社名", "担当者", "接触日", "状況メモ", "連絡先",
    "次回アクション日", "次回アクション内容", "自社対応状況"];
  const NOTE_COLS = ["対象月", "種別", "見出し", "内容"];
  const BLOCKS = {
    monthly: { first: "A", last: "N", cols: MONTHLY_COLS, color: "#0e7a5f" },
    follow:  { first: "P", last: "Y", cols: FOLLOW_COLS,  color: "#2c6e9b" },
    note:    { first: "AA", last: "AD", cols: NOTE_COLS,  color: "#7b4baf" },
  };
  /* 架電結果の内訳（表示順・色） */
  const BREAKDOWN = [
    { key: "prospect", label: "見込み",       color: "#0e7a5f" },
    { key: "callback", label: "折り返し待ち", color: "#d9a038" },
    { key: "recall",   label: "再架電予定",   color: "#2c6e9b" },
    { key: "absent",   label: "担当者不在",   color: "#8e5aa8" },
    { key: "ngNo",     label: "NG（不要）",   color: "#c0392b" },
    { key: "ngEarly",  label: "NG（時期尚早）", color: "#e8a49c" },
    { key: "noAnswer", label: "不通",         color: "#9aa5ad" },
    { key: "other",    label: "未対応・その他", color: "#cfd6db" },
  ];
  const MONTHLY_KEYS = ["month", "vendor", "vendorPerson", "calls", "appo", "prospect", "callback",
    "recall", "absent", "ngNo", "ngEarly", "noAnswer", "other", "targetNext"];

  let skMonthly = [];   // {month, vendor, ..., row}
  let skFollow = [];    // {month, kind, company, person, contact(Date), memo, mail, nextDate, nextAction, ownStatus, row}
  let skNotes = [];     // {month, kind, title, body, row}
  let skMonth = null;   // 表示中の対象月 "YYYY/MM"
  let skTerm = null;    // 推移グラフの期
  let skPending = null; // 取込プレビュー中の解析結果
  let skLoaded = false;

  /* ============================================================
     ユーティリティ
     ============================================================ */
  function colIndex(letter) {
    let n = 0;
    for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function colLetter(idx) {
    let s = "", n = idx + 1;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function num(v) {
    if (v === "" || v == null) return 0;
    const n = Number(String(v).replace(/[^\d.\-]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  /* 対象月の正規化：Excelが日付に変換していても "YYYY/MM" に戻す */
  function normMonth(v) {
    if (v === "" || v == null) return "";
    if (typeof v === "number") {
      const d = toDate(v);
      return d ? monthKeyYM(d.getFullYear(), d.getMonth() + 1) : "";
    }
    const m = String(v).match(/(\d{4})\D+(\d{1,2})/);
    return m ? monthKeyYM(+m[1], +m[2]) : String(v).trim();
  }
  function monthLabel(mk) {
    const m = String(mk).match(/(\d{4})\/(\d{2})/);
    return m ? `${+m[1]}年${+m[2]}月` : mk;
  }
  function prevMonth(mk) {
    const [y, m] = mk.split("/").map(Number);
    return m === 1 ? monthKeyYM(y - 1, 12) : monthKeyYM(y, m - 1);
  }
  /* 会社名の名寄せキー（法人格・空白を除去） */
  function companyKey(s) {
    return String(s || "").replace(/株式会社|有限会社|合同会社|（株）|\(株\)|（有）|\(有\)|\s|　/g, "");
  }
  function sum(rec) {
    return BREAKDOWN.reduce((a, b) => a + (rec[b.key] || 0), 0);
  }
  function md(d) { return d ? `${d.getMonth() + 1}/${d.getDate()}` : ""; }
  function today0() { const t = new Date(); t.setHours(0, 0, 0, 0); return t; }
  function isClosed(f) { return /完了|済|終了|クローズ|失注/.test(f.ownStatus || ""); }

  /* ============================================================
     読み込み（Excel／デモ）
     ============================================================ */
  async function skLoad() {
    if (demoMode || !window.Excel) { if (!skLoaded) skLoadDemo(); skLoaded = true; return; }
    try {
      await Excel.run(async ctx => {
        let sheet = ctx.workbook.worksheets.getItemOrNullObject(SK_SHEET);
        await ctx.sync();
        if (sheet.isNullObject) {
          sheet = ctx.workbook.worksheets.add(SK_SHEET);
          await ctx.sync();
        }
        await ensureHeaders(ctx, sheet);
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load("rowCount,rowIndex");
        await ctx.sync();
        const lastRow = used.isNullObject ? 1 : Math.min(used.rowIndex + used.rowCount, SK_MAX_ROWS + 1);
        if (lastRow < 2) { skMonthly = []; skFollow = []; skNotes = []; return; }
        const rng = sheet.getRange(`A2:AD${lastRow}`);
        rng.load("values");
        await ctx.sync();
        parseSheetValues(rng.values);
      });
      skLoaded = true;
    } catch (e) {
      console.error("新規営業シートの読み込みに失敗:", e);
      skMonthly = []; skFollow = []; skNotes = [];
    }
  }

  async function ensureHeaders(ctx, sheet) {
    const hdr = sheet.getRange("A1:AD1");
    hdr.load("values");
    await ctx.sync();
    const cur = hdr.values[0];
    for (const b of Object.values(BLOCKS)) {
      const s = colIndex(b.first);
      const ok = b.cols.every((h, i) => String(cur[s + i] || "").trim() === h);
      if (ok) continue;
      const r = sheet.getRange(`${b.first}1:${b.last}1`);
      r.values = [b.cols];
      r.format.fill.color = b.color;
      r.format.font.color = "#FFFFFF";
      r.format.font.bold = true;
    }
    await ctx.sync();
  }

  function parseSheetValues(values) {
    const mS = colIndex("A"), fS = colIndex("P"), nS = colIndex("AA");
    skMonthly = []; skFollow = []; skNotes = [];
    values.forEach((r, i) => {
      const row = i + 2;
      const mm = normMonth(r[mS]);
      if (mm) {
        const rec = { row };
        MONTHLY_KEYS.forEach((k, j) => {
          const v = r[mS + j];
          rec[k] = (j === 0) ? mm : (j <= 2 ? String(v || "") : num(v));
        });
        skMonthly.push(rec);
      }
      const fm = normMonth(r[fS]);
      if (fm && r[fS + 2]) {
        skFollow.push({
          row, month: fm, kind: String(r[fS + 1] || ""), company: String(r[fS + 2] || ""),
          person: String(r[fS + 3] || ""), contact: toDate(r[fS + 4]), memo: String(r[fS + 5] || ""),
          mail: String(r[fS + 6] || ""), nextDate: toDate(r[fS + 7]), nextAction: String(r[fS + 8] || ""),
          ownStatus: String(r[fS + 9] || ""),
        });
      }
      const nm = normMonth(r[nS]);
      if (nm && (r[nS + 2] || r[nS + 3])) {
        skNotes.push({ row, month: nm, kind: String(r[nS + 1] || ""), title: String(r[nS + 2] || ""), body: String(r[nS + 3] || "") });
      }
    });
  }

  /* ============================================================
     書き込み（対象月を置き換えて各ブロックを書き戻す）
     ============================================================ */
  function monthlyToRow(m) { return MONTHLY_KEYS.map(k => m[k] ?? ""); }
  function followToRow(f) {
    return [f.month, f.kind, f.company, f.person, f.contact ? toSerial(f.contact) : "", f.memo, f.mail,
      f.nextDate ? toSerial(f.nextDate) : "", f.nextAction, f.ownStatus || ""];
  }
  function noteToRow(n) { return [n.month, n.kind, n.title, n.body]; }

  function mergeParsed(parsed) {
    const mk = parsed.month;
    /* 自社対応状況（Y列）は手入力欄なので、同じ月・同じ会社の既存値を引き継ぐ */
    const keep = {};
    skFollow.filter(f => f.month === mk && f.ownStatus).forEach(f => { keep[companyKey(f.company)] = f.ownStatus; });
    parsed.follow.forEach(f => { if (!f.ownStatus && keep[companyKey(f.company)]) f.ownStatus = keep[companyKey(f.company)]; });

    const byMonth = (a, b) => (a.month < b.month ? -1 : a.month > b.month ? 1 : 0);
    const monthly = skMonthly.filter(m => m.month !== mk).concat([parsed.monthly]).sort(byMonth);
    const follow = skFollow.filter(f => f.month !== mk).concat(parsed.follow).sort(byMonth);
    const notes = skNotes.filter(n => n.month !== mk).concat(parsed.notes).sort(byMonth);
    return { monthly, follow, notes };
  }

  async function writeAll(merged) {
    if (demoMode || !window.Excel) {
      skMonthly = merged.monthly.map((m, i) => ({ ...m, row: i + 2 }));
      skFollow = merged.follow.map((f, i) => ({ ...f, row: i + 2 }));
      skNotes = merged.notes.map((n, i) => ({ ...n, row: i + 2 }));
      return;
    }
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SK_SHEET);
      const oldMax = Math.max(skMonthly.length, skFollow.length, skNotes.length, 1);
      const plan = [
        ["monthly", merged.monthly.map(monthlyToRow)],
        ["follow", merged.follow.map(followToRow)],
        ["note", merged.notes.map(noteToRow)],
      ];
      for (const [key, rows] of plan) {
        const b = BLOCKS[key];
        const clearTo = Math.max(oldMax, rows.length) + 1;
        sheet.getRange(`${b.first}2:${b.last}${clearTo}`).clear("Contents");
        if (!rows.length) continue;
        /* 対象月は文字列で保持（"2026/10" がExcelで日付に変換されないよう、値より先に書式を設定） */
        sheet.getRange(`${b.first}2:${b.first}${rows.length + 1}`).numberFormat = rows.map(() => ["@"]);
        sheet.getRange(`${b.first}2:${b.last}${rows.length + 1}`).values = rows;
      }
      const fRows = merged.follow.length + 1;
      if (fRows >= 2) {
        sheet.getRange(`T2:T${fRows}`).numberFormat = Array.from({ length: fRows - 1 }, () => ["m/d"]);
        sheet.getRange(`W2:W${fRows}`).numberFormat = Array.from({ length: fRows - 1 }, () => ["m/d"]);
      }
      await ctx.sync();
    });
    await skLoad();
  }

  /* ============================================================
     docx 解析
     ------------------------------------------------------------
     word/document.xml の body 直下を上から順に走査し、見出し
     （「1. アプローチ実績」等）ごとに表と段落を振り分ける。
     ============================================================ */
  function loadJSZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = JSZIP_URL;
      s.onload = () => window.JSZip ? resolve(window.JSZip) : reject(new Error("JSZip未定義"));
      s.onerror = () => reject(new Error("JSZipの読み込みに失敗しました"));
      document.head.appendChild(s);
    });
  }

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  function wAll(el, tag) { return Array.from(el.getElementsByTagNameNS(W_NS, tag)); }
  function pText(p) {
    let s = "";
    const walk = n => {
      for (const c of n.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.localName === "t") s += c.textContent;
        else if (c.localName === "tab") s += "\t";
        else if (c.localName === "br") s += "\n";
        else walk(c);
      }
    };
    walk(p);
    return s.trim();
  }
  function pStyle(p) {
    const ps = wAll(p, "pStyle")[0];
    return ps ? (ps.getAttributeNS(W_NS, "val") || ps.getAttribute("w:val") || "") : "";
  }
  function tblRows(tbl) {
    return Array.from(tbl.childNodes).filter(n => n.localName === "tr").map(tr =>
      Array.from(tr.childNodes).filter(n => n.localName === "tc").map(tc =>
        Array.from(tc.childNodes).filter(n => n.localName === "p").map(pText).join("\n").trim()));
  }

  async function extractBlocks(arrayBuffer) {
    const JSZip = await loadJSZip();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const f = zip.file("word/document.xml");
    if (!f) throw new Error("Word文書（.docx）ではありません。");
    const xml = await f.async("string");
    return blocksFromXml(xml);
  }
  function blocksFromXml(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    const blocks = [];
    for (const n of body.childNodes) {
      if (n.localName === "p") {
        const t = pText(n);
        if (t) blocks.push({ type: "p", text: t, style: pStyle(n) });
      } else if (n.localName === "tbl") {
        blocks.push({ type: "tbl", rows: tblRows(n) });
      }
    }
    return blocks;
  }

  /* 「10月21日」→ Date（報告月から年を補完） */
  function jpDate(text, baseMonth) {
    const m = String(text || "").match(/(\d{1,2})月(\d{1,2})日/);
    if (!m) return null;
    const [by, bm] = baseMonth.split("/").map(Number);
    let y = by;
    const mo = +m[1];
    if (mo < bm - 6) y++;
    else if (mo > bm + 6) y--;
    return new Date(y, mo - 1, +m[2]);
  }
  function stripBullet(s) { return String(s).replace(/^[\s・●○◆■\-－–*]+/, "").trim(); }
  function afterColon(s) { return String(s).replace(/^[^：:]*[：:]\s*/, "").trim(); }

  function parseReport(blocks) {
    const warn = [];
    const out = { month: "", monthly: {}, follow: [], notes: [], warnings: warn };
    let section = 0;
    let curProspect = null;
    let lastLead = "";      // 直前の段落（折り返し待ち表の見出し判定用）
    let inShokan = false;   // 3. の「所感」以降

    const headingNo = b => {
      if (b.type !== "p") return 0;
      const m = b.text.match(/^([0-9０-９]+)[\.．]\s*/);
      if (!m) return 0;
      if (!/heading|見出し|Title/i.test(b.style) && b.text.length > 30) return 0;
      return +String(m[1]).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
    };

    /* 対象月・委託先 */
    for (const b of blocks) {
      if (b.type !== "p") continue;
      if (!out.month) {
        const m = b.text.match(/対象期間[：:]\s*(\d{4})年(\d{1,2})月/) || b.text.match(/(\d{4})年(\d{1,2})月/);
        if (m) out.month = monthKeyYM(+m[1], +m[2]);
      }
      if (/TEL|ＴＥＬ|電話[：:]/.test(b.text) && /株式会社|有限会社|合同会社/.test(b.text)) {
        const first = b.text.split(/[｜|]/)[0].trim();
        const parts = first.split(/[\s　]+/).filter(Boolean);
        out.monthly.vendor = parts[0] || "";
        out.monthly.vendorPerson = parts.slice(1).join(" ");
      }
    }
    if (!out.month) throw new Error("対象期間（○○年○月）が見つかりませんでした。");
    const mk = out.month;
    Object.assign(out.monthly, { month: mk, calls: 0, appo: 0, targetNext: "" });
    BREAKDOWN.forEach(b => { out.monthly[b.key] = 0; });
    let breakdownFound = false;

    const pushProspect = () => { if (curProspect) { out.follow.push(curProspect); curProspect = null; } };

    for (const b of blocks) {
      const h = headingNo(b);
      if (h) { pushProspect(); section = h; inShokan = false; lastLead = ""; continue; }

      if (b.type === "tbl") {
        const rows = b.rows;
        const head = (rows[0] || []).map(s => s.replace(/\s/g, ""));
        const body = rows.slice(1);
        /* 1. 実績サマリー */
        if (head[0] === "項目") {
          body.forEach(r => {
            const label = r[0] || "", v = num(r[1]);
            if (/確認した企業|^架電(数|件数|社数)?$|^架電・/.test(label)) out.monthly.calls = v;
            else if (/アポ/.test(label)) out.monthly.appo = v;
          });
        }
        /* 1. 架電結果の内訳 */
        else if (head[0] === "結果") {
          breakdownFound = true;
          body.forEach(r => {
            const label = r[0] || "", v = num(r[1]);
            if (/合計/.test(label)) { if (!out.monthly.calls) out.monthly.calls = v; return; }
            if (/見込み/.test(label)) out.monthly.prospect = v;
            else if (/折り返し/.test(label)) out.monthly.callback = v;
            else if (/再架電/.test(label)) out.monthly.recall = v;
            else if (/不在/.test(label)) out.monthly.absent = v;
            else if (/時期尚早/.test(label)) out.monthly.ngEarly = v;
            else if (/NG|お断り|不要/.test(label)) out.monthly.ngNo = v;
            else if (/不通/.test(label)) out.monthly.noAnswer = v;
            else if (/アポ/.test(label)) out.monthly.appo = Math.max(out.monthly.appo, v);
            else out.monthly.other += v;
          });
        }
        /* 2. 見込み先（1行2列：会社名｜日付・区分） */
        else if (section === 2 && rows.length === 1 && rows[0].length >= 2) {
          pushProspect();
          const info = rows[0][1] || "";
          const label = (info.split(/[｜|]/)[1] || "").trim() || "見込み";
          const ch = (info.match(/[（(]([^）)]+)[）)]/) || [])[1] || "";
          curProspect = {
            month: mk, kind: /見込み/.test(label) ? label : `見込み（${label}）`,
            company: rows[0][0].replace(/\s+/g, " ").trim(), person: "",
            contact: jpDate(info, mk), memo: "", mail: "", nextDate: null, nextAction: "", ownStatus: "",
            _channel: ch,
          };
        }
        /* 2. 折り返し待ち／再架電予定の一覧 */
        else if (section === 2 && head[0] === "会社名") {
          pushProspect();
          const kind = /再架電/.test(lastLead) ? "再架電予定" : "折り返し待ち";
          const plan = (lastLead.split(/[―—\-－]/).slice(1).join("").trim()) || "";
          body.forEach(r => {
            if (!r[0]) return;
            out.follow.push({
              month: mk, kind, company: r[0].trim(), person: (r[1] || "").trim(), contact: null,
              memo: "", mail: "", nextDate: null, nextAction: plan, ownStatus: "",
            });
          });
        }
        /* 3. NG傾向 */
        else if (head[0] === "傾向") {
          body.forEach(r => { if (r[0] || r[1]) out.notes.push({ month: mk, kind: "NG傾向", title: r[0] || "", body: r[1] || "" }); });
        }
        /* 4. 来月の計画 */
        else if (/^No/i.test(head[0]) || head[1] === "アクション") {
          body.forEach(r => {
            const act = r[1] || r[0] || "";
            if (!act) return;
            out.notes.push({ month: mk, kind: "翌月計画", title: r[0] || "", body: act });
            const t = act.match(/月間\s*(\d+)\s*件/);
            if (t) out.monthly.targetNext = +t[1];
          });
        }
        lastLead = "";
        continue;
      }

      /* 段落 */
      const text = stripBullet(b.text);
      if (section === 2 && curProspect) {
        if (/^ステータス[：:]/.test(text)) { curProspect.memo += (curProspect.memo ? "\n" : "") + "【状況】" + afterColon(text); continue; }
        if (/^連絡先[：:]/.test(text)) { curProspect.mail = afterColon(text); continue; }
        if (/^ネクストアクション[：:]|^次回[：:]/.test(text)) {
          const v = afterColon(text);
          curProspect.nextDate = jpDate(v, mk);
          curProspect.nextAction = v.replace(/\d{1,2}月\d{1,2}日[\s　]*/, "").trim() || v;
          continue;
        }
        if (/^▼/.test(text)) { pushProspect(); lastLead = text; continue; }
        const pm = text.match(/^(.+?様)とお話/);
        if (pm && !curProspect.person) curProspect.person = pm[1];
        curProspect.memo += (curProspect.memo ? "\n" : "") + text;
        continue;
      }
      if (section === 3) {
        if (/^所感$/.test(text)) { inShokan = true; continue; }
        if (inShokan) out.notes.push({ month: mk, kind: "所感", title: "", body: text });
        continue;
      }
      lastLead = text;
    }
    pushProspect();
    out.follow.forEach(f => { delete f._channel; });

    if (!breakdownFound) warn.push("架電結果の内訳表が見つかりませんでした。");
    const total = sum(out.monthly);
    if (out.monthly.calls && total !== out.monthly.calls)
      warn.push(`内訳の合計（${total}社）と架電数（${out.monthly.calls}社）が一致しません。`);
    const prospects = out.follow.filter(f => /見込み/.test(f.kind)).length;
    if (prospects !== out.monthly.prospect)
      warn.push(`見込み件数（${out.monthly.prospect}）と見込み先の記載（${prospects}社）が一致しません。`);
    const cbs = out.follow.filter(f => f.kind === "折り返し待ち").length;
    if (cbs && cbs !== out.monthly.callback)
      warn.push(`折り返し待ち件数（${out.monthly.callback}）と一覧（${cbs}社）が一致しません。`);
    return out;
  }

  /* ============================================================
     取込 UI
     ============================================================ */
  function pickFile() {
    const inp = document.getElementById("sk-file");
    if (!inp) return;
    inp.value = "";
    inp.click();
  }
  async function onFileChosen(inp) {
    const file = inp.files && inp.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const blocks = await extractBlocks(buf);
      skPending = parseReport(blocks);
      skPending.fileName = file.name;
      inp.value = "";   // 同じファイルを続けて選んでも change が発火するように
      showPreview();
    } catch (e) {
      console.error(e);
      uiAlert("報告書を読み取れませんでした。\n" + (e.message || e));
    }
  }

  function showPreview() {
    const p = skPending;
    const exists = skMonthly.some(m => m.month === p.month);
    const kinds = {};
    p.follow.forEach(f => { kinds[f.kind] = (kinds[f.kind] || 0) + 1; });
    const notesBy = {};
    p.notes.forEach(n => { notesBy[n.kind] = (notesBy[n.kind] || 0) + 1; });
    const m = p.monthly;
    const html = `
      <div class="sk-pv-file">${esc(p.fileName || "")}</div>
      <table class="agg-table sk-pv-table">
        <tr><td class="l">対象月</td><td class="r"><b>${esc(monthLabel(p.month))}</b></td></tr>
        <tr><td class="l">委託先</td><td class="r">${esc(m.vendor || "（不明）")} ${esc(m.vendorPerson || "")}</td></tr>
        <tr><td class="l">架電数 / アポ</td><td class="r">${m.calls}社 / ${m.appo}件</td></tr>
        <tr><td class="l">内訳</td><td class="r sk-pv-bd">${BREAKDOWN.map(b => `${esc(b.label)} ${m[b.key]}`).join("、")}</td></tr>
        <tr><td class="l">企業フォロー</td><td class="r">${Object.entries(kinds).map(([k, v]) => `${esc(k)} ${v}社`).join("、") || "なし"}</td></tr>
        <tr><td class="l">所感・傾向</td><td class="r">${Object.entries(notesBy).map(([k, v]) => `${esc(k)} ${v}件`).join("、") || "なし"}</td></tr>
        <tr><td class="l">翌月目標</td><td class="r">${m.targetNext ? m.targetNext + "件" : "—"}</td></tr>
      </table>
      ${p.warnings.length ? `<div class="sk-warn">${p.warnings.map(w => `<div>${esc(w)}</div>`).join("")}</div>` : ""}
      ${exists ? `<div class="sk-over">${esc(monthLabel(p.month))}のデータは登録済みです。書き込むとこの月の行を置き換えます（自社対応状況の入力は引き継ぎます）。</div>` : ""}`;
    document.getElementById("sk-pv-body").innerHTML = html;
    document.getElementById("sk-pv-ok").textContent = exists ? "上書きする" : "シートに書き込む";
    document.getElementById("sk-pv-msg").textContent = "";
    document.getElementById("sk-preview").style.display = "";
  }
  function closePreview() {
    document.getElementById("sk-preview").style.display = "none";
    skPending = null;
  }
  async function commitPreview() {
    if (!skPending) return;
    const btn = document.getElementById("sk-pv-ok");
    const msg = document.getElementById("sk-pv-msg");
    btn.disabled = true;
    msg.textContent = "書き込み中…";
    try {
      const mk = skPending.month;
      await writeAll(mergeParsed(skPending));
      skMonth = mk;
      skTerm = termOfDate(new Date(+mk.slice(0, 4), +mk.slice(5) - 1, 1));
      closePreview();
      renderShinki();
    } catch (e) {
      console.error(e);
      msg.textContent = "書き込みに失敗しました：" + (e.message || e);
    } finally {
      btn.disabled = false;
    }
  }

  /* ============================================================
     描画
     ============================================================ */
  async function renderShinki() {
    const cont = document.getElementById("shinki-container");
    if (!cont) return;
    if (!skLoaded) {
      cont.innerHTML = `<div class="sk-empty">読み込み中…</div>`;
      await skLoad();
    }
    const months = [...new Set(skMonthly.map(m => m.month))].sort();
    if (!months.length) { cont.innerHTML = emptyHtml(); return; }
    if (!skMonth || !months.includes(skMonth)) skMonth = months[months.length - 1];
    if (skTerm == null) skTerm = termOfDate(new Date(+skMonth.slice(0, 4), +skMonth.slice(5) - 1, 1));

    const cur = skMonthly.find(m => m.month === skMonth);
    const prev = skMonthly.find(m => m.month === prevMonth(skMonth));
    cont.innerHTML =
      toolbarHtml(months, cur) +
      kpiHtml(cur, prev) +
      breakdownHtml(cur) +
      trendHtml() +
      followHtml() +
      notesHtml();
  }

  function emptyHtml() {
    return `<div class="sk-emptybox">
      <div class="sk-empty-ttl">営業代行の月次報告書を取り込みます</div>
      <div class="sk-empty-txt">受領した報告書（.docx）を選ぶと内容を読み取り、シート「新規営業」に月ごとに蓄積します。</div>
      <button class="btn-primary" onclick="shinkiPickFile()">報告書を取り込む</button>
    </div>`;
  }

  function toolbarHtml(months, cur) {
    const opts = months.slice().reverse().map(m =>
      `<option value="${m}"${m === skMonth ? " selected" : ""}>${esc(monthLabel(m))}</option>`).join("");
    const vendor = [cur.vendor, cur.vendorPerson].filter(Boolean).join(" ");
    return `<div class="sk-toolbar">
      <select class="filter-select" onchange="shinkiSelectMonth(this.value)">${opts}</select>
      <span class="sk-vendor">${esc(vendor)}</span>
      <span class="spacer"></span>
      <button class="btn-sub" onclick="shinkiOpenSheet()">シートを開く</button>
      <button class="btn-primary" onclick="shinkiPickFile()">報告書を取り込む</button>
    </div>`;
  }

  function deltaHtml(curV, prevV, suffix) {
    if (prevV == null) return "";
    const d = curV - prevV;
    if (!d) return `<div class="sk-delta">前月比 ±0</div>`;
    return `<div class="sk-delta ${d > 0 ? "up" : "down"}">前月比 ${d > 0 ? "+" : ""}${Math.round(d * 10) / 10}${suffix || ""}</div>`;
  }
  function reachRate(m) {
    return m.calls ? Math.round((m.calls - m.noAnswer - m.absent) / m.calls * 1000) / 10 : 0;
  }
  function kpiHtml(m, p) {
    const target = p && p.targetNext ? p.targetNext : null;
    const callsSub = target ? `<div class="sk-delta">目標 ${target}件（${Math.round(m.calls / target * 100)}%）</div>` : deltaHtml(m.calls, p ? p.calls : null, "社");
    return `<div class="kpi-row">
      <div class="kpi"><div class="kv">${m.calls}</div><div class="kl">架電数（社）</div>${callsSub}</div>
      <div class="kpi"><div class="kv">${m.appo}</div><div class="kl">アポ獲得</div>${deltaHtml(m.appo, p ? p.appo : null, "件")}</div>
      <div class="kpi"><div class="kv">${m.prospect}</div><div class="kl">見込み（社）</div>${deltaHtml(m.prospect, p ? p.prospect : null, "社")}</div>
      <div class="kpi"><div class="kv">${reachRate(m)}<span class="sk-unit">%</span></div><div class="kl">会話到達率</div>${deltaHtml(reachRate(m), p ? reachRate(p) : null, "pt")}</div>
    </div>`;
  }

  function breakdownHtml(m) {
    const total = sum(m) || 1;
    const segs = BREAKDOWN.filter(b => m[b.key]).map(b =>
      `<div class="sk-seg" style="flex:${m[b.key]};background:${b.color}" title="${esc(b.label)} ${m[b.key]}社（${(m[b.key] / total * 100).toFixed(1)}%）"></div>`).join("");
    const legend = BREAKDOWN.map(b =>
      `<span><span class="sw" style="background:${b.color}"></span>${esc(b.label)} <b>${m[b.key]}</b><span class="sk-pct">${(m[b.key] / total * 100).toFixed(1)}%</span></span>`).join("");
    const ng = m.ngNo + m.ngEarly;
    return `<div class="agg-card">
      <h3>架電結果の内訳<span class="sk-h-sub">${sum(m)}社 ／ NG計 ${ng}社</span></h3>
      <div class="sk-bar">${segs}</div>
      <div class="legend sk-legend">${legend}</div>
    </div>`;
  }

  function trendHtml() {
    const months = fiscalMonths(skTerm);
    const get = (mk, k) => { const r = skMonthly.find(x => x.month === mk); return r ? (r[k] || 0) : 0; };
    const calls = { "架電数": months.map(mk => get(mk, "calls")) };
    const react = {
      "アポ獲得": months.map(mk => get(mk, "appo")),
      "見込み": months.map(mk => get(mk, "prospect")),
      "折り返し待ち": months.map(mk => get(mk, "callback")),
      "再架電予定": months.map(mk => get(mk, "recall")),
    };
    const reactColors = { "アポ獲得": "#c0392b", "見込み": "#0e7a5f", "折り返し待ち": "#d9a038", "再架電予定": "#2c6e9b" };
    return `<div class="agg-card">
      <h3>月次推移
        <span class="sk-termbar">
          <button class="term-btn" onclick="shinkiShiftTerm(-1)">◀</button>
          <span class="term-label">${esc(termLabel(skTerm))}</span>
          <button class="term-btn" onclick="shinkiShiftTerm(1)">▶</button>
        </span>
      </h3>
      <div class="sk-sub-ttl">架電数（社）</div>
      <div class="chart-wrap">${groupedBarChart(months, calls, { "架電数": "#9fc9bb" })}</div>
      <div class="sk-sub-ttl">反応（件・社）</div>
      ${legendHtml(reactColors)}
      <div class="chart-wrap">${groupedBarChart(months, react, reactColors)}</div>
    </div>`;
  }

  function followHtml() {
    const list = skFollow.filter(f => f.month === skMonth);
    const t0 = today0();
    /* 名寄せ：初出月（何か月前から追っている先か） */
    const firstSeen = {};
    skFollow.forEach(f => {
      const k = companyKey(f.company);
      if (!firstSeen[k] || f.month < firstSeen[k]) firstSeen[k] = f.month;
    });
    const order = k => /見込み/.test(k) ? 0 : k === "折り返し待ち" ? 1 : k === "再架電予定" ? 2 : 3;
    const groups = {};
    list.forEach(f => { (groups[f.kind] = groups[f.kind] || []).push(f); });
    const kinds = Object.keys(groups).sort((a, b) => order(a) - order(b));
    const overdue = list.filter(f => f.nextDate && f.nextDate < t0 && !isClosed(f)).length;
    const pillCls = k => /見込み/.test(k) ? "sk-pill-pros" : k === "折り返し待ち" ? "sk-pill-cb" : "sk-pill-rc";

    const rows = kinds.map(k => {
      const items = groups[k].map(f => {
        const late = f.nextDate && f.nextDate < t0 && !isClosed(f);
        const since = firstSeen[companyKey(f.company)];
        const cont = since && since < skMonth ? `<span class="sk-since">${esc(monthLabel(since).replace(/^\d+年/, ""))}から継続</span>` : "";
        const next = f.nextDate || f.nextAction
          ? `<span class="sk-next${late ? " late" : ""}">${f.nextDate ? "次回 " + md(f.nextDate) : ""} ${esc(f.nextAction)}</span>` : "";
        const own = f.ownStatus ? `<span class="sk-own">${esc(f.ownStatus)}</span>` : "";
        const memo = f.memo ? `<div class="sk-memo">${esc(f.memo).replace(/\n/g, "<br>")}${f.mail ? `<div class="sk-mail">${esc(f.mail)}</div>` : ""}</div>` : "";
        return `<div class="sk-frow" onclick="shinkiJump(${f.row}, 'P')" title="シートの該当行へ移動">
          <div class="sk-fline">
            <span class="sk-co">${esc(f.company)}</span>
            <span class="sk-person">${esc(f.person)}</span>
            ${cont}
            <span class="spacer"></span>
            ${own}${next}
          </div>${memo}
        </div>`;
      }).join("");
      return `<div class="sk-fgroup"><div class="sk-fhead"><span class="sk-pill ${pillCls(k)}">${esc(k)}</span><span class="sk-cnt">${groups[k].length}社</span></div>${items}</div>`;
    }).join("");

    return `<div class="agg-card">
      <h3>フォロー対象<span class="sk-h-sub">${list.length}社${overdue ? ` ／ <span class="sk-late-cnt">期日超過 ${overdue}件</span>` : ""}</span></h3>
      ${rows || `<div class="sk-empty">この月のフォロー対象はありません。</div>`}
      <div class="sk-foot">行をクリックするとシートの該当行へ移動します。Y列「自社対応状況」に「完了」「済」等を入力すると期日超過の対象外になります。</div>
    </div>`;
  }

  function notesHtml() {
    const list = skNotes.filter(n => n.month === skMonth);
    if (!list.length) return "";
    const ng = list.filter(n => n.kind === "NG傾向");
    const shokan = list.filter(n => n.kind === "所感");
    const plan = list.filter(n => n.kind === "翌月計画");
    const other = list.filter(n => !["NG傾向", "所感", "翌月計画"].includes(n.kind));
    const ngHtml = ng.length ? `<div class="sk-sub-ttl">NG傾向</div><div class="sk-ng-grid">${ng.map(n =>
      `<div class="sk-ng"><div class="sk-ng-t">${esc(n.title)}</div><div class="sk-ng-b">${esc(n.body)}</div></div>`).join("")}</div>` : "";
    const shHtml = shokan.length ? `<div class="sk-sub-ttl">所感</div>${shokan.map(n => `<p class="sk-para">${esc(n.body)}</p>`).join("")}` : "";
    const plHtml = plan.length ? `<div class="sk-sub-ttl">翌月の計画</div><ol class="sk-plan">${plan.map(n => `<li>${esc(n.body)}</li>`).join("")}</ol>` : "";
    const otHtml = other.map(n => `<p class="sk-para"><b>${esc(n.kind)}</b> ${esc(n.title)} ${esc(n.body)}</p>`).join("");
    return `<div class="agg-card"><h3>所感・傾向</h3>${ngHtml}${shHtml}${plHtml}${otHtml}</div>`;
  }

  /* ============================================================
     操作
     ============================================================ */
  function selectMonth(v) { skMonth = v; renderShinki(); }
  function shiftTerm(d) { skTerm += d; renderShinki(); }
  async function jump(row, col) {
    if (demoMode || !window.Excel) return;
    try {
      await Excel.run(async ctx => {
        const s = ctx.workbook.worksheets.getItem(SK_SHEET);
        s.activate();
        s.getRange(`${col || "A"}${row}`).select();
        await ctx.sync();
      });
    } catch (e) { console.error(e); }
  }
  async function openSheet() {
    const r = skMonthly.find(m => m.month === skMonth);
    await jump(r ? r.row : 1, "A");
  }
  async function reload() { skLoaded = false; await renderShinki(); }

  /* ============================================================
     デモデータ（Excel外での表示確認用：2026年10月の報告書内容）
     ============================================================ */
  function skLoadDemo() {
    const d = (m, day) => new Date(2026, m - 1, day);
    skMonthly = [{
      row: 2, month: "2026/10", vendor: "ワンビジョンパートナーズ株式会社", vendorPerson: "宮下 孝太",
      calls: 404, appo: 0, prospect: 2, callback: 9, recall: 12, absent: 56, ngNo: 174, ngEarly: 9,
      noAnswer: 116, other: 26, targetNext: 300,
    }];
    const cb = [["株式会社アクティブレイン", "女性担当者"], ["サンコウ物流株式会社", "大木代表"], ["ジャパンロジコム株式会社", "野沢様"],
      ["ピーアンドピーロジスティクス株式会社", "女性担当者"], ["マイスカイ交通株式会社", "白石様（採用担当）"], ["今村運送株式会社", "鹿倉様"],
      ["原商事運輸株式会社", "桜井様"], ["城南観光バス株式会社", "担当者不在"], ["明成物流株式会社", "鈴木様（採用担当）"]];
    skFollow = [
      { month: "2026/10", kind: "見込み（長期フォロー）", company: "有限会社今井屋商事運送", person: "片岡様", contact: d(10, 21),
        memo: "配車管理は紙のみで対応。積み荷の条件が人ごとに異なり既存システムでは対応できない。\n【状況】カスタム開発の実績・事例を軸にご案内メール送付済み。",
        mail: "imaiya.ktwp@future.ocn.ne.jp", nextDate: d(11, 2), nextAction: "再架電", ownStatus: "" },
      { month: "2026/10", kind: "見込み（長期フォロー）", company: "有限会社八千代運送", person: "今井様", contact: d(10, 21),
        memo: "紙・エクセル・荷主指定システムが混在。荷主が費用負担するケースもあり一本化にハードル。\n【状況】状況が整ったタイミングでの提案を念頭にご案内メール送付済み。",
        mail: "imai-h@yachiyo-exp.co.jp", nextDate: d(11, 2), nextAction: "再架電", ownStatus: "" },
      ...cb.map(([c, p]) => ({ month: "2026/10", kind: "折り返し待ち", company: c, person: p, contact: null, memo: "", mail: "",
        nextDate: null, nextAction: "11月上旬に全件フォローコール予定", ownStatus: "" })),
    ].map((f, i) => ({ ...f, row: i + 2 }));
    skNotes = [
      ["NG傾向", "既にシステム導入済み", "今年から新システムを入れたという反応が多数。導入直後のため切り替えは困難。"],
      ["NG傾向", "従業員の年齢層による懸念", "ベテランドライバーが多くシステムへの拒否反応があるという声が複数。"],
      ["NG傾向", "小規模のため費用対効果が見えない", "10台前後の事業者は紙やエクセルで十分という認識が根強い。"],
      ["所感", "", "導入済みだがカスタマイズできない・複数システム混在の層へ切り口を変えると反応が変わる可能性。"],
      ["所感", "", "長期フォロー2社は提案を急がず関係構築を優先。"],
      ["翌月計画", "1", "折り返し待ち9件への再コンタクト（11月上旬）"],
      ["翌月計画", "2", "再架電予定12件への二次架電"],
      ["翌月計画", "3", "長期フォロー2社への継続アプローチ"],
      ["翌月計画", "4", "新規架電の継続（月間300件ペースを維持）"],
    ].map(([kind, title, body], i) => ({ row: i + 2, month: "2026/10", kind, title, body }));
  }

  /* ---------- 公開（HTML の onclick と app.js から呼ぶもの） ---------- */
  window.renderShinki = renderShinki;
  window.shinkiPickFile = pickFile;
  window.shinkiOnFile = onFileChosen;
  window.shinkiCommit = commitPreview;
  window.shinkiClosePreview = closePreview;
  window.shinkiSelectMonth = selectMonth;
  window.shinkiShiftTerm = shiftTerm;
  window.shinkiJump = jump;
  window.shinkiOpenSheet = openSheet;
  window.shinkiReload = reload;
  window.shinkiReset = () => { skLoaded = false; };
  /* テスト用（Node から解析ロジックを検証する） */
  window.__shinkiTest = { blocksFromXml, parseReport, mergeParsed };
})();
