/* 受注件数分析アドイン：Excelに依存しない計算ロジック */
(function (root) {
  "use strict";

  const GROUP_SHEET = "グループ";
  const LIST_SHEET = "_グループ名";
  const OUTPUT_PREFIX = "集計_";
  const EXCLUDE_KEYS = new Set(["その他"]);

  const str = (v) => (v === null || v === undefined ? "" : String(v)).trim();
  const num = (v) => {
    if (typeof v === "number") return v;
    const s = str(v).replace(/,/g, "");
    return s !== "" && !isNaN(Number(s)) ? Number(s) : NaN;
  };

  /** 見出し行を探す。請求先CD・請求先・受注件数がそろった行を見出しとみなす */
  function findHeader(values, maxScan) {
    const lim = Math.min(values.length, maxScan || 15);
    for (let r = 0; r < lim; r++) {
      const row = values[r].map(str);
      const cd = row.indexOf("請求先CD");
      const nm = row.indexOf("請求先");
      const cnt = row.indexOf("受注件数");
      if (cd >= 0 && nm >= 0 && cnt >= 0) {
        return {
          row: r,
          cdCol: cd,
          nameCol: nm,
          cntCol: cnt,
          noteCol: row.indexOf("備考"),
          isDaily: row.indexOf("受注日") >= 0
        };
      }
    }
    return null;
  }

  /** 見出し行より下のデータを読み取る */
  function parseRecords(values, h) {
    const out = [];
    for (let r = h.row + 1; r < values.length; r++) {
      const row = values[r];
      const cd = num(row[h.cdCol]);
      const n = num(row[h.cntCol]);
      if (isNaN(cd) || isNaN(n)) continue;
      out.push({
        cd: cd,
        name: str(row[h.nameCol]),
        n: n,
        note: h.noteCol >= 0 ? str(row[h.noteCol]) : null
      });
    }
    return out;
  }

  /** 対象外のシートか（アドインが作るシート） */
  function isSystemSheet(name) {
    return name === GROUP_SHEET || name === LIST_SHEET || name.indexOf(OUTPUT_PREFIX) === 0;
  }

  /** 備考列のない期間でも連携区分がわかるよう、請求先CDごとの備考を集める */
  function buildNoteMap(sources) {
    const m = new Map();
    sources.forEach((s) => s.records.forEach((r) => {
      if (r.note !== null && !m.has(r.cd)) m.set(r.cd, r.note);
      else if (r.note && !m.get(r.cd)) m.set(r.cd, r.note);
    }));
    return m;
  }

  /** 社名のゆれを除いた比較用キー（株式会社・支店名・先頭コードを除く） */
  function groupKey(name) {
    let t = str(name).normalize("NFKC")
      .replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)|\(合\)/g, " ");
    t = t.replace(/^\s*(?=[A-Z0-9\-]*\d)[A-Z0-9\-]{4,}\s+/, "").replace(/\s+/g, " ").trim();
    let k = (t.split(" ")[0] || t).replace(/[(（].*$/, "").trim();
    return k.length >= 2 ? k : t;
  }

  /** 社名が同じ請求先が2件以上あるものをグループ候補にする。戻り値：Map(cd → グループ名) */
  function suggestGroups(customers) {
    const by = new Map();
    customers.forEach((c) => {
      const k = groupKey(c.name);
      if (!k || EXCLUDE_KEYS.has(k)) return;
      if (!by.has(k)) by.set(k, []);
      by.get(k).push(c.cd);
    });
    const m = new Map();
    by.forEach((cds, k) => { if (cds.length >= 2) cds.forEach((cd) => m.set(cd, k)); });
    return m;
  }

  /** 新しく見つかった請求先に、既存グループと同じ社名なら自動でグループ名を入れる */
  function suggestForNew(newCustomers, existing) {
    const keyToGroup = new Map();
    existing.forEach((e) => { if (e.group) keyToGroup.set(groupKey(e.name), e.group); });
    const m = new Map();
    newCustomers.forEach((c) => {
      const g = keyToGroup.get(groupKey(c.name));
      if (g) m.set(c.cd, g);
    });
    return m;
  }

  /**
   * 集計する。
   * records: [{cd,name,n}]、groupMap: Map(cd → グループ名)、noteMap: Map(cd → 備考)
   * mode: "group" | "customer"
   */
  function aggregate(records, groupMap, noteMap, mode) {
    const units = new Map();
    let total = 0;
    records.forEach((r) => {
      const g = mode === "group" ? (groupMap.get(r.cd) || "") : "";
      const key = g ? "g:" + g : "c:" + r.cd;
      if (!units.has(key)) units.set(key, { key: key, name: g || r.name, isGroup: !!g, n: 0, linked: 0, members: [] });
      const u = units.get(key);
      const note = r.note !== null && r.note !== undefined ? r.note : (noteMap.get(r.cd) || "");
      u.n += r.n;
      if (note) u.linked += r.n;
      u.members.push({ cd: r.cd, name: r.name, n: r.n, note: note });
      total += r.n;
    });
    const list = Array.from(units.values()).sort((a, b) => b.n - a.n || a.name.localeCompare(b.name, "ja"));
    let s = 0;
    list.forEach((u) => {
      u.members.sort((a, b) => b.n - a.n);
      s += u.n;
      u.share = total ? u.n / total * 100 : 0;
      u.cum = total ? s / total * 100 : 0;
    });
    const linked = list.reduce((t, u) => t + u.linked, 0);
    return { units: list, total: total, linked: linked };
  }

  /** 上位k件の累計シェア */
  function topShare(agg, k) {
    if (!agg.units.length) return 0;
    return agg.units[Math.min(k, agg.units.length) - 1].cum;
  }

  /** グループシートの行（A:E）を解釈する */
  function parseGroupSheet(values) {
    const rows = [];
    for (let r = 1; r < values.length; r++) {
      const cd = num(values[r][0]);
      if (isNaN(cd)) continue;
      rows.push({
        row: r,
        cd: cd,
        name: str(values[r][1]),
        group: str(values[r][2]),
        auto: str(values[r][4])
      });
    }
    return rows;
  }

  /** グループシートの状態別件数 */
  function groupStatus(rows) {
    let auto = 0, confirmed = 0, blank = 0;
    rows.forEach((r) => {
      if (!r.group) blank++;
      else if (r.auto && r.auto === r.group) auto++;
      else confirmed++;
    });
    const names = Array.from(new Set(rows.map((r) => r.group).filter(Boolean)));
    return { auto: auto, confirmed: confirmed, blank: blank, groups: names.length, names: names.sort((a, b) => a.localeCompare(b, "ja")) };
  }

  /** グループシートを初めて作るときの行を並べる（グループの件数順→グループ内の件数順→未所属） */
  function buildInitialGroupRows(customers, suggestion) {
    const gTotal = new Map();
    customers.forEach((c) => {
      const g = suggestion.get(c.cd);
      if (g) gTotal.set(g, (gTotal.get(g) || 0) + c.n);
    });
    return customers.map((c) => ({ cd: c.cd, name: c.name, n: c.n, group: suggestion.get(c.cd) || "" }))
      .sort((a, b) => {
        if (!!a.group !== !!b.group) return a.group ? -1 : 1;
        if (a.group && b.group && a.group !== b.group) return gTotal.get(b.group) - gTotal.get(a.group) || a.group.localeCompare(b.group, "ja");
        return b.n - a.n;
      });
  }

  /** 件数の合計がいちばん大きい期間（グループシートの参考件数に使う） */
  function primarySource(sources) {
    let best = null, bt = -1;
    sources.forEach((s) => {
      const t = s.records.reduce((a, r) => a + r.n, 0);
      if (t > bt) { bt = t; best = s; }
    });
    return best;
  }

  /** 全期間の請求先を重複なく集める。期間は重なることがあるので、件数は基準期間の値を使う */
  function collectCustomers(sources, primary) {
    const m = new Map();
    if (primary) primary.records.forEach((r) => m.set(r.cd, { cd: r.cd, name: r.name, n: r.n }));
    sources.forEach((s) => s.records.forEach((r) => {
      if (!m.has(r.cd)) m.set(r.cd, { cd: r.cd, name: r.name, n: 0 });
    }));
    return Array.from(m.values());
  }

  const api = {
    GROUP_SHEET, LIST_SHEET, OUTPUT_PREFIX,
    findHeader, parseRecords, isSystemSheet, buildNoteMap, groupKey, suggestGroups, suggestForNew,
    aggregate, topShare, parseGroupSheet, groupStatus, buildInitialGroupRows, collectCustomers, primarySource
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Core = api;
})(typeof self !== "undefined" ? self : this);
