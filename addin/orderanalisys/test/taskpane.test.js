const fs = require("fs"), path = require("path"), assert = require("assert");
const { JSDOM } = require("jsdom");
const SHEETS = require("./sheets.json");

/* ---- 最小限の Excel モック ---- */
const colNum = (s) => s.split("").reduce((t, c) => t * 26 + c.charCodeAt(0) - 64, 0);
function parseAddr(a) {
  a = a.replace(/\$/g, "").split("!").pop();
  const [p, q] = a.split(":");
  const one = (x) => { const m = x.match(/^([A-Z]*)(\d*)$/); return { c: m[1] ? colNum(m[1]) - 1 : null, r: m[2] ? Number(m[2]) - 1 : null }; };
  const A = one(p), B = q ? one(q) : A;
  return { r1: A.r ?? 0, c1: A.c ?? 0, r2: B.r ?? 1048575, c2: B.c ?? 16383 };
}
const permissive = () => new Proxy(function () {}, {
  get: (t, k) => (k === "then" ? undefined : (t[k] !== undefined ? t[k] : (t[k] = permissive()))),
  set: (t, k, v) => { t[k] = v; return true; }, apply: () => permissive()
});
const log = [];
function makeWorkbook(data) {
  const sheets = [];
  function mkSheet(name, grid) {
    const ws = { name, grid: grid.map((r) => r.slice()), visibility: "Visible", charts: [], dv: [], fills: {} };
    ws.cell = (r, c) => (ws.grid[r] && ws.grid[r][c] !== undefined ? ws.grid[r][c] : "");
    ws.setCell = (r, c, v) => { while (ws.grid.length <= r) ws.grid.push([]); ws.grid[r][c] = v; };
    function range(b) {
      const rg = {
        isNullObject: false,
        format: { fill: { set color(v) { for (let r = b.r1; r <= Math.min(b.r2, 5000); r++) ws.fills[r + ":" + b.c1] = v; }, clear() { for (let r = b.r1; r <= Math.min(b.r2, 5000); r++) delete ws.fills[r + ":" + b.c1]; } }, font: {}, set columnWidth(v) {} },
        dataValidation: { clear() { ws.dv = []; }, set rule(v) { ws.dv.push(v); }, set errorAlert(v) {} },
        set columnHidden(v) { ws.hiddenCol = b.c1; },
        set numberFormat(v) { assert(Array.isArray(v) && v.length === b.r2 - b.r1 + 1, "numberFormat shape"); },
        load() {},
        clear() { for (let r = b.r1; r <= Math.min(b.r2, ws.grid.length - 1); r++) for (let c = b.c1; c <= b.c2; c++) if (ws.grid[r]) ws.grid[r][c] = ""; },
        getIntersectionOrNullObject(o) { const x = o._b; const nb = { r1: Math.max(b.r1, x.r1), c1: Math.max(b.c1, x.c1), r2: Math.min(b.r2, x.r2), c2: Math.min(b.c2, x.c2) }; const z = range(nb); if (nb.r1 > nb.r2 || nb.c1 > nb.c2) z.isNullObject = true; return z; },
        get rowIndex() { return b.r1; }, get rowCount() { return b.r2 - b.r1 + 1; },
        get values() { const out = []; for (let r = b.r1; r <= b.r2; r++) { const row = []; for (let c = b.c1; c <= b.c2; c++) row.push(ws.cell(r, c)); out.push(row); } return out; },
        set values(v) { assert.equal(v.length, b.r2 - b.r1 + 1, "values rows"); v.forEach((row, i) => { assert.equal(row.length, b.c2 - b.c1 + 1, "values cols"); row.forEach((x, j) => ws.setCell(b.r1 + i, b.c1 + j, x)); }); },
        _b: b, worksheet: ws
      };
      return rg;
    }
    ws.getRange = (a) => range(parseAddr(a));
    ws.getRangeByIndexes = (r, c, rc, cc) => range({ r1: r, c1: c, r2: r + rc - 1, c2: c + cc - 1 });
    ws.getUsedRangeOrNullObject = () => {
      let maxR = -1, maxC = -1;
      ws.grid.forEach((row, r) => row.forEach((v, c) => { if (v !== "" && v !== null && v !== undefined) { maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); } }));
      const z = range({ r1: 0, c1: 0, r2: Math.max(maxR, 0), c2: Math.max(maxC, 0) }); if (maxR < 0) z.isNullObject = true; return z;
    };
    ws.getUsedRange = ws.getUsedRangeOrNullObject;
    ws.freezePanes = { freezeRows() {} }; ws.autoFilter = { apply() {} };
    ws.activate = () => { wb.active = ws.name; };
    ws.delete = () => { sheets.splice(sheets.indexOf(ws), 1); };
    ws.onChanged = { add(h) { ws.handler = h; return {}; } };
    ws.charts = { list: [], add(type, rg) { const ch = permissive(); ch._type = type; ch._src = rg._b; this.list.push(ch); log.push("chart " + type); return ch; } };
    ws.isNullObject = false;
    return ws;
  }
  const wb = {
    sheets,
    worksheets: {
      get items() { return sheets; }, load() {},
      getItem(n) { const s = sheets.find((x) => x.name === n); if (!s) throw new Error("no sheet " + n); return s; },
      getItemOrNullObject(n) { return sheets.find((x) => x.name === n) || { isNullObject: true }; },
      add(n) { if (sheets.find((x) => x.name === n)) throw new Error("dup " + n); const s = mkSheet(n, []); sheets.push(s); return s; }
    },
    getSelectedRange() { return wb.selected; }
  };
  Object.entries(data).forEach(([n, g]) => sheets.push(mkSheet(n, g)));
  return wb;
}

(async () => {
  const wb = makeWorkbook(SHEETS);
  const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
  const dom = new JSDOM(html, { runScripts: "outside-only" });
  const w = dom.window;
  let ready;
  w.Office = { HostType: { Excel: "Excel" }, onReady: (f) => { ready = f; }, context: { requirements: { isSetSupported: () => true }}, actions: { associate: (n, f) => { w.__actions = w.__actions || {}; w.__actions[n] = f; } } };
  w.Excel = {
    run: async (f) => f({ workbook: wb, runtime: {}, sync: async () => {} }),
    SheetVisibility: { hidden: "Hidden" }, ClearApplyTo: { contents: "Contents" },
    ChartType: { barStacked: "BarStacked", line: "Line" }, ChartSeriesBy: { columns: "Columns" }, ChartLegendPosition: { top: "Top" }
  };
  w.eval(fs.readFileSync(path.join(__dirname, "../core.js"), "utf8"));
  w.eval(fs.readFileSync(path.join(__dirname, "../taskpane.js"), "utf8"));
  const $ = (id) => w.document.getElementById(id);
  const tick = () => new Promise((r) => setTimeout(r, 20));

  await ready({ host: "Excel" });
  assert.equal($("main").hidden, false, "main visible");
  console.log("period options:", Array.from($("period").options).map((o) => o.value));
  console.log("banner:", $("bnSugTitle").textContent, "|", $("bnSugDesc").textContent);
  assert.equal($("bnSuggest").hidden, false);
  assert.equal($("vGroup").disabled, true);
  console.log("customer view top10:", $("k10").textContent, "rank rows:", $("rank").children.length);

  $("btnCreate").click(); await tick(); await tick();
  assert.equal($("msg").className, "msg", "no error: " + $("msg").textContent);
  const gs = wb.worksheets.getItem("グループ");
  console.log("group sheet rows:", gs.grid.length, "header:", gs.grid[0]);
  console.log("validation:", JSON.stringify(gs.dv[0]));
  console.log("after create: view group pressed", $("vGroup").getAttribute("aria-pressed"), "top10", $("k10").textContent, $("k10d").textContent, "| status", $("stAuto").textContent, $("stOk").textContent, $("stBlank").textContent);
  assert.equal($("k10").textContent, "51.6%");

  // ユーザーがC列を書き換える（西原商会九州 → 西原商会）
  const idx = gs.grid.findIndex((r) => String(r[1]).includes("西原商会九州　鹿児島"));
  gs.grid[idx][2] = "西原商会";
  await new Promise((r) => setTimeout(r, 1600));
  await gs.handler({ address: "グループ!C" + (idx + 1) }); await tick();
  console.log("changed banner:", $("bnChanged").hidden, $("bnChgText").textContent, "fill now:", gs.fills[idx + ":2"]);
  $("btnRecalc2").click(); await tick(); await tick();
  console.log("after recalc status:", $("stAuto").textContent, $("stOk").textContent, $("stBlank").textContent, "banner hidden", $("bnChanged").hidden);
  const rank3 = Array.from($("rank").children).slice(0, 3).map((d) => d.querySelector(".nm").textContent + " " + d.querySelector(".n").textContent);
  console.log("top3:", rank3);

  // 期間の切り替え
  $("period").value = "202607"; $("period").onchange({ target: $("period") });
  console.log("July:", $("kTot").textContent, $("k10").textContent, $("kMan").textContent);

  // 新しい請求先が追加された場合
  const src = wb.worksheets.getItem("202601から202608");
  src.grid.push(["", 77777, "株式会社マツヤ 札幌支店 XXXXXX", 5, 0, 0, ""]);
  $("btnRecalc").click(); await tick(); await tick();
  console.log("msg:", $("msg").textContent, "| last row:", gs.grid[gs.grid.length - 1]);

  // 確認済み
  $("btnConfirmAll").click(); await tick(); await tick();
  console.log("after confirm:", $("stAuto").textContent, $("stOk").textContent, "| msg", $("msg").textContent);

  // 出力
  $("btnOutput").click(); await tick(); await tick();
  const out = wb.sheets.find((s) => s.name.startsWith("集計_"));
  console.log("output sheet:", out && out.name, out && out.grid[4], "charts", out && out.charts.list.length, "| msg", $("msg").textContent);
  assert(out && out.charts.list.length === 2);
  // 出力シートは再読み込みで集計対象にならない
  $("btnRecalc").click(); await tick(); await tick();
  assert.deepEqual(Array.from($("period").options).map((o) => o.value).sort(), ["202601から202608", "202607"]);
  // セル範囲の解析
  const f = w.__taskpane.cellsInColumnC;
  assert.equal(f("グループ!C5"), 1); assert.equal(f("A1:B9"), 0); assert.equal(f("C1:C3"), 2); assert.equal(f("A3:E4"), 2);
  wb.active = null; await new Promise((r) => w.__actions.openGroupSheet({ completed: r })); assert.equal(wb.active, "グループ");
  console.log("ALL OK");
})().catch((e) => { console.error("FAIL", e); process.exit(1); });
