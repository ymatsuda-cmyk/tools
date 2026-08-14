/* ステップ数集計アドイン
   データ  : パス | 拡張子 | 総ステップ | 実ステップ
   割当    : フォルダパス | 取引先名
   除外    : フォルダパス | メモ
   割当・除外はいずれも「最長一致した祖先が勝つ」で解決する。 */

var SHEET = {
  data: "データ",
  assign: "割当",
  exclude: "除外",
  outVendor: "集計_取引先",
  outExt: "集計_拡張子",
  outCross: "集計_クロス"
};

var UNASSIGNED = "（未割当）";
var DIRECT = "（直下）";

/* 相対パス行に直前のUNCルートを引き継ぐ。既に正規化済みのデータなら false のままで良い。 */
var CARRY_UNC_ROOT = false;

var state = {
  rows: [],
  tree: null,
  assignMap: new Map(),
  excludeMap: new Map(),
  vendors: [],
  tab: "assign",
  openNodes: new Set(),
  checked: new Set(),
  openL1: new Set(),
  openL2: new Set(),
  restChecked: new Set(),
  summary: null
};

/* ---------- 起動 ---------- */

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    document.querySelector(".boot-msg").textContent = "Excel で開いてください。";
    return;
  }
  bindUi();
  reload();
});

function bindUi() {
  document.querySelectorAll(".tab").forEach(function (b) {
    b.addEventListener("click", function () { switchTab(b.dataset.tab); });
  });
  byId("btn-reload").addEventListener("click", reload);

  byId("show-assigned").addEventListener("change", renderTree);
  byId("depth-open").addEventListener("change", function () {
    state.openNodes = new Set();
    expandTo(parseInt(this.value, 10));
    renderTree();
  });
  byId("btn-assign").addEventListener("click", function () { applyAssign("assign"); });
  byId("btn-unassign").addEventListener("click", function () { applyAssign("unassign"); });
  byId("btn-exclude").addEventListener("click", function () { applyAssign("exclude"); });
  byId("btn-include").addEventListener("click", function () { applyAssign("include"); });

  byId("kind").addEventListener("change", renderSummary);
  byId("drop-excluded").addEventListener("change", renderSummary);
  byId("group-mode").addEventListener("change", function () {
    state.openL2 = new Set();
    renderSummary();
  });
  byId("btn-collapse").addEventListener("click", function () {
    state.openL1 = new Set();
    state.openL2 = new Set();
    renderSummary();
  });
  byId("btn-export").addEventListener("click", exportSheets);

  byId("rest-depth").addEventListener("change", renderRest);
  byId("btn-rest-assign").addEventListener("click", assignFromRest);

  byId("tree").addEventListener("click", onTreeClick);
  byId("drill").addEventListener("click", onDrillClick);
  byId("rest").addEventListener("change", onRestChange);
}

function switchTab(name) {
  state.tab = name;
  document.querySelectorAll(".tab").forEach(function (b) {
    var on = b.dataset.tab === name;
    b.classList.toggle("is-active", on);
    b.setAttribute("aria-selected", on ? "true" : "false");
  });
  ["assign", "summary", "rest"].forEach(function (n) {
    byId("panel-" + n).hidden = n !== name;
  });
  if (name === "summary") renderSummary();
  if (name === "rest") renderRest();
}

/* ---------- 読み込み ---------- */

async function reload() {
  try {
    setStatus("読み込み中…");
    var raw = await readWorkbook();
    buildModel(raw);
    byId("boot").hidden = true;
    byId("app").hidden = false;
    expandTo(parseInt(byId("depth-open").value, 10));
    renderTree();
    if (state.tab === "summary") renderSummary();
    if (state.tab === "rest") renderRest();
    setStatus(fmt(state.rows.length) + " ファイル / 割当 " + state.assignMap.size +
      " 件 / 除外 " + state.excludeMap.size + " 件");
  } catch (e) {
    toast(String(e && e.message ? e.message : e), true);
    document.querySelector(".boot-msg").textContent = "読み込みに失敗しました：" + e;
  }
}

async function readWorkbook() {
  return await Excel.run(async function (ctx) {
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    var names = sheets.items.map(function (s) { return s.name; });
    if (names.indexOf(SHEET.data) < 0) {
      throw new Error("シート「" + SHEET.data + "」が見つかりません。");
    }
    [SHEET.assign, SHEET.exclude].forEach(function (n) {
      if (names.indexOf(n) < 0) {
        var s = sheets.add(n);
        s.getRange("A1:B1").values = n === SHEET.assign
          ? [["フォルダパス", "取引先名"]]
          : [["フォルダパス", "メモ"]];
      }
    });
    await ctx.sync();

    var d = sheets.getItem(SHEET.data).getUsedRange(true);
    var a = sheets.getItem(SHEET.assign).getRange("A1:B20000");
    var x = sheets.getItem(SHEET.exclude).getRange("A1:B20000");
    d.load("values");
    a.load("values");
    x.load("values");
    await ctx.sync();

    return { data: d.values || [], assign: a.values || [], exclude: x.values || [] };
  });
}

function buildModel(raw) {
  state.assignMap = pairsToMap(raw.assign);
  state.excludeMap = pairsToMap(raw.exclude);

  var rows = [];
  var carry = [];
  var body = raw.data;
  var start = looksLikeHeader(body[0]) ? 1 : 0;

  for (var i = start; i < body.length; i++) {
    var r = body[i];
    if (!r || r[0] === null || r[0] === undefined || r[0] === "") continue;

    var segs = splitPath(String(r[0]), carry);
    if (segs.length < 2) continue;

    var ext = String(r[1] === undefined || r[1] === null || r[1] === ""
      ? extOf(segs[segs.length - 1]) : r[1]).toLowerCase();

    rows.push({
      segs: segs,
      folder: segs.slice(0, segs.length - 1),
      ext: ext,
      total: num(r[2]),
      real: num(r[3])
    });
  }

  for (var j = 0; j < rows.length; j++) {
    var f = rows[j];
    var hitA = lookup(state.assignMap, f.folder);
    var hitX = lookup(state.excludeMap, f.folder);
    f.vendor = hitA ? hitA.value : null;
    f.assignDepth = hitA ? hitA.depth : 0;
    f.excluded = !!hitX;
  }

  state.rows = rows;
  state.tree = buildTree(rows);
  state.vendors = uniq(Array.from(state.assignMap.values())).sort(cmpJa);
  var dl = byId("vendor-list");
  dl.innerHTML = state.vendors.map(function (v) {
    return "<option value=\"" + esc(v) + "\"></option>";
  }).join("");
}

function pairsToMap(values) {
  var m = new Map();
  if (!values) return m;
  for (var i = 0; i < values.length; i++) {
    var k = values[i][0];
    if (k === null || k === undefined || k === "") continue;
    var key = normKey(String(k));
    if (i === 0 && (key === "フォルダパス" || key.toLowerCase() === "folder")) continue;
    var v = values[i][1];
    m.set(key, v === null || v === undefined ? "" : String(v));
  }
  return m;
}

function looksLikeHeader(r) {
  if (!r) return false;
  return typeof r[2] !== "number" || String(r[0]).indexOf("\\") < 0;
}

function splitPath(p, carry) {
  var s = String(p).replace(/\//g, "\\").trim();
  var unc = s.indexOf("\\\\") === 0;
  var segs = s.split("\\").filter(function (x) { return x.length > 0; });
  if (unc) {
    if (CARRY_UNC_ROOT && segs.length >= 3) {
      carry.length = 0;
      segs.slice(0, segs.length - 2).forEach(function (x) { carry.push(x); });
    }
    return segs;
  }
  if (CARRY_UNC_ROOT && carry.length) return carry.concat(segs);
  return segs;
}

function extOf(name) {
  var i = String(name).lastIndexOf(".");
  return i < 0 ? "" : String(name).slice(i + 1);
}

function normKey(s) {
  return String(s).replace(/\//g, "\\").replace(/^\\+/, "").replace(/\\+$/, "");
}

function lookup(map, folderSegs) {
  for (var i = folderSegs.length; i > 0; i--) {
    var k = folderSegs.slice(0, i).join("\\");
    if (map.has(k)) return { value: map.get(k), key: k, depth: i };
  }
  return null;
}

/* ---------- ツリー ---------- */

function buildTree(rows) {
  var root = node("", "");
  for (var i = 0; i < rows.length; i++) {
    var f = rows[i];
    var cur = root;
    for (var d = 0; d < f.folder.length; d++) {
      var name = f.folder[d];
      var kid = cur.kids.get(name);
      if (!kid) {
        kid = node(name, cur.path ? cur.path + "\\" + name : name);
        cur.kids.set(name, kid);
      }
      cur = kid;
      cur.files++;
      cur.total += f.total;
      cur.real += f.real;
    }
  }
  markTree(root);
  return root;
}

function node(name, path) {
  return {
    name: name, path: path, kids: new Map(),
    files: 0, total: 0, real: 0,
    vendor: null, inherited: null, excluded: false, partial: false
  };
}

function markTree(root) {
  walk(root, null, false);
  function walk(n, inheritedVendor, inheritedEx) {
    var own = state.assignMap.has(n.path) ? state.assignMap.get(n.path) : null;
    var ownEx = state.excludeMap.has(n.path);
    n.vendor = own;
    n.inherited = own ? null : inheritedVendor;
    n.excluded = ownEx || inheritedEx;
    n.ownExcluded = ownEx;
    var vend = own || inheritedVendor;
    var any = false;
    n.kids.forEach(function (k) {
      walk(k, vend, n.excluded);
      if (k.vendor || k.partial) any = true;
    });
    n.partial = !own && !inheritedVendor && any;
  }
}

function expandTo(depth) {
  if (!state.tree) return;
  (function walk(n, d) {
    if (d >= depth) return;
    n.kids.forEach(function (k) {
      if (k.kids.size) state.openNodes.add(k.path);
      walk(k, d + 1);
    });
  })(state.tree, 0);
}

function renderTree() {
  var host = byId("tree");
  var showAssigned = byId("show-assigned").checked;
  var kids = sortKids(state.tree);
  if (!kids.length) { host.innerHTML = "<p class=\"empty\">データがありません。</p>"; return; }
  var html = [];
  kids.forEach(function (k) { emit(k, 0); });
  host.innerHTML = html.join("");

  function emit(n, depth) {
    var hidden = !showAssigned && (n.vendor || n.inherited);
    if (hidden) return;
    var open = state.openNodes.has(n.path);
    var locked = !!(n.vendor || n.inherited);
    var badge = "";
    if (n.vendor) badge = "<span class=\"badge badge-vendor\">" + esc(n.vendor) + "</span>";
    else if (n.inherited) badge = "<span class=\"badge badge-inherit\">継承 " + esc(n.inherited) + "</span>";
    else if (n.partial) badge = "<span class=\"badge badge-partial\">一部割当済</span>";
    if (n.excluded) badge += "<span class=\"badge badge-excluded\">除外</span>";

    html.push(
      "<div class=\"node" + (locked ? " is-off" : "") + "\" style=\"padding-left:" + (depth * 14 + 4) + "px\">" +
      "<button class=\"twisty" + (n.kids.size ? "" : " is-leaf") + "\" data-toggle=\"" + esc(n.path) + "\"" +
      " aria-label=\"" + (open ? "折りたたむ" : "展開する") + "\">" + (open ? "▼" : "▶") + "</button>" +
      "<input type=\"checkbox\" data-pick=\"" + esc(n.path) + "\"" +
      (state.checked.has(n.path) ? " checked" : "") + (locked ? " disabled" : "") + ">" +
      "<span class=\"node-name\" title=\"" + esc(n.path) + "\">" + esc(n.name) + "</span>" +
      badge +
      "<span class=\"node-num\">" + fmt(n.total) + "</span>" +
      "</div>"
    );
    if (open) sortKids(n).forEach(function (c) { emit(c, depth + 1); });
  }
}

function sortKids(n) {
  var a = [];
  n.kids.forEach(function (k) { a.push(k); });
  a.sort(function (x, y) { return y.total - x.total || cmpJa(x.name, y.name); });
  return a;
}

function onTreeClick(ev) {
  var t = ev.target.closest("[data-toggle]");
  if (t) {
    var p = t.dataset.toggle;
    if (state.openNodes.has(p)) state.openNodes.delete(p); else state.openNodes.add(p);
    renderTree();
    return;
  }
  var c = ev.target.closest("[data-pick]");
  if (c) {
    if (c.checked) state.checked.add(c.dataset.pick);
    else state.checked.delete(c.dataset.pick);
  }
}

/* ---------- 割当の書き込み ---------- */

async function applyAssign(mode) {
  var picks = Array.from(state.checked);
  if (mode === "unassign" || mode === "include") {
    picks = picks.length ? picks : [];
  }
  if (!picks.length) { toast("フォルダを選択してください。", true); return; }

  var vendor = byId("vendor").value.trim();
  if (mode === "assign" && !vendor) { toast("取引先名を入力してください。", true); return; }

  if (mode === "assign") {
    picks.forEach(function (p) { state.assignMap.set(normKey(p), vendor); });
  } else if (mode === "unassign") {
    picks.forEach(function (p) { state.assignMap.delete(normKey(p)); });
  } else if (mode === "exclude") {
    picks.forEach(function (p) { state.excludeMap.set(normKey(p), "旧版・バックアップ"); });
  } else {
    picks.forEach(function (p) { state.excludeMap.delete(normKey(p)); });
  }

  try {
    await writeMap(SHEET.assign, ["フォルダパス", "取引先名"], state.assignMap);
    await writeMap(SHEET.exclude, ["フォルダパス", "メモ"], state.excludeMap);
    state.checked = new Set();
    await reload();
    toast(picks.length + " フォルダを更新しました。");
  } catch (e) {
    toast("書き込みに失敗しました：" + e, true);
  }
}

async function writeMap(sheetName, header, map) {
  await Excel.run(async function (ctx) {
    var s = ctx.workbook.worksheets.getItem(sheetName);
    s.getRange("A1:B20000").clear(Excel.ClearApplyTo.contents);
    await ctx.sync();

    var keys = Array.from(map.keys()).sort(cmpJa);
    var vals = [header];
    keys.forEach(function (k) { vals.push([k, map.get(k)]); });
    s.getRangeByIndexes(0, 0, vals.length, 2).values = vals;
    s.getRange("A:B").format.autofitColumns();
    await ctx.sync();
  });
}

/* ---------- 集計 ---------- */

function aggregate() {
  var kind = byId("kind").value;
  var drop = byId("drop-excluded").checked;
  var mode = byId("group-mode").value;

  var vendors = new Map();
  var extTotals = new Map();
  var dropped = 0;
  var files = 0;
  var grand = 0;

  for (var i = 0; i < state.rows.length; i++) {
    var f = state.rows[i];
    var v = kind === "real" ? f.real : f.total;
    if (f.excluded) {
      dropped += v;
      if (drop) continue;
    }
    var vName = f.vendor || UNASSIGNED;
    var gName = groupOf(f, mode);

    var ve = vendors.get(vName);
    if (!ve) { ve = { name: vName, value: 0, files: 0, groups: new Map() }; vendors.set(vName, ve); }
    ve.value += v; ve.files++;

    var ge = ve.groups.get(gName);
    if (!ge) { ge = { name: gName, value: 0, files: 0, exts: new Map() }; ve.groups.set(gName, ge); }
    ge.value += v; ge.files++;

    ge.exts.set(f.ext, (ge.exts.get(f.ext) || 0) + v);
    extTotals.set(f.ext, (extTotals.get(f.ext) || 0) + v);

    files++; grand += v;
  }

  var list = Array.from(vendors.values()).sort(function (a, b) {
    if (a.name === UNASSIGNED) return 1;
    if (b.name === UNASSIGNED) return -1;
    return b.value - a.value;
  });
  list.forEach(function (ve) {
    ve.groupList = Array.from(ve.groups.values()).sort(function (a, b) { return b.value - a.value; });
    ve.groupList.forEach(function (ge) {
      ge.extList = Array.from(ge.exts.entries())
        .map(function (e) { return { name: e[0] || "(なし)", value: e[1] }; })
        .sort(function (a, b) { return b.value - a.value; });
    });
  });

  return {
    kind: kind, mode: mode, vendors: list, grand: grand, files: files, dropped: dropped,
    exts: Array.from(extTotals.entries())
      .map(function (e) { return { name: e[0] || "(なし)", value: e[1] }; })
      .sort(function (a, b) { return b.value - a.value; })
  };
}

function groupOf(f, mode) {
  if (mode === "ext") return null;
  if (mode === "under") {
    var d = f.assignDepth;
    return f.folder.length > d ? f.folder[d] : DIRECT;
  }
  var idx = mode === "d1" ? 0 : 1;
  return f.folder.length > idx ? f.folder[idx] : DIRECT;
}

function renderSummary() {
  var s = aggregate();
  state.summary = s;
  byId("m-steps").textContent = fmt(s.grand);
  byId("m-files").textContent = fmt(s.files);
  byId("m-dropped").textContent = fmt(s.dropped);

  var flat = s.mode === "ext";
  var host = byId("drill");
  if (!s.vendors.length) { host.innerHTML = "<p class=\"empty\">集計対象がありません。</p>"; return; }

  var max = s.vendors[0].value || 1;
  var html = ["<div class=\"dl-head\"><span class=\"g\">取引先" +
    (flat ? " › 拡張子" : " › " + labelOfMode(s.mode) + " › 拡張子") +
    "</span><span class=\"n\">ステップ</span><span class=\"p\">構成比</span></div>"];

  s.vendors.forEach(function (ve, i) {
    var open = state.openL1.has(ve.name);
    var pct = s.grand ? ve.value / s.grand * 100 : 0;
    var rest = ve.name === UNASSIGNED;
    html.push(
      "<div class=\"dl-row is-click" + (open ? " is-open" : "") + "\" data-l1=\"" + esc(ve.name) + "\"" +
      " role=\"button\" tabindex=\"0\" aria-expanded=\"" + open + "\">" +
      "<div class=\"dl-line\">" +
      "<span class=\"twisty\" aria-hidden=\"true\">" + (open ? "▼" : "▶") + "</span>" +
      "<span class=\"dl-name\">" + esc(ve.name) + "</span>" +
      "<span class=\"dl-num\">" + fmt(ve.value) + "</span>" +
      "<span class=\"dl-pct\">" + pct.toFixed(1) + "%</span>" +
      "</div>" +
      bar(ve.value / max * 100, rest) +
      "</div>"
    );
    if (!open) return;

    if (flat) {
      html.push("<div class=\"lvl2\">");
      var ex = mergeExts(ve);
      ex.forEach(function (e) {
        html.push(leafRow(e.name, e.value, ve.value, true));
      });
      html.push("</div>");
      return;
    }

    html.push("<div class=\"lvl2\">");
    ve.groupList.forEach(function (ge) {
      var key = ve.name + "\u0000" + ge.name;
      var o2 = state.openL2.has(key);
      var p2 = ve.value ? ge.value / ve.value * 100 : 0;
      html.push(
        "<div class=\"dl-row is-click" + (o2 ? " is-open" : "") + "\" data-l2=\"" + esc(key) + "\"" +
        " role=\"button\" tabindex=\"0\" aria-expanded=\"" + o2 + "\">" +
        "<div class=\"dl-line\">" +
        "<span class=\"twisty\" aria-hidden=\"true\">" + (o2 ? "▼" : "▶") + "</span>" +
        "<span class=\"dl-name mono\">" + esc(ge.name) + "</span>" +
        "<span class=\"dl-num\">" + fmt(ge.value) + "</span>" +
        "<span class=\"dl-pct\">" + p2.toFixed(1) + "%</span>" +
        "</div>" + bar(p2, rest) + "</div>"
      );
      if (o2) {
        html.push("<div class=\"lvl3\">");
        ge.extList.forEach(function (e) {
          html.push(leafRow(e.name, e.value, ge.value, false));
        });
        html.push("</div>");
      }
    });
    html.push("</div>");
  });

  html.push("<div class=\"dl-total\"><span class=\"g\">合計</span><span class=\"n\">" +
    fmt(s.grand) + "</span><span class=\"p\"></span></div>");
  host.innerHTML = html.join("");
}

function leafRow(name, value, denom, indent) {
  var p = denom ? value / denom * 100 : 0;
  return "<div class=\"dl-row\">" +
    "<div class=\"dl-line\">" +
    (indent ? "<span class=\"twisty is-leaf\" aria-hidden=\"true\"></span>" : "") +
    "<span class=\"dl-name mono\">" + esc(name) + "</span>" +
    "<span class=\"dl-num\">" + fmt(value) + "</span>" +
    "<span class=\"dl-pct\">" + p.toFixed(1) + "%</span>" +
    "</div>" + bar(p, false) + "</div>";
}

function mergeExts(ve) {
  var m = new Map();
  ve.groupList.forEach(function (ge) {
    ge.extList.forEach(function (e) { m.set(e.name, (m.get(e.name) || 0) + e.value); });
  });
  return Array.from(m.entries())
    .map(function (e) { return { name: e[0], value: e[1] }; })
    .sort(function (a, b) { return b.value - a.value; });
}

function bar(pct, rest) {
  var w = Math.max(0, Math.min(100, pct));
  return "<div class=\"bar" + (rest ? " is-rest" : "") + "\"><i style=\"width:" + w.toFixed(1) + "%\"></i></div>";
}

function labelOfMode(m) {
  return m === "under" ? "割当直下フォルダ" : m === "d1" ? "階層1" : "階層2";
}

function onDrillClick(ev) {
  var l2 = ev.target.closest("[data-l2]");
  if (l2) {
    var k = l2.dataset.l2;
    if (state.openL2.has(k)) state.openL2.delete(k); else state.openL2.add(k);
    renderSummary();
    return;
  }
  var l1 = ev.target.closest("[data-l1]");
  if (l1) {
    var n = l1.dataset.l1;
    if (state.openL1.has(n)) state.openL1.delete(n); else state.openL1.add(n);
    renderSummary();
  }
}

document.addEventListener("keydown", function (ev) {
  if (ev.key !== "Enter" && ev.key !== " ") return;
  var t = ev.target.closest("[data-l1],[data-l2]");
  if (!t) return;
  ev.preventDefault();
  t.click();
});

/* ---------- 未割当 ---------- */

function renderRest() {
  var depth = parseInt(byId("rest-depth").value, 10);
  var kind = byId("kind").value;
  var m = new Map();
  var sum = 0;

  state.rows.forEach(function (f) {
    if (f.vendor) return;
    var v = kind === "real" ? f.real : f.total;
    var key = f.folder.slice(0, Math.min(depth, f.folder.length)).join("\\");
    var e = m.get(key);
    if (!e) { e = { path: key, value: 0, files: 0 }; m.set(key, e); }
    e.value += v; e.files++;
    sum += v;
  });

  var list = Array.from(m.values()).sort(function (a, b) { return b.value - a.value; });
  byId("rest-total").textContent = list.length + " フォルダ / " + fmt(sum);

  var host = byId("rest");
  if (!list.length) { host.innerHTML = "<p class=\"empty\">未割当はありません。</p>"; return; }
  host.innerHTML = list.map(function (e) {
    return "<div class=\"rest-row\">" +
      "<input type=\"checkbox\" data-rest=\"" + esc(e.path) + "\"" +
      (state.restChecked.has(e.path) ? " checked" : "") + ">" +
      "<span class=\"rest-path\" title=\"" + esc(e.path) + "\">" + esc(e.path) + "</span>" +
      "<span class=\"rest-num\">" + fmt(e.value) + "</span>" +
      "</div>";
  }).join("");
}

function onRestChange(ev) {
  var c = ev.target.closest("[data-rest]");
  if (!c) return;
  if (c.checked) state.restChecked.add(c.dataset.rest);
  else state.restChecked.delete(c.dataset.rest);
}

async function assignFromRest() {
  var vendor = byId("vendor").value.trim();
  if (!vendor) { toast("「割当」タブで取引先名を入力してください。", true); switchTab("assign"); return; }
  if (!state.restChecked.size) { toast("フォルダを選択してください。", true); return; }
  state.restChecked.forEach(function (p) { state.assignMap.set(normKey(p), vendor); });
  state.restChecked = new Set();
  try {
    await writeMap(SHEET.assign, ["フォルダパス", "取引先名"], state.assignMap);
    await reload();
    switchTab("rest");
    toast("割り当てました。");
  } catch (e) {
    toast("書き込みに失敗しました：" + e, true);
  }
}

/* ---------- 出力 ---------- */

async function exportSheets() {
  var s = state.summary || aggregate();
  var label = s.kind === "real" ? "実ステップ" : "総ステップ";

  var v = [["取引先名", "ファイル数", label, "構成比"]];
  s.vendors.forEach(function (ve) {
    v.push([ve.name, ve.files, ve.value, s.grand ? ve.value / s.grand : 0]);
  });
  v.push(["合計", s.files, s.grand, 1]);

  var x = [["拡張子", label, "構成比"]];
  s.exts.forEach(function (e) {
    x.push([e.name, e.value, s.grand ? e.value / s.grand : 0]);
  });

  var c = [["取引先名", labelOfMode(s.mode), "拡張子", "ファイル数", label]];
  s.vendors.forEach(function (ve) {
    ve.groupList.forEach(function (ge) {
      ge.extList.forEach(function (e) {
        c.push([ve.name, ge.name, e.name, ge.files, e.value]);
      });
    });
  });

  try {
    await Excel.run(async function (ctx) {
      await putSheet(ctx, SHEET.outVendor, v, [3]);
      await putSheet(ctx, SHEET.outExt, x, [2]);
      await putSheet(ctx, SHEET.outCross, c, []);
      ctx.workbook.worksheets.getItem(SHEET.outVendor).activate();
      await ctx.sync();
    });
    toast("3 シートに出力しました。");
  } catch (e) {
    toast("出力に失敗しました：" + e, true);
  }
}

async function putSheet(ctx, name, values, pctCols) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  await ctx.sync();
  var exists = sheets.items.some(function (i) { return i.name === name; });
  var s = exists ? sheets.getItem(name) : sheets.add(name);
  if (exists) {
    s.getRange("A1:Z50000").clear(Excel.ClearApplyTo.all);
    await ctx.sync();
  }
  var r = s.getRangeByIndexes(0, 0, values.length, values[0].length);
  r.values = values;
  s.getRangeByIndexes(0, 0, 1, values[0].length).format.font.bold = true;
  pctCols.forEach(function (ci) {
    s.getRangeByIndexes(1, ci, values.length - 1, 1).numberFormat = [["0.0%"]];
  });
  s.getRangeByIndexes(0, 0, values.length, values[0].length).format.autofitColumns();
  await ctx.sync();
}

/* ---------- ユーティリティ ---------- */

function byId(id) { return document.getElementById(id); }
function num(v) { return typeof v === "number" ? v : (parseFloat(String(v).replace(/,/g, "")) || 0); }
function fmt(n) { return Math.round(n).toLocaleString("ja-JP"); }
function uniq(a) { return Array.from(new Set(a.filter(function (x) { return x; }))); }
function cmpJa(a, b) { return String(a).localeCompare(String(b), "ja"); }
function esc(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c];
  });
}
function setStatus(t) { byId("status-count").textContent = t; }

var toastTimer = null;
function toast(msg, isError) {
  var el = byId("toast");
  el.textContent = msg;
  el.classList.toggle("is-error", !!isError);
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.hidden = true; }, 4000);
}
