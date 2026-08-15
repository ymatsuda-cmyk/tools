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
/* 継承を断ち切るための特別マーカー。割当シートには読める形で書き込む。 */
var UNASSIGN_MARK = "（解除）";

/* 相対パス行に直前のUNCルートを引き継ぐ。既に正規化済みのデータなら false のままで良い。 */
var CARRY_UNC_ROOT = false;

/* Excel for Web は1回の sync で扱えるセル数に上限があるため、
   1万行規模のシートは CHUNK 行ずつに分けて読み書きする。 */
var CHUNK = 2000;

var state = {
  rows: [],
  tree: null,
  nodeIndex: new Map(),
  assignMap: new Map(),
  excludeMap: new Map(),
  vendors: [],
  tab: "assign",
  openNodes: new Set(),
  checked: new Set(),
  openL1: new Set(),
  openL2: new Set(),
  openFiles: new Set(),
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

  byId("kind").addEventListener("change", function () { state.openFiles = new Set(); renderSummary(); });
  byId("drop-excluded").addEventListener("change", function () { state.openFiles = new Set(); renderSummary(); });
  byId("group-mode").addEventListener("change", function () {
    state.openL2 = new Set();
    state.openFiles = new Set();
    renderSummary();
  });
  byId("btn-collapse").addEventListener("click", function () {
    state.openL1 = new Set();
    state.openL2 = new Set();
    state.openFiles = new Set();
    renderSummary();
  });
  byId("btn-export").addEventListener("click", exportSheets);

  byId("tree").addEventListener("click", onTreeClick);
  byId("drill").addEventListener("click", onDrillClick);

  document.addEventListener("keydown", function (ev) {
    if (ev.key !== "Enter" && ev.key !== " ") return;
    var t = ev.target.closest("[data-l1],[data-l2],[data-vendor]");
    if (!t) return;
    ev.preventDefault();
    t.click();
  });
}

function switchTab(name) {
  state.tab = name;
  document.querySelectorAll(".tab").forEach(function (b) {
    var on = b.dataset.tab === name;
    b.classList.toggle("is-active", on);
    b.setAttribute("aria-selected", on ? "true" : "false");
  });
  ["assign", "summary"].forEach(function (n) {
    byId("panel-" + n).hidden = n !== name;
  });
  if (name === "summary") renderSummary();
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
    setStatus(fmt(state.rows.length) + " ファイル / 割当 " + state.assignMap.size +
      " 件 / 除外 " + state.excludeMap.size + " 件");
  } catch (e) {
    var d = describe(e);
    toast("読み込みに失敗しました：" + d, true);
    var b = document.querySelector(".boot-msg");
    if (b) b.textContent = "読み込みに失敗しました：" + d;
    setStatus("読み込み失敗");
    console.error(e);
  }
}

async function readWorkbook() {
  await ensureSheets();
  var data = await readArea(SHEET.data, 4);
  var assign = await readArea(SHEET.assign, 2);
  var exclude = await readArea(SHEET.exclude, 2);
  return { data: data, assign: assign, exclude: exclude };
}

async function ensureSheets() {
  await Excel.run(async function (ctx) {
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    var names = sheets.items.map(function (s) { return s.name; });
    if (names.indexOf(SHEET.data) < 0) {
      throw new Error("シート「" + SHEET.data + "」が見つかりません。");
    }
    var added = false;
    [SHEET.assign, SHEET.exclude].forEach(function (n) {
      if (names.indexOf(n) >= 0) return;
      var s = sheets.add(n);
      s.getRange("A1:B1").values = n === SHEET.assign
        ? [["フォルダパス", "取引先名"]]
        : [["フォルダパス", "メモ"]];
      added = true;
    });
    if (added) await ctx.sync();
  });
}

/* シートの使用範囲の行数を測り、A列から cols 列ぶんを分割して読む */
async function readArea(sheetName, cols) {
  var dim = await areaSize(sheetName);
  if (!dim || !dim.rows) return [];

  var out = [];
  var read = 0;
  while (read < dim.rows) {
    var take = Math.min(CHUNK, dim.rows - read);
    var offset = read;
    var chunk = await Excel.run(async function (ctx) {
      var r = ctx.workbook.worksheets.getItem(sheetName)
        .getRangeByIndexes(dim.top + offset, 0, take, cols);
      r.load("values");
      await ctx.sync();
      return r.values;
    });
    for (var i = 0; i < chunk.length; i++) out.push(chunk[i]);
    read += take;
    if (dim.rows > CHUNK) {
      setStatus("読み込み中… " + fmt(read) + " / " + fmt(dim.rows) + " 行");
    }
  }
  return out;
}

async function areaSize(sheetName) {
  try {
    return await Excel.run(async function (ctx) {
      var r = ctx.workbook.worksheets.getItem(sheetName).getUsedRange(true);
      r.load(["rowCount", "columnCount", "rowIndex"]);
      await ctx.sync();
      return { rows: r.rowCount, cols: r.columnCount, top: r.rowIndex };
    });
  } catch (e) {
    return null; /* 空シート */
  }
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
    f.vendor = (hitA && hitA.value !== UNASSIGN_MARK) ? hitA.value : null;
    f.assignDepth = hitA ? hitA.depth : 0;
    f.excluded = !!hitX;
  }

  state.rows = rows;
  state.nodeIndex = new Map();
  state.tree = buildTree(rows);
  state.vendors = uniq(Array.from(state.assignMap.values())
    .filter(function (v) { return v !== UNASSIGN_MARK; })).sort(cmpJa);
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
  state.nodeIndex.set("", root);
  for (var i = 0; i < rows.length; i++) {
    var f = rows[i];
    var cur = root;
    for (var d = 0; d < f.folder.length; d++) {
      var name = f.folder[d];
      var kid = cur.kids.get(name);
      if (!kid) {
        var path = cur.path ? cur.path + "\\" + name : name;
        kid = node(name, path);
        cur.kids.set(name, kid);
        state.nodeIndex.set(path, kid);
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
    vendor: null, inherited: null, excluded: false, partial: false, overridden: false
  };
}

function markTree(root) {
  walk(root, null, false);
  function walk(n, inheritedVendor, inheritedEx) {
    var raw = state.assignMap.has(n.path) ? state.assignMap.get(n.path) : null;
    var isOverride = raw === UNASSIGN_MARK;
    var own = (raw && !isOverride) ? raw : null;
    var ownEx = state.excludeMap.has(n.path);
    n.vendor = own;
    n.overridden = isOverride;
    n.inherited = (own || isOverride) ? null : inheritedVendor;
    n.excluded = ownEx || inheritedEx;
    n.ownExcluded = ownEx;
    var vend = own || (isOverride ? null : inheritedVendor);
    var any = false;
    n.kids.forEach(function (k) {
      walk(k, vend, n.excluded);
      if (k.vendor || k.partial) any = true;
    });
    n.partial = !own && !isOverride && !inheritedVendor && any;
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

/* チェック状態を配下すべてに伝播させる */
function setSubtreeChecked(n, val) {
  if (val) state.checked.add(n.path); else state.checked.delete(n.path);
  n.kids.forEach(function (k) { setSubtreeChecked(k, val); });
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
    var badge = "";
    if (n.vendor) badge = "<span class=\"badge badge-vendor\">" + esc(n.vendor) + "</span>";
    else if (n.overridden) badge = "<span class=\"badge badge-override\">解除済み</span>";
    else if (n.inherited) badge = "<span class=\"badge badge-inherit\">継承 " + esc(n.inherited) + "</span>";
    else if (n.partial) badge = "<span class=\"badge badge-partial\">一部割当済</span>";
    if (n.excluded) badge += "<span class=\"badge badge-excluded\">除外</span>";

    html.push(
      "<div class=\"node" + ((n.vendor || n.inherited) ? " is-off" : "") +
      "\" style=\"padding-left:" + (depth * 14 + 4) + "px\">" +
      "<button class=\"twisty" + (n.kids.size ? "" : " is-leaf") + "\" data-toggle=\"" + esc(n.path) + "\"" +
      " aria-label=\"" + (open ? "折りたたむ" : "展開する") + "\">" + (open ? "▼" : "▶") + "</button>" +
      "<input type=\"checkbox\" data-pick=\"" + esc(n.path) + "\"" +
      (state.checked.has(n.path) ? " checked" : "") + ">" +
      "<span class=\"node-name\" title=\"" + esc(n.path) + "\">" + esc(n.name) + "</span>" +
      badge +
      "<span class=\"node-num\">" + fmtPair(n.files, n.total) + "</span>" +
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
    var n = state.nodeIndex.get(c.dataset.pick);
    if (n) setSubtreeChecked(n, c.checked);
    else if (c.checked) state.checked.add(c.dataset.pick);
    else state.checked.delete(c.dataset.pick);
    renderTree();
  }
}

/* ---------- 割当の書き込み ---------- */

/* 選択された経路のうち、祖先も選択されている経路を取り除く
   （配下を丸ごとチェックしても、割当は最上位のフォルダにだけ書けば済むため） */
function topmostPicks(paths) {
  var set = new Set(paths);
  return paths.filter(function (p) {
    var segs = p.split("\\");
    for (var i = 1; i < segs.length; i++) {
      if (set.has(segs.slice(0, i).join("\\"))) return false;
    }
    return true;
  });
}

async function applyAssign(mode) {
  var all = Array.from(state.checked);
  if (!all.length) { toast("フォルダを選択してください。", true); return; }

  var vendor = byId("vendor").value.trim();
  if (mode === "assign" && !vendor) { toast("取引先名を入力してください。", true); return; }

  var count = 0;
  if (mode === "assign") {
    var picks = topmostPicks(all);
    picks.forEach(function (p) { state.assignMap.set(normKey(p), vendor); });
    count = picks.length;
  } else if (mode === "exclude") {
    var picksX = topmostPicks(all);
    picksX.forEach(function (p) { state.excludeMap.set(normKey(p), "旧版・バックアップ"); });
    count = picksX.length;
  } else if (mode === "unassign") {
    var picksU = topmostPicks(all);
    var removed = 0, cut = 0;
    picksU.forEach(function (p) {
      var key = normKey(p);
      var n = state.nodeIndex.get(p);
      var raw = state.assignMap.has(key) ? state.assignMap.get(key) : undefined;
      if (raw !== undefined) {
        state.assignMap.delete(key);
        removed++;
      } else if (n && n.inherited) {
        state.assignMap.set(key, UNASSIGN_MARK);
        cut++;
      }
    });
    if (!removed && !cut) {
      toast("選択した項目には割当も継承もありません。", true);
      return;
    }
    count = removed + cut;
  } else {
    all.forEach(function (p) { if (state.excludeMap.delete(normKey(p))) count++; });
    if (!count) { toast("選択した項目に個別の除外がありません。除外元のフォルダを選んでください。", true); return; }
  }

  try {
    await writeMap(SHEET.assign, ["フォルダパス", "取引先名"], state.assignMap);
    await writeMap(SHEET.exclude, ["フォルダパス", "メモ"], state.excludeMap);
    state.checked = new Set();
    await reload();
    toast(count + " フォルダを更新しました。");
  } catch (e) {
    toast("書き込みに失敗しました：" + describe(e), true);
  }
}

async function writeMap(sheetName, header, map) {
  var keys = Array.from(map.keys()).sort(cmpJa);
  var vals = [header];
  keys.forEach(function (k) { vals.push([k, map.get(k)]); });

  var dim = await areaSize(sheetName);
  if (dim && dim.rows) {
    await Excel.run(async function (ctx) {
      ctx.workbook.worksheets.getItem(sheetName)
        .getRangeByIndexes(0, 0, dim.top + dim.rows, 2)
        .clear(Excel.ClearApplyTo.contents);
      await ctx.sync();
    });
  }
  await writeChunked(sheetName, vals);
  await Excel.run(async function (ctx) {
    ctx.workbook.worksheets.getItem(sheetName)
      .getRangeByIndexes(0, 0, vals.length, 2).format.autofitColumns();
    await ctx.sync();
  });
}

/* 値を CHUNK 行ずつ書き込む */
async function writeChunked(sheetName, values) {
  if (!values.length) return;
  var cols = values[0].length;
  var done = 0;
  while (done < values.length) {
    var slice = values.slice(done, done + CHUNK);
    var at = done;
    await Excel.run(async function (ctx) {
      ctx.workbook.worksheets.getItem(sheetName)
        .getRangeByIndexes(at, 0, slice.length, cols).values = slice;
      await ctx.sync();
    });
    done += slice.length;
  }
}

/* ---------- 集計 ---------- */
/* group-mode: "ext"（取引先 › 拡張子）/ "folder"（取引先 › フォルダ › 拡張子） */

function currentFilters() {
  return {
    kind: byId("kind").value,
    drop: byId("drop-excluded").checked,
    mode: byId("group-mode").value
  };
}

function groupOf(f) {
  var d = f.assignDepth;
  if (d > 0) return f.folder[d - 1];
  return f.folder.length ? f.folder[0] : DIRECT;
}

function aggregate() {
  var opt = currentFilters();
  var flat = opt.mode === "ext";

  var vendors = new Map();
  var extTotals = new Map();
  var dropped = 0;
  var files = 0;
  var grand = 0;

  for (var i = 0; i < state.rows.length; i++) {
    var f = state.rows[i];
    var v = opt.kind === "real" ? f.real : f.total;
    if (f.excluded) {
      dropped += v;
      if (opt.drop) continue;
    }
    var vName = f.vendor || UNASSIGNED;
    var gName = flat ? null : groupOf(f);

    var ve = vendors.get(vName);
    if (!ve) { ve = { name: vName, value: 0, files: 0, groups: new Map() }; vendors.set(vName, ve); }
    ve.value += v; ve.files++;

    var ge = ve.groups.get(gName);
    if (!ge) { ge = { name: gName, value: 0, files: 0, exts: new Map() }; ve.groups.set(gName, ge); }
    ge.value += v; ge.files++;

    var eb = ge.exts.get(f.ext);
    if (!eb) { eb = { value: 0, files: 0 }; ge.exts.set(f.ext, eb); }
    eb.value += v; eb.files++;

    var et = extTotals.get(f.ext);
    if (!et) { et = { value: 0, files: 0 }; extTotals.set(f.ext, et); }
    et.value += v; et.files++;

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
        .map(function (e) { return { name: e[0] || "(なし)", value: e[1].value, files: e[1].files }; })
        .sort(function (a, b) { return b.value - a.value; });
    });
  });

  return {
    kind: opt.kind, drop: opt.drop, mode: opt.mode, flat: flat,
    vendors: list, grand: grand, files: files, dropped: dropped,
    exts: Array.from(extTotals.entries())
      .map(function (e) { return { name: e[0] || "(なし)", value: e[1].value, files: e[1].files }; })
      .sort(function (a, b) { return b.value - a.value; })
  };
}

/* 指定した取引先・フォルダ・拡張子に絞った実ファイルの一覧を返す。
   group / ext は null で「絞らない」を表す。 */
function filesForBucket(vendorName, groupName, extName) {
  var opt = currentFilters();
  var out = [];
  for (var i = 0; i < state.rows.length; i++) {
    var f = state.rows[i];
    if (f.excluded && opt.drop) continue;
    var vName = f.vendor || UNASSIGNED;
    if (vName !== vendorName) continue;
    if (groupName !== null && groupOf(f) !== groupName) continue;
    var extName2 = f.ext || "(なし)";
    if (extName !== null && extName2 !== extName) continue;
    out.push({
      path: f.segs.join("\\"),
      relPath: f.folder.join("\\"),
      name: f.segs[f.segs.length - 1],
      ext: extName2,
      value: opt.kind === "real" ? f.real : f.total,
      excluded: f.excluded
    });
  }
  out.sort(function (a, b) { return b.value - a.value; });
  return out;
}

function renderSummary() {
  var s = aggregate();
  state.summary = s;
  byId("m-steps").textContent = fmt(s.grand);
  byId("m-files").textContent = fmt(s.files);
  byId("m-dropped").textContent = fmt(s.dropped);

  var host = byId("drill");
  if (!s.vendors.length) { host.innerHTML = "<p class=\"empty\">集計対象がありません。</p>"; return; }

  var max = s.vendors[0].value || 1;
  var groupLabel = s.flat ? "" : " › フォルダ";
  var html = ["<div class=\"dl-head\"><span class=\"g\">取引先" + groupLabel + " › 拡張子</span>" +
    "<span class=\"n\">件数 / ステップ</span><span class=\"p\">構成比</span></div>"];

  s.vendors.forEach(function (ve) {
    var open = state.openL1.has(ve.name);
    var pct = s.grand ? ve.value / s.grand * 100 : 0;
    var rest = ve.name === UNASSIGNED;
    html.push(
      "<div class=\"dl-row" + (open ? " is-open" : "") + "\">" +
      "<div class=\"dl-line is-click\" data-l1=\"" + esc(ve.name) + "\"" +
      " role=\"button\" tabindex=\"0\" aria-expanded=\"" + open + "\">" +
      "<span class=\"twisty\" aria-hidden=\"true\">" + (open ? "▼" : "▶") + "</span>" +
      "<span class=\"dl-name\">" + esc(ve.name) + "</span>" +
      "<span class=\"dl-num\">" + fmtPair(ve.files, ve.value) + "</span>" +
      "<span class=\"dl-pct\">" + pct.toFixed(1) + "%</span>" +
      "</div>" +
      bar(ve.value / max * 100, rest, ve.name, null, null) +
      inlineFiles(ve.name, null, null) +
      "</div>"
    );
    if (!open) return;

    if (s.flat) {
      html.push("<div class=\"lvl2\">");
      var el = ve.groupList[0] ? ve.groupList[0].extList : [];
      el.forEach(function (e) { html.push(leafRow(ve.name, null, e)); });
      html.push("</div>");
      return;
    }

    html.push("<div class=\"lvl2\">");
    ve.groupList.forEach(function (ge) {
      var key = ve.name + "\u0000" + ge.name;
      var o2 = state.openL2.has(key);
      var p2 = ve.value ? ge.value / ve.value * 100 : 0;
      html.push(
        "<div class=\"dl-row" + (o2 ? " is-open" : "") + "\">" +
        "<div class=\"dl-line\">" +
        "<button type=\"button\" class=\"twisty is-click\" data-l2=\"" + esc(key) + "\"" +
        " aria-expanded=\"" + o2 + "\" aria-label=\"" + (o2 ? "折りたたむ" : "拡張子内訳を展開") + "\">" +
        (o2 ? "▼" : "▶") + "</button>" +
        "<span class=\"dl-name mono is-click\" data-vendor=\"" + esc(ve.name) + "\" data-group=\"" + esc(ge.name) +
        "\" data-ext=\"\" role=\"button\" tabindex=\"0\" title=\"クリックでファイル一覧\">" + esc(ge.name) + "</span>" +
        "<span class=\"dl-num\">" + fmtPair(ge.files, ge.value) + "</span>" +
        "<span class=\"dl-pct\">" + p2.toFixed(1) + "%</span>" +
        "</div>" +
        bar(p2, rest, ve.name, ge.name, null) +
        inlineFiles(ve.name, ge.name, null) +
        "</div>"
      );
      if (o2) {
        html.push("<div class=\"lvl3\">");
        ge.extList.forEach(function (e) { html.push(leafRow(ve.name, ge.name, e)); });
        html.push("</div>");
      }
    });
    html.push("</div>");
  });

  html.push("<div class=\"dl-total\"><span class=\"g\">合計</span><span class=\"n\">" +
    fmtPair(s.files, s.grand) + "</span><span class=\"p\"></span></div>");
  host.innerHTML = html.join("");
}

function leafRow(vendorName, groupName, e) {
  var denom;
  if (groupName === null) {
    var ve = state.summary.vendors.filter(function (x) { return x.name === vendorName; })[0];
    denom = ve ? ve.value : 0;
  } else {
    var found = 0;
    state.summary.vendors.some(function (x) {
      if (x.name !== vendorName) return false;
      x.groupList.some(function (g) { if (g.name === groupName) found = g.value; });
      return true;
    });
    denom = found;
  }
  var p = denom ? e.value / denom * 100 : 0;
  return "<div class=\"dl-row\">" +
    "<div class=\"dl-line is-click\" data-vendor=\"" + esc(vendorName) + "\" data-group=\"" +
    esc(groupName === null ? "" : groupName) + "\" data-ext=\"" + esc(e.name) + "\">" +
    "<span class=\"twisty is-leaf\" aria-hidden=\"true\"></span>" +
    "<span class=\"dl-name mono\">" + esc(e.name) + "</span>" +
    "<span class=\"dl-num\">" + fmtPair(e.files, e.value) + "</span>" +
    "<span class=\"dl-pct\">" + p.toFixed(1) + "%</span>" +
    "</div>" +
    bar(p, false, vendorName, groupName, e.name) +
    inlineFiles(vendorName, groupName, e.name) +
    "</div>";
}

function bar(pct, rest, vendorName, groupName, extName) {
  var w = Math.max(0, Math.min(100, pct));
  return "<div class=\"bar" + (rest ? " is-rest" : "") + "\" role=\"button\" tabindex=\"0\"" +
    " data-vendor=\"" + esc(vendorName) + "\" data-group=\"" + esc(groupName === null || groupName === undefined ? "" : groupName) +
    "\" data-ext=\"" + esc(extName === null || extName === undefined ? "" : extName) +
    "\" aria-label=\"該当ファイルを表示\">" +
    "<i style=\"width:" + w.toFixed(1) + "%\"></i></div>";
}

/* 集計の各バケットを識別するキー。HTML属性には出さない（Set の管理専用）ので任意の文字が使える。 */
function fileKey(vendorName, groupName, extName) {
  return JSON.stringify([vendorName, groupName, extName]);
}

/* グラフ／フォルダ名／拡張子行のクリックで開閉する、行のすぐ下に差し込むファイル一覧 */
function inlineFiles(vendorName, groupName, extName) {
  var key = fileKey(vendorName, groupName, extName);
  if (!state.openFiles.has(key)) return "";

  var rows = filesForBucket(vendorName, groupName, extName);
  var sum = rows.reduce(function (s, r) { return s + r.value; }, 0);
  var html = ["<div class=\"file-list-inline\">", "<div class=\"fl-meta\">" + fmtPair(rows.length, sum) + "</div>"];
  if (!rows.length) {
    html.push("<p class=\"empty\">該当するファイルがありません。</p>");
  } else {
    rows.forEach(function (r) {
      html.push(
        "<div class=\"file-row" + (r.excluded ? " is-excluded" : "") + "\">" +
        "<div class=\"file-line1\">" +
        "<span class=\"file-name mono\" title=\"" + esc(r.name) + "\">" + esc(r.name) + "</span>" +
        "<span class=\"file-num\">" + fmt(r.value) + "</span>" +
        "</div>" +
        "<div class=\"file-line2\">" +
        "<span class=\"file-path mono\" title=\"" + esc(r.relPath) + "\">" + esc(r.relPath) + "</span>" +
        "<span class=\"file-ext\">" + esc(r.ext) + "</span>" +
        "</div>" +
        "</div>"
      );
    });
  }
  html.push("</div>");
  return html.join("");
}

function onDrillClick(ev) {
  var fb = ev.target.closest("[data-vendor]");
  if (fb) {
    ev.stopPropagation();
    var vendorName = fb.dataset.vendor;
    var groupName = fb.dataset.group === "" ? null : fb.dataset.group;
    var extName = fb.dataset.ext === "" ? null : fb.dataset.ext;
    var key = fileKey(vendorName, groupName, extName);
    if (state.openFiles.has(key)) state.openFiles.delete(key); else state.openFiles.add(key);
    renderSummary();
    return;
  }
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

/* ---------- 出力 ---------- */

async function exportSheets() {
  var s = state.summary || aggregate();
  var label = s.kind === "real" ? "実ステップ" : "総ステップ";

  var v = [["取引先名", "ファイル数", label, "構成比"]];
  s.vendors.forEach(function (ve) {
    v.push([ve.name, ve.files, ve.value, s.grand ? ve.value / s.grand : 0]);
  });
  v.push(["合計", s.files, s.grand, 1]);

  var x = [["拡張子", "ファイル数", label, "構成比"]];
  s.exts.forEach(function (e) {
    x.push([e.name, e.files, e.value, s.grand ? e.value / s.grand : 0]);
  });

  var c = [["取引先名", "フォルダ", "拡張子", "ファイル数", label]];
  s.vendors.forEach(function (ve) {
    ve.groupList.forEach(function (ge) {
      ge.extList.forEach(function (e) {
        c.push([ve.name, ge.name === null ? "" : ge.name, e.name, e.files, e.value]);
      });
    });
  });

  try {
    setStatus("出力中…");
    await putSheet(SHEET.outVendor, v, [3]);
    await putSheet(SHEET.outExt, x, [3]);
    await putSheet(SHEET.outCross, c, []);
    await Excel.run(async function (ctx) {
      ctx.workbook.worksheets.getItem(SHEET.outVendor).activate();
      await ctx.sync();
    });
    toast("3 シートに出力しました。");
  } catch (e) {
    toast("出力に失敗しました：" + describe(e), true);
  }
}

async function putSheet(name, values, pctCols) {
  await Excel.run(async function (ctx) {
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    if (!sheets.items.some(function (i) { return i.name === name; })) {
      sheets.add(name);
      await ctx.sync();
    }
  });

  var dim = await areaSize(name);
  if (dim && dim.rows) {
    await Excel.run(async function (ctx) {
      ctx.workbook.worksheets.getItem(name)
        .getRangeByIndexes(0, 0, dim.top + dim.rows, Math.max(dim.cols, 8))
        .clear(Excel.ClearApplyTo.all);
      await ctx.sync();
    });
  }

  await writeChunked(name, values);

  await Excel.run(async function (ctx) {
    var s = ctx.workbook.worksheets.getItem(name);
    var w = values[0].length;
    s.getRangeByIndexes(0, 0, 1, w).format.font.bold = true;
    var n = values.length - 1;
    if (n > 0) {
      var pf = [];
      for (var i = 0; i < n; i++) pf.push(["0.0%"]);
      pctCols.forEach(function (ci) {
        s.getRangeByIndexes(1, ci, n, 1).numberFormat = pf;
      });
    }
    s.getRangeByIndexes(0, 0, Math.min(values.length, 200), w).format.autofitColumns();
    await ctx.sync();
  });
}

/* ---------- ユーティリティ ---------- */

function byId(id) { return document.getElementById(id); }
function num(v) { return typeof v === "number" ? v : (parseFloat(String(v).replace(/,/g, "")) || 0); }
function fmt(n) { return Math.round(n).toLocaleString("ja-JP"); }
function fmtPair(files, steps) { return fmt(files) + " / " + fmt(steps); }
function uniq(a) { return Array.from(new Set(a.filter(function (x) { return x; }))); }
function cmpJa(a, b) { return String(a).localeCompare(String(b), "ja"); }
function esc(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c];
  });
}
function setStatus(t) { byId("status-count").textContent = t; }

/* Office API のエラーは message だけでは原因が分からないので debugInfo を添える */
function describe(e) {
  if (!e) return "不明なエラー";
  var msg = e.message || String(e);
  if (e.code) msg = e.code + ": " + msg;
  if (e.debugInfo) {
    var d = e.debugInfo;
    var extra = [d.errorLocation, d.statement, d.surroundingStatements]
      .filter(function (x) { return x; }).join(" / ");
    if (extra) msg += "（" + extra + "）";
  }
  return msg;
}

var toastTimer = null;
function toast(msg, isError) {
  var el = byId("toast");
  el.textContent = msg;
  el.classList.toggle("is-error", !!isError);
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.hidden = true; }, 4000);
}
