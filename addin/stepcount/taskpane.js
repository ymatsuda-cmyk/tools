/* ステップ数集計アドイン
   データ  : パス | 拡張子 | 総ステップ | 実ステップ
   割当    : フォルダパス | 取引先名
   除外    : フォルダパス | メモ
   割当・除外はいずれも「最長一致した祖先が勝つ」で解決する。
   ステップ数は常に実ステップを使用する（基準の切替は廃止）。 */

var SHEET = {
  data: "データ",
  assign: "割当",
  code: "コード",
  outVendor: "集計_取引先",
  outExt: "集計_拡張子",
  outCross: "集計_クロス"
};

var UNASSIGNED = "（未割当）";
var DIRECT = "（直下）";
var UNASSIGN_MARK = "（解除）";

var CARRY_UNC_ROOT = false;
var CHUNK = 2000;

var state = {
  rows: [],
  tree: null,
  nodeIndex: new Map(),
  assignMap: new Map(),
  vendors: [],
  codeVendors: [],
  screenExts: new Set(),
  configExts: new Set(),
  tab: "dashboard",
  openNodes: new Set(),
  srcVendor: "すべて",
  srcExt: "すべて",
  srcSort: { col: "folder", dir: "asc" },
  category: "total",
  editVendor: "",
  draftChecked: new Set()
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

  byId("depth-open").addEventListener("change", function () {
    state.openNodes = new Set();
    expandTo(parseInt(this.value, 10));
    renderTree();
  });
  byId("btn-add-vendor").addEventListener("click", addNewVendor);
  byId("new-vendor").addEventListener("keydown", function (ev) {
    if (ev.key === "Enter") { ev.preventDefault(); addNewVendor(); }
  });
  byId("btn-focus").addEventListener("click", focusChecked);
  byId("btn-save").addEventListener("click", saveAssignments);
  byId("btn-discard").addEventListener("click", discardDraft);

  byId("btn-export").addEventListener("click", exportSheets);
  byId("src-head").addEventListener("click", function (ev) {
    var b = ev.target.closest("[data-sort]");
    if (b) setSourceSort(b.dataset.sort);
  });

  byId("tree").addEventListener("click", onTreeClick);
}

function switchTab(name) {
  state.tab = name;
  document.querySelectorAll(".tab").forEach(function (b) {
    var on = b.dataset.tab === name;
    b.classList.toggle("is-active", on);
    b.setAttribute("aria-selected", on ? "true" : "false");
  });
  ["dashboard", "source", "assign"].forEach(function (n) {
    byId("panel-" + n).hidden = n !== name;
  });
  if (name === "dashboard") {
    state.srcExt = "すべて";
    renderDashboard();
  }
  if (name === "source") renderSourceList();
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
    if (!state.editVendor) {
      var order = assignPillOptions();
      state.editVendor = order.length ? order[0] : "";
    }
    state.draftChecked = committedSetFor(state.editVendor);
    renderAssignPills();
    renderTree();
    if (state.tab === "dashboard") renderDashboard();
    if (state.tab === "source") renderSourceList();
    setStatus(fmt(state.rows.length) + " ファイル / 割当 " + state.assignMap.size +
      " 件");
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
  var code = await readArea(SHEET.code, 3);
  return { data: data, assign: assign, code: code };
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
    if (names.indexOf(SHEET.assign) < 0) {
      var s = sheets.add(SHEET.assign);
      s.getRange("A1:B1").values = [["フォルダパス", "取引先名"]];
      added = true;
    }
    if (names.indexOf(SHEET.code) < 0) {
      var cs = sheets.add(SHEET.code);
      cs.getRange("A1:C1").values = [["取引先", "画面", "定義"]];
      added = true;
    }
    if (added) await ctx.sync();
  });
}

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
    return null;
  }
}

function buildModel(raw) {
  state.assignMap = pairsToMap(raw.assign);

  var code = parseCodeSheet(raw.code);
  state.codeVendors = code.vendors;
  state.screenExts = code.screenExts;
  state.configExts = code.configExts;

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
      real: num(r[3]),
      isScreen: state.screenExts.has(ext),
      isConfig: state.configExts.has(ext)
    });
  }

  for (var j = 0; j < rows.length; j++) {
    var f = rows[j];
    var hitA = lookup(state.assignMap, f.folder);
    f.vendor = (hitA && hitA.value !== UNASSIGN_MARK) ? hitA.value : null;
    f.assignDepth = hitA ? hitA.depth : 0;
  }

  state.rows = rows;
  state.nodeIndex = new Map();
  state.tree = buildTree(rows);
  state.vendors = uniq(Array.from(state.assignMap.values())
    .filter(function (v) { return v !== UNASSIGN_MARK; })).sort(cmpJa);
}

/* コードシート（取引先・画面）を読む。2列は独立したリストとして扱う
   （行が対応している必要はない。どちらかが空でももう一方は読む）。 */
function parseCodeSheet(values) {
  var vendors = [], screenExts = new Set(), configExts = new Set();
  if (!values) return { vendors: vendors, screenExts: screenExts, configExts: configExts };
  var start = 0;
  if (values[0] && (String(values[0][0]) === "取引先" || String(values[0][1]) === "画面" || String(values[0][2]) === "定義")) start = 1;
  var seen = new Set();
  for (var i = start; i < values.length; i++) {
    var r = values[i];
    if (!r) continue;
    var v = r[0];
    if (v !== null && v !== undefined && String(v).trim() !== "") {
      var name = String(v).trim();
      if (!seen.has(name)) { seen.add(name); vendors.push(name); }
    }
    var e = r[1];
    if (e !== null && e !== undefined && String(e).trim() !== "") {
      screenExts.add(String(e).trim().toLowerCase().replace(/^\./, ""));
    }
    var c = r[2];
    if (c !== null && c !== undefined && String(c).trim() !== "") {
      configExts.add(String(c).trim().toLowerCase().replace(/^\./, ""));
    }
  }
  return { vendors: vendors, screenExts: screenExts, configExts: configExts };
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

/* ---------- ツリー（割当タブ） ---------- */

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
    vendor: null, inherited: null, partial: false, overridden: false
  };
}

function markTree(root) {
  walk(root, null);
  function walk(n, inheritedVendor) {
    var raw = state.assignMap.has(n.path) ? state.assignMap.get(n.path) : null;
    var isOverride = raw === UNASSIGN_MARK;
    var own = (raw && !isOverride) ? raw : null;
    n.vendor = own;
    n.overridden = isOverride;
    n.inherited = (own || isOverride) ? null : inheritedVendor;
    var vend = own || (isOverride ? null : inheritedVendor);
    var any = false;
    n.kids.forEach(function (k) {
      walk(k, vend);
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

/* 選択中（チェック済み）のフォルダへ辿り着く経路だけを開き、それ以外は閉じる。 */
function focusChecked() {
  var open = new Set();
  state.draftChecked.forEach(function (path) {
    var segs = path.split("\\");
    for (var i = 1; i <= segs.length; i++) open.add(segs.slice(0, i).join("\\"));
  });
  state.openNodes = open;
  renderTree();
}

/* ツリー全体を1回辿り、各フォルダが実際に集計上どの取引先になるか（祖先からの継承込み）を求める。
   ダッシュボード・ソース一覧の集計ロジック（lookup方式）と同じ考え方をチェックボックスにも適用し、
   画面の状態と実際の集計を一致させる。 */
function computeEffectiveVendors() {
  var map = new Map();
  (function walk(n, inherited) {
    var raw = state.assignMap.has(n.path) ? state.assignMap.get(n.path) : null;
    var isOverride = raw === UNASSIGN_MARK;
    var own = (raw && !isOverride) ? raw : null;
    var effective = own || (isOverride ? null : inherited);
    map.set(n.path, effective);
    n.kids.forEach(function (k) { walk(k, effective); });
  })(state.tree, null);
  return map;
}

/* 現在 vendor に実際に属している（直接または継承で）パスの集合。 */
function committedSetFor(vendor) {
  var set = new Set();
  if (!vendor) return set;
  computeEffectiveVendors().forEach(function (v, path) {
    if (v === vendor) set.add(path);
  });
  return set;
}

function switchEditVendor(vendor) {
  state.editVendor = vendor;
  state.draftChecked = committedSetFor(vendor);
  renderAssignPills();
  renderTree();
}

function discardDraft() {
  state.draftChecked = committedSetFor(state.editVendor);
  renderTree();
}

function toggleDraft(path) {
  var n = state.nodeIndex.get(path);
  if (!n) return;
  var next = !state.draftChecked.has(path);
  (function walk(nd) {
    if (next) state.draftChecked.add(nd.path); else state.draftChecked.delete(nd.path);
    nd.kids.forEach(walk);
  })(n);
  renderTree();
}

function draftDiff() {
  var committed = committedSetFor(state.editVendor);
  var added = [], removed = [];
  state.draftChecked.forEach(function (p) { if (!committed.has(p)) added.push(p); });
  committed.forEach(function (p) { if (!state.draftChecked.has(p)) removed.push(p); });
  return { added: added, removed: removed };
}

function assignPillOptions() {
  return aggregateAll("total").vendors.map(function (v) { return v.name; });
}

function renderAssignPills() {
  pillRow("pav", assignPillOptions(), state.editVendor, function (v) { switchEditVendor(v); });
}

async function addNewVendor() {
  var input = byId("new-vendor");
  var name = input.value.trim();
  if (!name) { toast("取引先名を入力してください。", true); return; }
  if (state.codeVendors.indexOf(name) >= 0) {
    toast("その取引先はすでに登録されています。", true);
    switchEditVendor(name);
    return;
  }
  try {
    setStatus("追加中…");
    await appendCodeVendor(name);
    state.codeVendors.push(name);
    input.value = "";
    switchEditVendor(name);
    setStatus(fmt(state.rows.length) + " ファイル / 割当 " + state.assignMap.size +
      " 件");
    toast("取引先「" + name + "」を追加しました。");
  } catch (e) {
    toast("追加に失敗しました：" + describe(e), true);
  }
}

/* コードシートのA列（取引先）の末尾に1件だけ追記する。B列（画面）には触れない。 */
async function appendCodeVendor(name) {
  var dim = await areaSize(SHEET.code);
  var nextRow = dim ? dim.top + dim.rows : 1;
  await Excel.run(async function (ctx) {
    ctx.workbook.worksheets.getItem(SHEET.code)
      .getRangeByIndexes(nextRow, 0, 1, 1).values = [[name]];
    await ctx.sync();
  });
}

function renderTree() {
  var host = byId("tree");
  var keepScroll = host.scrollTop;
  var kids = sortKids(state.tree);
  if (!kids.length) { host.innerHTML = "<p class=\"empty\">データがありません。</p>"; return; }

  var diff = draftDiff();
  var diffLabel = (diff.added.length || diff.removed.length)
    ? ("未保存の変更 " + (diff.added.length + diff.removed.length) + " 件"
      + "（追加 " + diff.added.length + " / 解除 " + diff.removed.length + "）")
    : "変更なし";
  byId("diff-count").textContent = diffLabel;

  var effective = computeEffectiveVendors();

  var html = [];
  kids.forEach(function (k) { emit(k, 0); });
  host.innerHTML = html.join("");
  host.scrollTop = keepScroll;

  function emit(n, depth) {
    var open = state.openNodes.has(n.path);
    var checked = state.draftChecked.has(n.path);
    var owner = effective.get(n.path);
    var badge = "";
    if (checked) {
      badge = "<span class=\"badge badge-vendor\">" + esc(state.editVendor) + "</span>";
    } else if (owner && owner !== state.editVendor) {
      badge = "<span class=\"badge badge-inherit\">" + esc(owner) + "</span>";
    }

    html.push(
      "<div class=\"node" + ((owner && owner !== state.editVendor && !checked) ? " is-off" : "") +
      "\" style=\"padding-left:" + (depth * 14 + 4) + "px\">" +
      "<button class=\"twisty" + (n.kids.size ? "" : " is-leaf") + "\" data-toggle=\"" + esc(n.path) + "\"" +
      " aria-label=\"" + (open ? "折りたたむ" : "展開する") + "\">" + (open ? "▼" : "▶") + "</button>" +
      "<input type=\"checkbox\" data-pick=\"" + esc(n.path) + "\"" + (checked ? " checked" : "") + ">" +
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
  if (c) toggleDraft(c.dataset.pick);
}

/* ---------- 割当の保存 ---------- */

async function saveAssignments() {
  var vendor = state.editVendor;
  if (!vendor) { toast("編集する取引先を選択または入力してください。", true); return; }

  var diff = draftDiff();
  if (!diff.added.length && !diff.removed.length) { toast("変更がありません。", true); return; }

  diff.added.forEach(function (p) { state.assignMap.set(normKey(p), vendor); });
  diff.removed.forEach(function (p) { state.assignMap.delete(normKey(p)); });

  /* 追加・削除を反映した直後のツリー全体を見直し、チェックしていないのに
     祖先からの継承で editVendor になってしまうフォルダがあれば、そこだけ継承を断ち切る。
     （新しく上位フォルダへ割り当てたことで、チェックを外していた配下に継承が漏れるケースを防ぐ） */
  var leaked = [];
  computeEffectiveVendors().forEach(function (v, path) {
    if (v === vendor && !state.draftChecked.has(path)) {
      state.assignMap.set(normKey(path), UNASSIGN_MARK);
      leaked.push(path);
    }
  });

  try {
    setStatus("保存中…");
    await writeMap(SHEET.assign, ["フォルダパス", "取引先名"], state.assignMap);
    refreshAssignments(diff.added.concat(diff.removed).concat(leaked));
    toast("追加 " + diff.added.length + " 件 / 解除 " + diff.removed.length + " 件を保存しました。" +
      (leaked.length ? "（継承を断ち切ったフォルダ " + leaked.length + " 件）" : ""));
  } catch (e) {
    toast("書き込みに失敗しました：" + describe(e), true);
  }
}

/* 割当を変えたあと、Excel を読み直さずにメモリ上のモデルだけ更新する。
   changed に含まれるパスの配下だけ再解決すればよいので、行数が多くても軽い。 */
function refreshAssignments(changed) {
  var prefixes = changed.map(function (p) { return normKey(p); });

  for (var i = 0; i < state.rows.length; i++) {
    var f = state.rows[i];
    var full = f.folder.join("\\");
    var affected = false;
    for (var j = 0; j < prefixes.length; j++) {
      var p = prefixes[j];
      if (full === p || full.indexOf(p + "\\") === 0) { affected = true; break; }
    }
    if (!affected) continue;
    var hitA = lookup(state.assignMap, f.folder);
    f.vendor = (hitA && hitA.value !== UNASSIGN_MARK) ? hitA.value : null;
    f.assignDepth = hitA ? hitA.depth : 0;
  }

  markTree(state.tree);

  state.vendors = uniq(Array.from(state.assignMap.values())
    .filter(function (v) { return v !== UNASSIGN_MARK; })).sort(cmpJa);

  state.draftChecked = committedSetFor(state.editVendor);
  renderAssignPills();
  renderTree();
  setStatus(fmt(state.rows.length) + " ファイル / 割当 " + state.assignMap.size +
    " 件");
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

/* ---------- 共通：割当フォルダ名 ---------- */

function groupOf(f) {
  var d = f.assignDepth;
  if (d > 0) return f.folder[d - 1];
  return f.folder.length ? f.folder[0] : DIRECT;
}

/* ---------- 共通：合計／ソースコード／画面定義／設定 の区分 ---------- */

var CATEGORIES = [
  { key: "total", label: "合計" },
  { key: "source", label: "ソースコード" },
  { key: "screen", label: "画面定義" },
  { key: "config", label: "設定" }
];

function fileCategory(f) {
  if (f.isScreen) return "screen";
  if (f.isConfig) return "config";
  return "source";
}

function extCategory(ext) {
  if (state.screenExts.has(ext)) return "screen";
  if (state.configExts.has(ext)) return "config";
  return "source";
}

/* 取引先が割り当てられているファイルだけを対象に、区分ごとのファイル数を数える。
   4つのボタンに表示する参考値（現在の選択には左右されない）。 */
function categoryCounts(vendorFilter) {
  var counts = { total: 0, source: 0, screen: 0, config: 0 };
  state.rows.forEach(function (f) {
    if (!f.vendor) return;
    if (vendorFilter && vendorFilter !== "すべて" && f.vendor !== vendorFilter) return;
    counts.total++;
    counts[fileCategory(f)]++;
  });
  return counts;
}

function renderCategoryButtons(hostId, vendorFilter) {
  var counts = categoryCounts(vendorFilter);
  var host = byId(hostId);
  host.innerHTML = CATEGORIES.map(function (c) {
    var on = state.category === c.key;
    return "<button type=\"button\" class=\"cat-btn" + (on ? " is-active" : "") + "\" data-cat=\"" + c.key + "\">" +
      "<span class=\"cat-label\">" + esc(c.label) + "</span>" +
      "<span class=\"cat-count\">" + fmt(counts[c.key]) + "</span></button>";
  }).join("");
  host.onclick = function (ev) {
    var b = ev.target.closest("[data-cat]");
    if (!b) return;
    setCategory(b.dataset.cat);
  };
}

function setCategory(cat) {
  state.category = cat;
  if (cat !== "total" && state.srcExt !== "すべて" && extCategory(state.srcExt) !== cat) {
    state.srcExt = "すべて";
  }
  renderDashboard();
  renderSourceList();
}
/* ---------- ダッシュボード ---------- */

function aggregateAll(category) {
  var vendors = new Map();
  state.codeVendors.forEach(function (name) { vendors.set(name, { files: 0, steps: 0 }); });

  var extsAssigned = new Map();

  state.rows.forEach(function (f) {
    if (category !== "total" && fileCategory(f) !== category) return;
    var vName = f.vendor || UNASSIGNED;
    var ve = vendors.get(vName);
    if (!ve) { ve = { files: 0, steps: 0 }; vendors.set(vName, ve); }
    ve.files++; ve.steps += f.real;
    if (f.vendor) {
      extsAssigned.set(f.ext, (extsAssigned.get(f.ext) || 0) + f.real);
    }
  });

  var vList = Array.from(vendors.entries()).map(function (e) {
    return { name: e[0], files: e[1].files, steps: e[1].steps };
  }).filter(function (v) { return v.name !== UNASSIGNED; });
  vList.sort(function (a, b) { return b.steps - a.steps || cmpJa(a.name, b.name); });

  var grandAssigned = vList.reduce(function (s, v) { return s + v.steps; }, 0);
  var totalFiles = vList.reduce(function (s, v) { return s + v.files; }, 0);
  var vendorCount = vList.filter(function (v) { return v.files > 0; }).length;

  var extList = Array.from(extsAssigned.entries())
    .map(function (e) { return { name: e[0] || "(なし)", steps: e[1] }; })
    .sort(function (a, b) { return b.steps - a.steps; });

  return {
    vendors: vList, grandAssigned: grandAssigned, vendorCount: vendorCount,
    totalFiles: totalFiles, exts: extList
  };
}

function renderDashboard() {
  renderCategoryButtons("dash-cats");
  var d = aggregateAll(state.category);

  byId("dash-metrics").innerHTML =
    metricCard("合計ステップ", fmt(d.grandAssigned), "未割当を除く") +
    metricCard("取引先数", fmt(d.vendorCount), null);

  var vMax = d.vendors.length ? Math.max.apply(null, d.vendors.map(function (v) { return v.steps; })) || 1 : 1;
  byId("dash-vendors").innerHTML = d.vendors.length
    ? d.vendors.map(function (v) { return statRow(v.name, v.files, v.steps, vMax, false); }).join("")
    : "<p class=\"empty\">データがありません。</p>";

  var eMax = d.exts.length ? d.exts[0].steps || 1 : 1;
  byId("dash-exts").innerHTML = d.exts.length
    ? d.exts.map(function (e) { return statRow(e.name, null, e.steps, eMax, false); }).join("")
    : "<p class=\"empty\">データがありません。</p>";
}

function metricCard(label, value, note) {
  return "<div class=\"metric\"><span class=\"metric-label\">" + esc(label) + "</span>" +
    "<span class=\"metric-value\">" + value + "</span>" +
    (note ? "<span class=\"metric-note\">" + esc(note) + "</span>" : "") + "</div>";
}

function statRow(name, files, steps, max, rest) {
  var pct = max ? steps / max * 100 : 0;
  return "<div class=\"stat-row" + (rest ? " is-rest" : "") + "\">" +
    "<div class=\"stat-line\"><span class=\"stat-name\">" + esc(name) + "</span>" +
    "<span class=\"stat-num\">" + (files !== null ? fmt(files) + "件 / " : "") + fmt(steps) + "</span></div>" +
    "<div class=\"stat-bar\"><i style=\"width:" + pct.toFixed(1) + "%\"></i></div></div>";
}

/* ---------- ソース一覧（取引先 › 拡張子のピルで絞り込み） ---------- */

/* コードシートの取引先一覧と、実際に割当シートで使われている取引先名を統合した一覧。 */
function allVendorNames() {
  var set = new Set(state.codeVendors);
  state.assignMap.forEach(function (v) { if (v && v !== UNASSIGN_MARK) set.add(v); });
  return Array.from(set);
}

function vendorOptions() {
  return ["すべて"].concat(allVendorNames().sort(cmpJa));
}

function extOptions(vendorSel) {
  var set = [];
  state.rows.forEach(function (f) {
    if (!f.vendor) return;
    if (vendorSel !== "すべて" && f.vendor !== vendorSel) return;
    var e = f.ext || "(なし)";
    if (set.indexOf(e) < 0) set.push(e);
  });
  return ["すべて"].concat(set.sort(cmpJa));
}

function filteredSourceRows() {
  return state.rows.filter(function (f) {
    if (!f.vendor) return false;
    if (state.category !== "total" && fileCategory(f) !== state.category) return false;
    if (state.srcVendor !== "すべて" && f.vendor !== state.srcVendor) return false;
    var e = f.ext || "(なし)";
    if (state.srcExt !== "すべて" && e !== state.srcExt) return false;
    return true;
  }).map(function (f) {
    return {
      folder: f.folder.join("\\"),
      name: f.segs[f.segs.length - 1],
      ext: f.ext || "(なし)",
      steps: f.real
    };
  });
}

function sortSourceRows(rows) {
  var col = state.srcSort.col, dir = state.srcSort.dir === "desc" ? -1 : 1;
  var sorted = rows.slice();
  sorted.sort(function (a, b) {
    var cmp;
    if (col === "steps") cmp = a.steps - b.steps;
    else if (col === "folder") cmp = cmpJa(a.folder, b.folder) || cmpJa(a.name, b.name);
    else if (col === "name") cmp = cmpJa(a.name, b.name);
    else cmp = cmpJa(a.ext, b.ext) || cmpJa(a.name, b.name);
    return cmp * dir;
  });
  return sorted;
}

function setSourceSort(col) {
  if (state.srcSort.col === col) {
    state.srcSort.dir = state.srcSort.dir === "asc" ? "desc" : "asc";
  } else {
    state.srcSort = { col: col, dir: "asc" };
  }
  renderSourceList();
}

function sortIndicator(col) {
  if (state.srcSort.col !== col) return "";
  return state.srcSort.dir === "asc" ? " ▲" : " ▼";
}

function renderSourceList() {
  renderCategoryButtons("src-cats", state.srcVendor);

  pillRow("pv", vendorOptions(), state.srcVendor, function (v) {
    state.srcVendor = v; state.srcExt = "すべて"; renderSourceList();
  });
  pillRow("pe", extOptions(state.srcVendor), state.srcExt, function (v) {
    state.srcExt = v; renderSourceList();
  }, function (ext) { return state.category !== "total" && extCategory(ext) !== state.category; });

  var rows = sortSourceRows(filteredSourceRows());
  var sum = rows.reduce(function (s, r) { return s + r.steps; }, 0);

  byId("src-metrics").innerHTML =
    metricCard("ステップ", fmt(sum), null) + metricCard("ファイル数", fmt(rows.length), null);

  byId("src-head").innerHTML =
    "<button type=\"button\" class=\"src-col folder\" data-sort=\"folder\">フォルダ" + sortIndicator("folder") + "</button>" +
    "<button type=\"button\" class=\"src-col name\" data-sort=\"name\">ファイル名" + sortIndicator("name") + "</button>" +
    "<button type=\"button\" class=\"src-col ext\" data-sort=\"ext\">拡張子" + sortIndicator("ext") + "</button>" +
    "<button type=\"button\" class=\"src-col steps\" data-sort=\"steps\">ステップ数" + sortIndicator("steps") + "</button>";

  var host = byId("src-rows");
  host.innerHTML = rows.length ? rows.map(function (r) {
    return "<div class=\"src-row\">" +
      "<span class=\"src-col folder mono\" title=\"" + esc(r.folder) + "\">" + esc(r.folder) + "</span>" +
      "<span class=\"src-col name\" title=\"" + esc(r.name) + "\">" + esc(r.name) + "</span>" +
      "<span class=\"src-col ext mono\">" + esc(r.ext) + "</span>" +
      "<span class=\"src-col steps\">" + fmt(r.steps) + "</span></div>";
  }).join("") : "<p class=\"empty\">該当するファイルがありません。</p>";
}

function pillRow(hostId, options, selected, onPick, isDisabled) {
  var host = byId(hostId);
  host.innerHTML = options.map(function (o) {
    var disabled = isDisabled ? isDisabled(o) : false;
    var on = o === selected && !disabled;
    var rest = o === UNASSIGNED;
    return "<button type=\"button\" class=\"pill" + (on ? " is-active" : "") + (rest ? " is-rest" : "") +
      (disabled ? " is-disabled" : "") + "\"" + (disabled ? " disabled" : "") +
      " data-v=\"" + esc(o) + "\">" + esc(o) + "</button>";
  }).join("");
  host.onclick = function (ev) {
    var b = ev.target.closest("[data-v]");
    if (!b || b.disabled) return;
    onPick(b.dataset.v);
  };
}

/* ---------- 出力（ソース一覧の除外設定を使い、全件を集計） ---------- */

function aggregateFull() {
  var vendors = new Map();
  var extTotals = new Map();
  var totalFiles = 0, grand = 0;

  state.rows.forEach(function (f) {
    var v = f.real;
    var vName = f.vendor || UNASSIGNED;
    var gName = groupOf(f);

    var ve = vendors.get(vName);
    if (!ve) { ve = { name: vName, files: 0, steps: 0, groups: new Map() }; vendors.set(vName, ve); }
    ve.files++; ve.steps += v;

    var ge = ve.groups.get(gName);
    if (!ge) { ge = { name: gName, files: 0, steps: 0, exts: new Map() }; ve.groups.set(gName, ge); }
    ge.files++; ge.steps += v;

    var eb = ge.exts.get(f.ext);
    if (!eb) { eb = { files: 0, steps: 0 }; ge.exts.set(f.ext, eb); }
    eb.files++; eb.steps += v;

    var et = extTotals.get(f.ext);
    if (!et) { et = { files: 0, steps: 0 }; extTotals.set(f.ext, et); }
    et.files++; et.steps += v;

    totalFiles++; grand += v;
  });

  var list = Array.from(vendors.values()).sort(function (a, b) {
    if (a.name === UNASSIGNED) return 1;
    if (b.name === UNASSIGNED) return -1;
    return b.steps - a.steps;
  });
  list.forEach(function (ve) {
    ve.groupList = Array.from(ve.groups.values()).sort(function (a, b) { return b.steps - a.steps; });
    ve.groupList.forEach(function (ge) {
      ge.extList = Array.from(ge.exts.entries())
        .map(function (e) { return { name: e[0] || "(なし)", files: e[1].files, steps: e[1].steps }; })
        .sort(function (a, b) { return b.steps - a.steps; });
    });
  });

  return {
    vendors: list, files: totalFiles, grand: grand,
    exts: Array.from(extTotals.entries())
      .map(function (e) { return { name: e[0] || "(なし)", files: e[1].files, steps: e[1].steps }; })
      .sort(function (a, b) { return b.steps - a.steps; })
  };
}

async function exportSheets() {
  var s = aggregateFull();

  var v = [["取引先名", "ファイル数", "ステップ", "構成比"]];
  s.vendors.forEach(function (ve) {
    v.push([ve.name, ve.files, ve.steps, s.grand ? ve.steps / s.grand : 0]);
  });
  v.push(["合計", s.files, s.grand, 1]);

  var x = [["拡張子", "ファイル数", "ステップ", "構成比"]];
  s.exts.forEach(function (e) {
    x.push([e.name, e.files, e.steps, s.grand ? e.steps / s.grand : 0]);
  });

  var c = [["取引先名", "フォルダ", "拡張子", "ファイル数", "ステップ"]];
  s.vendors.forEach(function (ve) {
    ve.groupList.forEach(function (ge) {
      ge.extList.forEach(function (e) {
        c.push([ve.name, ge.name, e.name, e.files, e.steps]);
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
