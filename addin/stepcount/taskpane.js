/* ステップ数集計アドイン
   データ  : パス | 拡張子 | 総ステップ | 実ステップ
   割当    : フォルダパス | 取引先名
   除外    : フォルダパス | メモ
   割当・除外はいずれも「最長一致した祖先が勝つ」で解決する。
   ステップ数は常に実ステップを使用する（基準の切替は廃止）。 */

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
var UNASSIGN_MARK = "（解除）";

var CARRY_UNC_ROOT = false;
var CHUNK = 2000;

var state = {
  rows: [],
  tree: null,
  nodeIndex: new Map(),
  assignMap: new Map(),
  excludeMap: new Map(),
  vendors: [],
  tab: "dashboard",
  openNodes: new Set(),
  checked: new Set(),
  srcVendor: "すべて",
  srcFolder: "すべて",
  srcExt: "すべて"
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
  byId("btn-assign").addEventListener("click", function () {
    var vendor = byId("vendor").value.trim();
    applyAssign(vendor ? "assign" : "unassign");
  });
  byId("btn-exclude").addEventListener("click", function () { applyAssign("exclude"); });
  byId("btn-include").addEventListener("click", function () { applyAssign("include"); });

  byId("drop-excluded").addEventListener("change", renderSourceList);
  byId("btn-export").addEventListener("click", exportSheets);

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
  if (name === "dashboard") renderDashboard();
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
    renderTree();
    if (state.tab === "dashboard") renderDashboard();
    if (state.tab === "source") renderSourceList();
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

/* topPath 配下を辿り、チェックが外れている枝（＝配下ごと明示的に除外された枝）の
   最も浅い経路を集める。チェックされている枝はさらに下まで辿る。 */
function uncheckedBranches(topPath) {
  var root = state.nodeIndex.get(topPath);
  var out = [];
  if (!root) return out;
  (function walk(n) {
    n.kids.forEach(function (k) {
      if (state.checked.has(k.path)) walk(k);
      else out.push(k.path);
    });
  })(root);
  return out;
}

async function applyAssign(mode) {
  var all = Array.from(state.checked);
  if (!all.length) { toast("フォルダを選択してください。", true); return; }

  var vendor = byId("vendor").value.trim();
  if (mode === "assign" && !vendor) { toast("取引先名を入力してください。", true); return; }

  var count = 0;
  var excludedBranches = 0;
  if (mode === "assign") {
    var picks = topmostPicks(all);
    picks.forEach(function (p) {
      state.assignMap.set(normKey(p), vendor);
      var branches = uncheckedBranches(p);
      branches.forEach(function (bp) {
        state.assignMap.set(normKey(bp), UNASSIGN_MARK);
      });
      excludedBranches += branches.length;
    });
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
    toast(count + " フォルダを更新しました。" +
      (excludedBranches ? "（チェックを外した " + excludedBranches + " 件は対象外にしました）" : ""));
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

/* ---------- ダッシュボード ---------- */

function aggregateAll(dropExcluded) {
  var vendors = new Map();
  var extsAssigned = new Map();

  state.rows.forEach(function (f) {
    if (f.excluded && dropExcluded) return;
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
  vList.sort(function (a, b) { return b.steps - a.steps; });

  var grandAssigned = vList.reduce(function (s, v) { return s + v.steps; }, 0);
  var totalFiles = vList.reduce(function (s, v) { return s + v.files; }, 0);
  var vendorCount = vList.length;

  var extList = Array.from(extsAssigned.entries())
    .map(function (e) { return { name: e[0] || "(なし)", steps: e[1] }; })
    .sort(function (a, b) { return b.steps - a.steps; });

  return { vendors: vList, grandAssigned: grandAssigned, vendorCount: vendorCount, totalFiles: totalFiles, exts: extList };
}

function renderDashboard() {
  var d = aggregateAll(true);

  byId("dash-metrics").innerHTML =
    metricCard("合計ステップ", fmt(d.grandAssigned), "未割当を除く") +
    metricCard("ファイル数", fmt(d.totalFiles), "未割当を除く") +
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

/* ---------- ソース一覧（取引先 › 割当フォルダ › 拡張子のピルで絞り込み） ---------- */

function vendorOptions() {
  var names = uniq(state.rows.map(function (f) { return f.vendor; })
    .filter(function (v) { return v; })).sort(cmpJa);
  return ["すべて"].concat(names);
}

function folderOptions(vendorSel, dropExcluded) {
  var set = [];
  state.rows.forEach(function (f) {
    if (!f.vendor) return;
    if (dropExcluded && f.excluded) return;
    if (vendorSel !== "すべて" && f.vendor !== vendorSel) return;
    var g = groupOf(f);
    if (set.indexOf(g) < 0) set.push(g);
  });
  return ["すべて"].concat(set.sort(cmpJa));
}

function extOptions(vendorSel, folderSel, dropExcluded) {
  var set = [];
  state.rows.forEach(function (f) {
    if (!f.vendor) return;
    if (dropExcluded && f.excluded) return;
    if (vendorSel !== "すべて" && f.vendor !== vendorSel) return;
    if (folderSel !== "すべて" && groupOf(f) !== folderSel) return;
    var e = f.ext || "(なし)";
    if (set.indexOf(e) < 0) set.push(e);
  });
  return ["すべて"].concat(set.sort(cmpJa));
}

function filteredSourceRows(dropExcluded) {
  return state.rows.filter(function (f) {
    if (!f.vendor) return false;
    if (dropExcluded && f.excluded) return false;
    if (state.srcVendor !== "すべて" && f.vendor !== state.srcVendor) return false;
    if (state.srcFolder !== "すべて" && groupOf(f) !== state.srcFolder) return false;
    var e = f.ext || "(なし)";
    if (state.srcExt !== "すべて" && e !== state.srcExt) return false;
    return true;
  }).map(function (f) {
    return {
      folder: f.folder.join("\\"),
      name: f.segs[f.segs.length - 1],
      ext: f.ext || "(なし)",
      steps: f.real,
      excluded: f.excluded
    };
  }).sort(function (a, b) { return b.steps - a.steps; });
}

function renderSourceList() {
  var dropExcluded = byId("drop-excluded").checked;

  pillRow("pv", vendorOptions(), state.srcVendor, function (v) {
    state.srcVendor = v; state.srcFolder = "すべて"; state.srcExt = "すべて"; renderSourceList();
  });
  pillRow("pf", folderOptions(state.srcVendor, dropExcluded), state.srcFolder, function (v) {
    state.srcFolder = v; state.srcExt = "すべて"; renderSourceList();
  });
  pillRow("pe", extOptions(state.srcVendor, state.srcFolder, dropExcluded), state.srcExt, function (v) {
    state.srcExt = v; renderSourceList();
  });

  var rows = filteredSourceRows(dropExcluded);
  var sum = rows.reduce(function (s, r) { return s + r.steps; }, 0);
  byId("src-steps").textContent = fmt(sum);
  byId("src-files").textContent = fmt(rows.length);

  var host = byId("src-rows");
  host.innerHTML = rows.length ? rows.map(function (r) {
    return "<div class=\"src-row" + (r.excluded ? " is-excluded" : "") + "\">" +
      "<span class=\"g\"><span class=\"folder\">" + esc(r.folder) + "/</span>" +
      "<span class=\"name\">" + esc(r.name) + "</span></span>" +
      "<span class=\"e\">" + esc(r.ext) + "</span>" +
      "<span class=\"n\">" + fmt(r.steps) + "</span></div>";
  }).join("") : "<p class=\"empty\">該当するファイルがありません。</p>";
}

function pillRow(hostId, options, selected, onPick) {
  var host = byId(hostId);
  host.innerHTML = options.map(function (o) {
    var on = o === selected;
    var rest = o === UNASSIGNED;
    return "<button type=\"button\" class=\"pill" + (on ? " is-active" : "") + (rest ? " is-rest" : "") +
      "\" data-v=\"" + esc(o) + "\">" + esc(o) + "</button>";
  }).join("");
  host.onclick = function (ev) {
    var b = ev.target.closest("[data-v]");
    if (!b) return;
    onPick(b.dataset.v);
  };
}

/* ---------- 出力（ソース一覧の除外設定を使い、全件を集計） ---------- */

function aggregateFull(dropExcluded) {
  var vendors = new Map();
  var extTotals = new Map();
  var totalFiles = 0, grand = 0, dropped = 0;

  state.rows.forEach(function (f) {
    var v = f.real;
    if (f.excluded) { dropped += v; if (dropExcluded) return; }
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
    vendors: list, files: totalFiles, grand: grand, dropped: dropped,
    exts: Array.from(extTotals.entries())
      .map(function (e) { return { name: e[0] || "(なし)", files: e[1].files, steps: e[1].steps }; })
      .sort(function (a, b) { return b.steps - a.steps; })
  };
}

async function exportSheets() {
  var dropExcluded = byId("drop-excluded").checked;
  var s = aggregateFull(dropExcluded);

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
