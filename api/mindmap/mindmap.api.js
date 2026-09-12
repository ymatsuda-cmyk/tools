/*!
 * mindmap.api.js — 原文からマインドマップを作り、描画・編集するライブラリ
 *
 *   MindMap.mount(el, options)                 描画先を指定して初期化
 *   MindMap.fromText(原文, { useAI: true })     原文をツリーに構造化（Promise）
 *   MindMap.render(tree)                       ツリーを描画して編集可能にする
 *   MindMap.getData() / MindMap.toJSON()       現在の編集結果
 *   MindMap.on("change"|"zoom", fn)            変更通知
 *   MindMap.zoomIn() / zoomOut() / fit()       表示操作
 *   MindMap.palette                            マーカー色の定義
 *
 * ツリー形式:
 *   { "text": "10文字以内", "color": "green", "collapsed": false, "children": [ ... ] }
 *   color を省いたノードは白。
 */
(function (global) {
  "use strict";

  var PALETTE = {
    purple: { bg: "#cfcbf2", fg: "#2e2760", edge: "#a9a2e4" },
    green:  { bg: "#a6e0cb", fg: "#0b4a3a", edge: "#6fc9ab" },
    orange: { bg: "#f6c7a8", fg: "#5a2a10", edge: "#e5a377" },
    yellow: { bg: "#fadc8f", fg: "#4a3105", edge: "#edc45c" },
    pink:   { bg: "#f4c2d4", fg: "#4b182c", edge: "#e293b3" }
  };

  var NODE_H = 34, ROOT_H = 44, SLOT = 46,
      ROOT_W = 160, NODE_W = 136, HGAP = 64, PAD = 48,
      MAXLEN = 10;

  var SVG_NS = "http://www.w3.org/2000/svg";
  var ENDPOINT = "https://api.anthropic.com/v1/messages";
  var MODEL = "claude-sonnet-4-6";

  var viewport = null, stage = null, wires = null;
  var root = null, sel = null, editing = null, pop = null;
  var zoom = 1, panX = 0, panY = 0;
  var history = [], listeners = { change: [], zoom: [] };
  var uid = 0;

  function nid() { uid++; return "n" + uid; }
  function on(ev, fn) { if (listeners[ev]) listeners[ev].push(fn); }
  function fire(ev, arg) {
    var l = listeners[ev] || [];
    for (var i = 0; i < l.length; i++) l[i](arg);
  }

  /* ---------- モデル ---------- */

  function clean(t) {
    return (t || "").replace(/\s+/g, " ").trim().slice(0, MAXLEN);
  }
  function normalize(raw) {
    var n = {
      id: nid(),
      text: clean(raw.text) || "無題",
      color: (raw.color && PALETTE[raw.color]) ? raw.color : null,
      collapsed: !!raw.collapsed,
      children: []
    };
    var kids = raw.children || [];
    for (var i = 0; i < kids.length; i++) n.children.push(normalize(kids[i]));
    return n;
  }
  function serialize(n) {
    var o = { id: n.id, text: n.text };
    if (n.color) o.color = n.color;
    if (n.collapsed) o.collapsed = true;
    if (n.children.length) o.children = n.children.map(serialize);
    return o;
  }
  function findParent(n, target) {
    for (var i = 0; i < n.children.length; i++) {
      if (n.children[i] === target) return n;
      var r = findParent(n.children[i], target);
      if (r) return r;
    }
    return null;
  }
  function walk(n, fn) {
    fn(n);
    var cs = n.collapsed ? [] : n.children;
    for (var i = 0; i < cs.length; i++) walk(cs[i], fn);
  }
  function countAll(n) {
    var c = 0;
    (function rec(x) { c++; for (var i = 0; i < x.children.length; i++) rec(x.children[i]); })(n);
    return c;
  }
  function snapshot() {
    if (!root) return;
    history.push(JSON.stringify(serialize(root)));
    if (history.length > 60) history.shift();
  }
  function undo() {
    if (!history.length) return;
    var prev = history.pop();
    uid = 0;
    root = normalize(JSON.parse(prev));
    sel = null;
    render(); emit();
  }
  function emit() { fire("change", root ? serialize(root) : null); }

  /* ---------- レイアウト ---------- */

  function leaves(n) {
    if (n.collapsed || !n.children.length) return 1;
    var s = 0;
    for (var i = 0; i < n.children.length; i++) s += leaves(n.children[i]);
    return s;
  }
  function shift(n, dy) {
    n.y += dy;
    for (var i = 0; i < n.children.length; i++) shift(n.children[i], dy);
  }
  function layout() {
    var kids = root.collapsed ? [] : root.children;
    var right = [], left = [], rw = 0, lw = 0;
    for (var i = 0; i < kids.length; i++) {
      var w = leaves(kids[i]);
      if (rw <= lw) { right.push(kids[i]); rw += w; }
      else { left.push(kids[i]); lw += w; }
    }
    function side(list, dir) {
      var cursor = 0;
      function place(n, depth) {
        n.depth = depth; n.dir = dir;
        var cs = n.collapsed ? [] : n.children;
        if (!cs.length) { n.y = cursor + NODE_H / 2; cursor += SLOT; }
        else {
          for (var j = 0; j < cs.length; j++) place(cs[j], depth + 1);
          n.y = (cs[0].y + cs[cs.length - 1].y) / 2;
        }
        n.x = dir > 0
          ? ROOT_W / 2 + HGAP + (depth - 1) * (NODE_W + HGAP)
          : -(ROOT_W / 2 + HGAP + (depth - 1) * (NODE_W + HGAP)) - NODE_W;
      }
      for (var k = 0; k < list.length; k++) place(list[k], 1);
      if (list.length) {
        var anchor = (list[0].y + list[list.length - 1].y) / 2;
        for (var m = 0; m < list.length; m++) shift(list[m], -anchor);
      }
    }
    side(right, 1); side(left, -1);

    root.x = -ROOT_W / 2; root.y = 0; root.depth = 0; root.dir = 0;

    var minX = -ROOT_W / 2, maxX = ROOT_W / 2, minY = -ROOT_H / 2, maxY = ROOT_H / 2;
    walk(root, function (n) {
      if (n === root) return;
      minX = Math.min(minX, n.x); maxX = Math.max(maxX, n.x + NODE_W);
      minY = Math.min(minY, n.y - NODE_H / 2); maxY = Math.max(maxY, n.y + NODE_H / 2);
    });
    var ox = PAD - minX, oy = PAD - minY;
    walk(root, function (n) { n.px = n.x + ox; n.py = n.y + oy; });
    return { w: maxX - minX + PAD * 2, h: maxY - minY + PAD * 2 };
  }

  /* ---------- 描画 ---------- */

  function render() {
    if (!stage) return;
    closePop();
    var old = stage.querySelectorAll(".node,.count,.toggle");
    for (var i = 0; i < old.length; i++) old[i].remove();
    wires.innerHTML = "";
    if (!root) return;

    var size = layout();
    stage.style.width = size.w + "px";
    stage.style.height = size.h + "px";
    wires.setAttribute("viewBox", "0 0 " + size.w + " " + size.h);
    wires.setAttribute("width", size.w);
    wires.setAttribute("height", size.h);

    drawNode(root);
    walk(root, function (n) {
      var cs = n.collapsed ? [] : n.children;
      for (var j = 0; j < cs.length; j++) { drawWire(n, cs[j]); drawNode(cs[j]); }
    });
    applyView();
  }
  function drawWire(p, c) {
    var pw = (p === root ? ROOT_W : NODE_W);
    var dir = c.dir;
    var x1 = dir > 0 ? p.px + pw : p.px;
    var y1 = p.py;
    var x2 = dir > 0 ? c.px : c.px + NODE_W;
    var y2 = c.py;
    var mx = (x1 + x2) / 2;
    var path = document.createElementNS(SVG_NS, "path");
    path.setAttribute("d", "M" + x1 + " " + y1 + " C" + mx + " " + y1 + " " + mx + " " + y2 + " " + x2 + " " + y2);
    if (c.depth > 1) path.setAttribute("class", "deep");
    wires.appendChild(path);
  }
  function drawNode(n) {
    var isRoot = (n === root);
    var el = document.createElement("div");
    el.className = "node" + (isRoot ? " root" : "") + (n === sel ? " sel" : "");
    el.dataset.id = n.id;
    var w = isRoot ? ROOT_W : NODE_W, h = isRoot ? ROOT_H : NODE_H;
    el.style.left = n.px + "px";
    el.style.top = (n.py - h / 2) + "px";
    el.style.width = w + "px";
    if (isRoot) el.style.height = ROOT_H + "px";
    var p = n.color ? PALETTE[n.color] : null;
    if (p) { el.style.background = p.bg; el.style.color = p.fg; el.style.borderColor = p.edge; }
    el.textContent = n.text;
    el._node = n;
    stage.appendChild(el);

    if (n.children.length) {
      var t = document.createElement("button");
      t.className = "toggle" + (n.collapsed ? " on" : "");
      t.textContent = n.collapsed ? leaves(n) : "−";
      t.title = n.collapsed ? "開く" : "閉じる";
      var dir = isRoot ? 1 : n.dir;
      t.style.left = (dir > 0 ? n.px + w + 6 : n.px - 28) + "px";
      t.style.top = n.py + "px";
      t.onclick = function (e) {
        e.stopPropagation();
        snapshot(); n.collapsed = !n.collapsed; render(); emit();
      };
      stage.appendChild(t);
    }
  }
  function nodeEl(n) { return stage.querySelector('.node[data-id="' + n.id + '"]'); }

  /* ---------- 選択・編集 ---------- */

  function select(n) {
    if (editing) commitEdit();
    sel = n;
    var all = stage.querySelectorAll(".node");
    for (var i = 0; i < all.length; i++) all[i].classList.toggle("sel", all[i]._node === n);
    closePop();
    if (n) openPop(n);
  }
  function openPop(n) {
    var el = nodeEl(n); if (!el) return;
    pop = document.createElement("div");
    pop.className = "pop";
    var keys = Object.keys(PALETTE);
    for (var i = 0; i < keys.length; i++) pop.appendChild(swatch(n, keys[i]));
    pop.appendChild(swatch(n, null));
    var sep = document.createElement("div"); sep.className = "sep"; pop.appendChild(sep);
    pop.appendChild(actBtn("編集", function () { startEdit(n); }));
    pop.appendChild(actBtn("子を追加", function () { addChild(n); }));
    if (n !== root) {
      var d = actBtn("削除", function () { removeNode(n); });
      d.classList.add("del"); pop.appendChild(d);
    }
    stage.appendChild(pop);
    var h = (n === root ? ROOT_H : NODE_H);
    pop.style.left = n.px + "px";
    pop.style.top = (n.py - h / 2 - pop.offsetHeight - 10) + "px";
    if (parseFloat(pop.style.top) < 4) pop.style.top = (n.py + h / 2 + 10) + "px";
  }
  function swatch(n, key) {
    var b = document.createElement("button");
    b.className = "sw" + (key ? "" : " none") + (n.color === key ? " act" : "");
    b.title = key ? "色を変える" : "色なし";
    if (key) b.style.background = PALETTE[key].bg;
    b.onclick = function (e) {
      e.stopPropagation();
      snapshot(); n.color = key; render(); select(n); emit();
    };
    return b;
  }
  function actBtn(label, fn) {
    var b = document.createElement("button");
    b.className = "act-btn"; b.textContent = label;
    b.onclick = function (e) { e.stopPropagation(); fn(); };
    return b;
  }
  function closePop() { if (pop) { pop.remove(); pop = null; } }

  function startEdit(n) {
    if (editing) commitEdit();
    closePop();
    var el = nodeEl(n); if (!el) return;
    editing = n;
    el.textContent = "";
    var input = document.createElement("input");
    input.value = n.text; input.maxLength = MAXLEN;
    el.appendChild(input);

    var count = document.createElement("div");
    count.className = "count";
    count.style.left = (n.px + NODE_W - 42) + "px";
    count.style.top = (n.py + NODE_H / 2 + 4) + "px";
    stage.appendChild(count);
    function upd() {
      count.textContent = input.value.length + "/" + MAXLEN;
      count.classList.toggle("warn", input.value.length >= MAXLEN);
    }
    upd();
    input.addEventListener("input", upd);
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); commitEdit(); }
      else if (e.key === "Escape") { e.preventDefault(); cancelEdit(); }
      else if (e.key === "Tab") { e.preventDefault(); var t = editing; commitEdit(); addChild(t); }
    });
    input.addEventListener("blur", function () { if (editing === n) commitEdit(); });
    input.focus(); input.select();
  }
  function commitEdit() {
    if (!editing) return;
    var n = editing, el = nodeEl(n), input = el && el.querySelector("input");
    var v = input ? clean(input.value) : n.text;
    editing = null;
    if (v && v !== n.text) { snapshot(); n.text = v; emit(); }
    render(); if (sel) select(sel);
  }
  function cancelEdit() { var n = editing; editing = null; render(); if (n) select(n); }

  function addChild(n) {
    snapshot();
    if (n.collapsed) n.collapsed = false;
    var c = normalize({ text: "新しい要素" });
    n.children.push(c);
    sel = c;
    render(); emit();
    startEdit(c);
  }
  function removeNode(n) {
    var p = findParent(root, n); if (!p) return;
    var c = countAll(n);
    if (c > 1 && !confirm("この要素と、ぶら下がる" + (c - 1) + "件を削除します。よろしいですか。")) return;
    snapshot();
    p.children.splice(p.children.indexOf(n), 1);
    sel = null; render(); emit();
  }

  /* ---------- 表示位置 ---------- */

  function applyView() {
    stage.style.transform = "translate(" + panX + "px," + panY + "px) scale(" + zoom + ")";
    fire("zoom", zoom);
  }
  function setZoom(z, cx, cy) {
    z = Math.min(2, Math.max(.2, z));
    if (cx === undefined) { cx = viewport.clientWidth / 2; cy = viewport.clientHeight / 2; }
    panX = cx - (cx - panX) * (z / zoom);
    panY = cy - (cy - panY) * (z / zoom);
    zoom = z; applyView();
  }
  function fit() {
    if (!root || !stage) return;
    var w = stage.offsetWidth, h = stage.offsetHeight;
    var z = Math.min((viewport.clientWidth - 32) / w, (viewport.clientHeight - 32) / h, 1);
    zoom = Math.max(.2, z);
    panX = (viewport.clientWidth - w * zoom) / 2;
    panY = (viewport.clientHeight - h * zoom) / 2;
    applyView();
  }

  /* ---------- 原文の構造化 ---------- */

  function trimPhrase(s) {
    s = s.replace(/^[\s、,]+/, "").replace(/[。\.]+$/, "");
    s = s.replace(/(します|しました|である|ます|です)$/, "");
    return clean(s);
  }
  function localOutline(src) {
    var lines = src.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
    if (!lines.length) return { text: "原文", children: [] };
    var tree = { text: clean(lines[0]), children: [] };
    var cur = null;
    var bullet = /^([・\-*•▪]|\d+[\.\)、])\s*/;
    for (var i = 1; i < lines.length; i++) {
      var L = lines[i];
      if (bullet.test(L)) {
        if (!cur) { cur = { text: "内容", children: [] }; tree.children.push(cur); }
        cur.children.push({ text: trimPhrase(L.replace(bullet, "")), children: [] });
      } else if (L.length <= 20 && !/[。\.]$/.test(L)) {
        cur = { text: clean(L), children: [] }; tree.children.push(cur);
      } else {
        var parts = L.split(/[。\.]/).filter(Boolean);
        for (var j = 0; j < parts.length; j++) {
          var t = trimPhrase(parts[j]); if (!t) continue;
          if (!cur) { cur = { text: "内容", children: [] }; tree.children.push(cur); }
          cur.children.push({ text: t, children: [] });
        }
      }
    }
    if (!tree.children.length) tree.children.push({ text: "内容", children: [] });
    return tree;
  }

  function callModel(src) {
    var instruction =
      "次の原文を日本語のマインドマップに構造化してください。\n" +
      "出力はJSONのみ。前置き・コードフェンス・説明は一切書かないこと。\n" +
      '形式: {"text":"中心テーマ","children":[{"text":"見出し","children":[{"text":"要点"}]}]}\n' +
      "制約:\n" +
      "- すべてのtextは日本語" + MAXLEN + "文字以内。超えそうなら削って体言止めにする\n" +
      "- 第1階層は3〜7個、第2階層は各0〜4個、3階層まで\n" +
      "- 原文にない情報を足さない\n\n原文:\n" + src;

    return fetch(ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 1000,
        messages: [{ role: "user", content: instruction }]
      })
    }).then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    }).then(function (data) {
      var text = (data.content || [])
        .filter(function (b) { return b.type === "text"; })
        .map(function (b) { return b.text; }).join("\n");
      text = text.replace(/```json/g, "").replace(/```/g, "").trim();
      var s = text.indexOf("{"), e = text.lastIndexOf("}");
      if (s < 0 || e < 0) throw new Error("JSONが取り出せませんでした");
      return JSON.parse(text.slice(s, e + 1));
    });
  }

  function fromText(src, opts) {
    opts = opts || {};
    if (opts.useAI === false) {
      API.lastSource = "local"; API.lastError = null;
      return Promise.resolve(localOutline(src));
    }
    return callModel(src).then(function (tree) {
      API.lastSource = "ai"; API.lastError = null;
      return tree;
    }, function (err) {
      API.lastSource = "local"; API.lastError = err;
      return localOutline(src);
    });
  }

  /* ---------- 初期化 ---------- */

  function setData(tree) {
    if (!stage) throw new Error("mount() を先に呼んでください");
    uid = 0; history = []; sel = null;
    root = normalize(tree);
    render(); fit(); emit();
    return API;
  }

  function mount(el) {
    viewport = (typeof el === "string") ? document.querySelector(el) : el;
    if (!viewport) throw new Error("描画先が見つかりません");

    stage = document.createElement("div");
    stage.className = "mm-stage";
    wires = document.createElementNS(SVG_NS, "svg");
    wires.setAttribute("class", "wires");
    stage.appendChild(wires);
    viewport.appendChild(stage);

    stage.addEventListener("click", function (e) {
      var t = e.target.closest(".node");
      if (!t || !t._node) return;
      e.stopPropagation(); select(t._node);
    });
    stage.addEventListener("dblclick", function (e) {
      var t = e.target.closest(".node");
      if (!t || !t._node) return;
      e.stopPropagation(); startEdit(t._node);
    });

    viewport.addEventListener("pointerdown", function (e) {
      if (e.target.closest(".node") || e.target.closest(".pop") || e.target.closest(".toggle")) return;
      select(null);
      var sx = e.clientX, sy = e.clientY, ox = panX, oy = panY;
      viewport.setPointerCapture(e.pointerId);
      viewport.classList.add("dragging");
      function mv(ev) { panX = ox + (ev.clientX - sx); panY = oy + (ev.clientY - sy); applyView(); }
      function up() {
        viewport.classList.remove("dragging");
        viewport.removeEventListener("pointermove", mv);
        viewport.removeEventListener("pointerup", up);
        viewport.removeEventListener("pointercancel", up);
      }
      viewport.addEventListener("pointermove", mv);
      viewport.addEventListener("pointerup", up);
      viewport.addEventListener("pointercancel", up);
    });

    viewport.addEventListener("wheel", function (e) {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) {
        var r = viewport.getBoundingClientRect();
        setZoom(zoom - e.deltaY * 0.0025, e.clientX - r.left, e.clientY - r.top);
      } else {
        panX -= e.deltaX; panY -= e.deltaY; applyView();
      }
    }, { passive: false });

    document.addEventListener("keydown", function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") { e.preventDefault(); undo(); return; }
      if (editing || !sel) return;
      if (e.target.tagName === "TEXTAREA" || e.target.tagName === "INPUT") return;
      if (e.key === "Enter") { e.preventDefault(); startEdit(sel); }
      else if (e.key === "Tab") { e.preventDefault(); addChild(sel); }
      else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); removeNode(sel); }
    });

    global.addEventListener("resize", function () { if (root) fit(); });
    return API;
  }

  var API = {
    mount: mount,
    render: setData,
    fromText: fromText,
    localOutline: localOutline,
    getData: function () { return root ? serialize(root) : null; },
    toJSON: function () { return root ? JSON.stringify(serialize(root), null, 2) : "null"; },
    count: function () { return root ? countAll(root) : 0; },
    on: on,
    undo: undo,
    fit: fit,
    setZoom: setZoom,
    zoomIn: function () { setZoom(zoom + .1); },
    zoomOut: function () { setZoom(zoom - .1); },
    palette: PALETTE,
    maxLength: MAXLEN,
    lastSource: null,
    lastError: null
  };

  global.MindMap = API;
})(window);
