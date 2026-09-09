/* ============================================================
 * 財務ダッシュボード アドイン app.js
 * ------------------------------------------------------------
 * 役割: Excelからの読み取り・変更監視・画面制御
 *       計算と描画は dashboard.js
 *
 * 読み取り対象シート（名前は前方一致で探すので多少の表記ゆれは吸収）
 *   月次PL(実績入力)   … 科目 × 12ヶ月の発生額、年間計画、決算賞与ブロック
 *   月次BS(実績入力)   … 科目 × 12ヶ月の月末残高
 *   36期実績(決算報告書) … 前期の年間実績
 *   前提条件           … 年間計画の前提値
 *
 * 月の列は見出し行の「2025/10」形式のセルから自動判定するため、
 * 列を挿入しても・期が変わっても読み替え不要。
 * ============================================================ */

const APP_VERSION = "rev_20260909_a1";
const SHEET_PL = "月次PL";
const SHEET_BS = "月次BS";
const SHEET_36 = "期実績";
const SHEET_ASSUM = "前提条件";
const STORE_KEY = "zaimu_settings";

let demoMode = false;
let liveOn = true;
let reloadTimer = null;
let eventsBound = false;
let sheetNames = {};

/* ============================================================
   起動
   ============================================================ */
if (window.Office && window.Office.onReady) {
  Office.onReady(() => whenDomReady(init));
} else {
  window.addEventListener("DOMContentLoaded", init);
}
function whenDomReady(fn) {
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", fn, { once: true });
  else fn();
}

async function init() {
  const vl = document.getElementById("version-label");
  if (vl) vl.textContent = APP_VERSION;
  const stored = restoreSettings();
  bindControls();
  await loadData(!stored);
  if (!demoMode && liveOn) bindExcelEvents();
  bindResize();
}

/* ============================================================
   データ読み込み
   ============================================================ */
async function loadData(applyDefaults) {
  let data = null, err = "";
  if (window.Office && window.Excel && Office.context && Office.context.host) {
    try {
      data = await readWorkbook();
      demoMode = false;
    } catch (e) {
      err = (e && e.message) ? e.message : String(e);
      console.warn("Excel読込に失敗しました。デモデータで表示します。", e);
    }
  }
  if (!data) { data = JSON.parse(JSON.stringify(DEMO_DATA)); demoMode = true; }

  const keepPeople = BN.peopleDirty ? BN.people.slice() : null;
  setData(data);
  if (applyDefaults) applyDataDefaults();
  BN.people = keepPeople || (data.bonusPeople && data.bonusPeople.length
    ? data.bonusPeople.map(p => ({ name: p.name, amt: p.amt }))
    : [{ name: "支給対象1", amt: 0 }]);

  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
  document.getElementById("loadError").innerHTML = err
    ? '<div class="warn">Excelから読み取れませんでした（' + esc(err) + '）。デモデータを表示しています。対象ブックを開いた状態で再読み込みしてください。</div>' : "";
  buildMonthSelect();
  buildTuners();
  buildPeople();
  syncControls();
  render();
  stampFetched();
}

function esc(s) { return String(s).replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c])); }

async function readWorkbook() {
  return await Excel.run(async ctx => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    const names = wss.items.map(w => w.name);
    const pick = (frag) => names.find(n => n.indexOf(frag) === 0) || names.find(n => n.indexOf(frag) >= 0);
    sheetNames = {
      pl: pick(SHEET_PL), bs: pick(SHEET_BS),
      f36: names.find(n => /\d+期実績/.test(n)) || pick(SHEET_36),
      assum: pick(SHEET_ASSUM)
    };
    if (!sheetNames.pl || !sheetNames.bs) {
      throw new Error("「月次PL」「月次BS」で始まるシートが見つかりません");
    }
    const grab = nm => {
      if (!nm) return null;
      const ws = ctx.workbook.worksheets.getItem(nm);
      const used = ws.getUsedRangeOrNullObject(true);
      const rect = used.getBoundingRect(ws.getRange("A1"));
      rect.load("values");
      return rect;
    };
    const rPL = grab(sheetNames.pl), rBS = grab(sheetNames.bs),
      r36 = grab(sheetNames.f36), rAS = grab(sheetNames.assum);
    await ctx.sync();
    return buildData(
      rPL.values, rBS.values,
      r36 ? r36.values : [], rAS ? rAS.values : []
    );
  });
}

/* ---------- パース ---------- */
const isMonthCell = s => /^\d{4}\/\d{1,2}$/.test(String(s == null ? "" : s).trim());
function num(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = String(v).replace(/[,\s円]/g, "");
  if (s === "" || isNaN(Number(s))) return null;
  return Number(s);
}
function findMonthHeader(vals) {
  for (let r = 0; r < vals.length; r++) {
    const cols = [];
    (vals[r] || []).forEach((c, i) => { if (isMonthCell(c)) cols.push(i); });
    if (cols.length >= 6) return { row: r, cols: cols };
  }
  return null;
}

function buildData(plVals, bsVals, f36Vals, asVals) {
  const plHdr = findMonthHeader(plVals);
  if (!plHdr) throw new Error("月次PLの見出し行（2025/10 形式）が見つかりません");
  const bsHdr = findMonthHeader(bsVals);
  if (!bsHdr) throw new Error("月次BSの見出し行（2025/10 形式）が見つかりません");

  const months = plHdr.cols.map(c => String(plVals[plHdr.row][c]).trim());
  const hdrRow = plVals[plHdr.row] || [];
  let annualCol = 1;
  hdrRow.forEach((c, i) => { if (/年間計画/.test(String(c || ""))) annualCol = i; });

  /* --- PL --- */
  const pl = [];
  let sgaStart = -1, sgaEnd = -1;
  for (let r = plHdr.row + 1; r < plVals.length; r++) {
    const nm = String((plVals[r] || [])[0] || "").trim();
    if (!nm) continue;
    if (/^【販売費/.test(nm)) { sgaStart = pl.length; continue; }
    if (/^【/.test(nm) || /^※/.test(nm)) continue;
    pl.push({
      name: nm,
      annualPlan: num(plVals[r][annualCol]),
      v: plHdr.cols.map(c => num(plVals[r][c]))
    });
    if (/販管費/.test(nm) && /合計/.test(nm)) sgaEnd = pl.length - 1;
    if (normName(nm) === "経常利益") break;
  }
  if (sgaStart >= 0 && sgaEnd > sgaStart) for (let i = sgaStart; i < sgaEnd; i++) pl[i].isSga = true;

  /* --- PL内の決算賞与ブロック・みなし仕入率 --- */
  const bonusPeople = [];
  let mnashi = null;
  plVals.forEach(row => {
    const nm = String((row || [])[0] || "").trim();
    const mm = nm.match(/^決算賞与[\s　]*(.+)$/);
    if (mm && mm[1].trim() !== "計") {
      const amt = num(row[1]);
      if (amt !== null) bonusPeople.push({ name: mm[1].trim(), amt: Math.round(amt) });
    }
    if (/みなし仕入率/.test(nm)) { const v = num(row[1]); if (v !== null) mnashi = v > 1 ? v / 100 : v; }
  });

  /* --- BS --- */
  const bs = [];
  for (let r = bsHdr.row + 1; r < bsVals.length; r++) {
    const nm = String((bsVals[r] || [])[0] || "").trim();
    if (!nm) continue;
    if (/^■/.test(nm)) break;               // 検算ブロック以降は読まない（同名行の上書きを防ぐ）
    bs.push({ name: nm, v: bsHdr.cols.map(c => num(bsVals[r][c])) });
  }

  /* --- 前期実績 --- */
  const fy36 = [];
  (f36Vals || []).forEach(row => {
    const nm = String((row || [])[0] || "").trim();
    const a = num((row || [])[1]);
    if (nm && a !== null && !/^【/.test(nm)) fy36.push({ name: nm, amount: a });
  });

  /* --- 前提条件 --- */
  const assum = { salesPlan: null, sgaPlan: null, cogsRate: null };
  let termLabel = "";
  (asVals || []).forEach(row => {
    const nm = String((row || [])[0] || "").trim();
    const v = num((row || [])[1]);
    if (/想定年間売上高/.test(nm)) assum.salesPlan = v;
    else if (/想定年間販管費/.test(nm)) assum.sgaPlan = v;
    else if (/売上原価率/.test(nm)) assum.cogsRate = v;
    const t = nm.match(/第(\d+)期/);
    if (t && !termLabel) termLabel = "第" + t[1] + "期";
  });

  return { months, pl, bs, fy36, assum, bonusPeople, mnashi, termLabel };
}

/* ============================================================
   Excelの変更監視（リアルタイム反映）
   ============================================================ */
async function bindExcelEvents() {
  if (eventsBound || demoMode) return;
  try {
    await Excel.run(async ctx => {
      [sheetNames.pl, sheetNames.bs, sheetNames.assum].forEach(nm => {
        if (!nm) return;
        const ws = ctx.workbook.worksheets.getItem(nm);
        ws.onChanged.add(onExcelChanged);
      });
      await ctx.sync();
    });
    const has18 = Office.context.requirements && Office.context.requirements.isSetSupported("ExcelApi", "1.8");
    if (has18) {
      await Excel.run(async ctx => {
        [sheetNames.pl, sheetNames.bs].forEach(nm => {
          if (!nm) return;
          ctx.workbook.worksheets.getItem(nm).onCalculated.add(onExcelChanged);
        });
        await ctx.sync();
      });
    }
    eventsBound = true;
  } catch (e) {
    console.warn("変更監視の登録に失敗しました。手動の再読み込みをご利用ください。", e);
    setLive(false, true);
  }
}
async function onExcelChanged() {
  if (!liveOn) return;
  scheduleReload();
}
function scheduleReload() {
  setBadge("busy", "更新中");
  clearTimeout(reloadTimer);
  reloadTimer = setTimeout(() => reload(false), 600);
}
async function reload(manual) {
  if (manual) setBadge("busy", "更新中");
  const pane = document.querySelector(".pane:not([style*='display: none'])");
  const top = pane ? pane.scrollTop : 0;
  await loadData(false);
  if (pane) pane.scrollTop = top;
  setBadge(liveOn ? "on" : "off", liveOn ? "自動更新" : "手動");
}
function toggleLive() {
  setLive(!liveOn, false);
  if (liveOn) { bindExcelEvents(); reload(true); }
}
function setLive(on, silent) {
  liveOn = on;
  saveSettings();
  setBadge(on ? "on" : "off", on ? "自動更新" : "手動");
}
function setBadge(state, text) {
  const b = document.getElementById("live-badge");
  b.className = "live-badge" + (state === "on" ? "" : " " + state);
  document.getElementById("live-text").textContent = text;
}
function stampFetched() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  document.getElementById("fetched").textContent =
    (demoMode ? "デモデータ" : "Excelから取得") + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
}

/* ============================================================
   画面制御
   ============================================================ */
function switchTab(id) {
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.p === id));
  document.querySelectorAll(".pane").forEach(p => p.style.display = (p.id === id ? "" : "none"));
  requestAnimationFrame(() => render());
}
let resizeTimer = null;
function bindResize() {
  const on = () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => render(), 220); };
  window.addEventListener("resize", on);
  if (window.ResizeObserver) { try { new ResizeObserver(on).observe(document.body); } catch (e) { } }
}

function buildMonthSelect() {
  const sel = document.getElementById("lastActual");
  const sig = M.join("|");
  if (sel.dataset.sig !== sig) {
    sel.innerHTML = M.map((m, i) => '<option value="' + i + '">' + m + '（' + (i + 1) + 'ヶ月目）</option>').join("");
    sel.dataset.sig = sig;
  }
  sel.value = S.last;
}

function applyDataDefaults() {
  S.cogsRate = Math.round(actualCogsRate(S.last) * 1000) / 10;
  const c = compute();
  S.nyCogs = c.tot.sales ? Math.round(c.tot.cogs / c.tot.sales * 1000) / 10 : S.cogsRate;
  if (DATA.mnashi != null) BN.mnashi = Math.round(DATA.mnashi * 100);
}

function buildTuners() {
  [["sgaTuner", S.sgaAdj], ["nySgaTuner", S.nySgaAdj]].forEach(([hostId, store]) => {
    const host = document.getElementById(hostId);
    if (host.dataset.built) { syncCatLabels(hostId, store); return; }
    host.innerHTML = CATS.map(c =>
      '<div class="ctl"><div class="cl"><span>' + c + '</span><span class="cv" id="v_' + hostId + '_' + c + '">0%</span></div>'
      + '<input type="range" data-cat="' + c + '" min="-30" max="30" step="1" value="0"></div>').join("");
    host.addEventListener("input", e => {
      const c = e.target.dataset.cat; if (!c) return;
      store[c] = +e.target.value;
      document.getElementById("v_" + hostId + "_" + c).textContent = (store[c] > 0 ? "+" : "") + store[c] + "%";
      saveSettings(); render();
    });
    host.dataset.built = "1";
    syncCatLabels(hostId, store);
  });
}
function syncCatLabels(hostId, store) {
  CATS.forEach(c => {
    const lab = document.getElementById("v_" + hostId + "_" + c);
    const inp = document.querySelector("#" + hostId + " input[data-cat='" + c + "']");
    if (lab) lab.textContent = (store[c] > 0 ? "+" : "") + (store[c] || 0) + "%";
    if (inp) inp.value = store[c] || 0;
  });
}

function bindControls() {
  document.getElementById("lastActual").addEventListener("change", e => {
    S.last = +e.target.value; saveSettings(); render();
  });
  seg("segMethod", v => S.method = v);
  seg("segCogs", v => S.cogsMode = v);
  seg("segBase", v => S.base = v);
  seg("segSoc", v => BN.socWhen = v);
  rng("salesAdj", v => S.salesAdj = v);
  rng("cogsRate", v => S.cogsRate = v);
  rng("growth", v => S.growth = v);
  rng("season", v => S.season = v);
  rng("nyCogsRate", v => S.nyCogs = v);
  ["socRate", "mnashi", "kintou", "empSoc", "empTax", "tgtRate", "resMonths"].forEach(id =>
    rng(id, v => BN[id === "resMonths" ? "resMonths" : id] = v));

  document.getElementById("btRange").addEventListener("input", e => {
    const t = +e.target.value, cur = btTotal();
    if (cur > 0) BN.people.forEach(p => p.amt = Math.round(p.amt * t / cur / 1000) * 1000);
    else BN.people.forEach(p => p.amt = Math.round(t / BN.people.length / 1000) * 1000);
    BN.peopleDirty = true;
    document.querySelectorAll("#peopleBody .pa").forEach((el, i) => el.value = BN.people[i].amt);
    saveSettings(); render();
  });
  document.getElementById("btnAddPerson").addEventListener("click", () => {
    BN.people.push({ name: "新しい支給対象", amt: 0 });
    BN.peopleDirty = true; buildPeople(); saveSettings(); render();
  });
  document.getElementById("btnSyncPeople").addEventListener("click", () => {
    BN.peopleDirty = false;
    BN.people = (DATA.bonusPeople || []).map(p => ({ name: p.name, amt: p.amt }));
    if (!BN.people.length) BN.people = [{ name: "支給対象1", amt: 0 }];
    buildPeople(); document.getElementById("btRange").value = Math.min(4000000, btTotal());
    saveSettings(); render();
  });
  document.getElementById("btnReset").addEventListener("click", () => {
    S.method = "sheet"; S.cogsMode = "method"; S.salesAdj = 0; S.base = "fc37";
    S.growth = 0; S.season = 100; S.last = defaultLastActual();
    CATS.forEach(c => { S.sgaAdj[c] = 0; S.nySgaAdj[c] = 0; });
    Object.assign(BN, { socRate: 15, socWhen: "this", kintou: 180000, empSoc: 15, empTax: 10.21, tgtRate: 3, resMonths: 3 });
    BN.peopleDirty = false;
    BN.people = (DATA.bonusPeople || []).map(p => ({ name: p.name, amt: p.amt }));
    if (!BN.people.length) BN.people = [{ name: "支給対象1", amt: 0 }];
    applyDataDefaults();
    buildMonthSelect(); buildTuners(); buildPeople(); syncControls(); saveSettings(); render();
  });
  document.getElementById("btnCopy").addEventListener("click", copyResult);
}
function seg(id, set) {
  const el = document.getElementById(id);
  el.addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    Array.prototype.forEach.call(el.children, x => x.setAttribute("aria-pressed", x === b));
    set(b.dataset.v); saveSettings(); render();
  });
}
function rng(id, set) {
  document.getElementById(id).addEventListener("input", e => { set(+e.target.value); saveSettings(); render(); });
}
function syncControls() {
  const put = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
  put("salesAdj", S.salesAdj); put("cogsRate", S.cogsRate); put("growth", S.growth);
  put("season", S.season); put("nyCogsRate", S.nyCogs);
  put("socRate", BN.socRate); put("mnashi", BN.mnashi); put("kintou", BN.kintou);
  put("empSoc", BN.empSoc); put("empTax", BN.empTax); put("tgtRate", BN.tgtRate); put("resMonths", BN.resMonths);
  put("btRange", Math.min(4000000, btTotal()));
  const press = (id, v) => Array.prototype.forEach.call(document.getElementById(id).children,
    x => x.setAttribute("aria-pressed", x.dataset.v === v));
  press("segMethod", S.method); press("segCogs", S.cogsMode); press("segBase", S.base); press("segSoc", BN.socWhen);
  setBadge(liveOn ? "on" : "off", liveOn ? "自動更新" : "手動");
}

function copyResult() {
  const cur = compute(), nx = computeNext(cur), m = bonusModel(cur), bt = btTotal();
  const t = [
    (DATA.termLabel || "今期") + " 財務ダッシュボード",
    "", "【着地見込み】",
    "売上高 " + yen(cur.tot.sales), "売上原価 " + yen(cur.tot.cogs),
    "販管費 " + yen(cur.tot.sga), "営業利益 " + syen(cur.tot.op),
    "", "【次期 計画】",
    "売上高 " + yen(nx.tot.sales), "販管費 " + yen(nx.tot.sga),
    "営業利益 " + syen(nx.tot.op), "損益分岐点売上高 " + yen(nx.bep),
    "", "【決算賞与 試算】",
    "賞与前 営業利益率 " + pct(m.base.rate),
    "賞与総額 " + yen(bt) + "（" + BN.people.map(p => p.name + " " + yen(p.amt)).join(" / ") + "）",
    "会社総コスト " + yen(m.now.cashout) + "　節税 " + yen(m.base.T.total - m.now.T.total)
    + "　正味 " + yen(m.now.cashout - (m.base.T.total - m.now.T.total)),
    "賞与後 営業利益率 " + pct(m.now.rate),
    "税引後利益 " + syen(m.now.NI) + "　期末後現預金 " + yen(m.now.endCash),
    "上限：利益ゼロ " + yen(m.limZero) + " / 目標率" + BN.tgtRate + "% " + yen(m.limTgt)
    + " / 運転資金" + BN.resMonths + "ヶ月 " + yen(m.limCash) + " → 採用 " + yen(m.rec),
    "", "【設定】",
    "実績確定月 " + M[S.last] + "／残り月の推計 " + S.method + "／売上調整 " + S.salesAdj + "%",
    "販管費調整 " + CATS.map(c => c + " " + S.sgaAdj[c] + "%").join(" / "),
    "次期基準 " + S.base + "／成長率 " + S.growth + "%／季節性 " + S.season + "%／原価率 " + S.nyCogs + "%",
    "次期販管費 " + CATS.map(c => c + " " + S.nySgaAdj[c] + "%").join(" / ")
  ].join("\n");
  const done = ok => {
    const el = document.getElementById("copyMsg");
    el.textContent = ok ? "コピーしました" : "コピーできませんでした";
    setTimeout(() => el.textContent = "", 2200);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(t).then(() => done(true)).catch(() => fallbackCopy(t, done));
  } else fallbackCopy(t, done);
}
function fallbackCopy(text, done) {
  try {
    const ta = document.createElement("textarea");
    ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta); ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta); done(ok);
  } catch (e) { done(false); }
}

/* ============================================================
   設定の保存
   ============================================================ */
function saveSettings() {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify({
      live: liveOn,
      s: { last: S.last, method: S.method, cogsMode: S.cogsMode, salesAdj: S.salesAdj, cogsRate: S.cogsRate, sgaAdj: S.sgaAdj, base: S.base, growth: S.growth, season: S.season, nyCogs: S.nyCogs, nySgaAdj: S.nySgaAdj },
      b: { people: BN.people, peopleDirty: BN.peopleDirty, socRate: BN.socRate, socWhen: BN.socWhen, mnashi: BN.mnashi, kintou: BN.kintou, empSoc: BN.empSoc, empTax: BN.empTax, tgtRate: BN.tgtRate, resMonths: BN.resMonths }
    }));
  } catch (e) { }
}
function restoreSettings() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return false;
    const o = JSON.parse(raw);
    if (typeof o.live === "boolean") liveOn = o.live;
    if (o.s) Object.keys(o.s).forEach(k => { if (o.s[k] !== undefined && o.s[k] !== null) S[k] = o.s[k]; });
    if (o.b) Object.keys(o.b).forEach(k => { if (o.b[k] !== undefined && o.b[k] !== null) BN[k] = o.b[k]; });
    CATS.forEach(c => { if (S.sgaAdj[c] == null) S.sgaAdj[c] = 0; if (S.nySgaAdj[c] == null) S.nySgaAdj[c] = 0; });
    return true;
  } catch (e) { return false; }
}

/* ============================================================
   共通スライドメニュー
   ============================================================ */
const COMMON_BASE = "https://ymatsuda-cmyk.github.io/tools/common";
let menuReady = null;
function openMenu() {
  if (!menuReady) {
    menuReady = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = COMMON_BASE + "/slide-menu.js";
      s.onload = () => {
        SlideMenu.init({
          appName: "財務ダッシュボード", version: APP_VERSION, position: "left",
          currentId: "zaimu", menuUrl: COMMON_BASE + "/menu.json",
          localItems: [{ label: "今すぐ再読み込み", icon: "", onClick: () => reload(true) }]
        });
        resolve();
      };
      s.onerror = e => { menuReady = null; reject(e); };
      document.head.appendChild(s);
    });
  }
  menuReady.then(() => SlideMenu.open()).catch(() => { menuReady = null; });
}
