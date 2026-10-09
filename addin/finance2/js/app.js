/* ============================================================
 * finance2 — 合計残高試算表の取込とBS/PLダッシュボード
 * タブ：取込 ／ 全体 ／ 月別 ／ 科目 ／ 予測 ／ 計画 ／ 予実
 * ============================================================ */
const APP_VERSION = "rev_20261009_yj02";
window.APP_VERSION = APP_VERSION;

(function () {
  "use strict";
  const $ = (s, el = document) => el.querySelector(s);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const norm = (s) => String(s || "").replace(/[\s\u3000]/g, "");
  const yen = (n) => (n == null ? "—" : (n < 0 ? "−" : "") + Math.abs(Math.round(n)).toLocaleString());
  const sgn = (n) => (n == null ? "—" : (n > 0 ? "+" : n < 0 ? "−" : "±") + Math.abs(Math.round(n)).toLocaleString());
  // 金額はすべて円のフル桁で表示する
  const k = yen, sk = sgn, mil = yen, smil = sgn;
  const pct = (a, b) => (a == null || b == null || b === 0 ? null : (a - b) / Math.abs(b) * 100);
  const spct = (v) => (v == null ? "—" : (v > 0 ? "+" : v < 0 ? "−" : "±") + Math.abs(v).toFixed(1) + "%");
  const mLabel = (p) => Number(p.slice(5)) + "月";
  const pLabel = (p) => p ? p.slice(0, 4) + "年" + Number(p.slice(5)) + "月" : "期間不明";
  const LS = "finance2-ui";

  const TABS = [["import", "取込"], ["dash", "全体"], ["monthly", "月別"], ["account", "科目"], ["forecast", "予測"], ["plan", "計画"], ["yojitsu", "予実"]];
  const state = {
    tab: "import", led: null, M: null, loadError: "",
    imports: [], seq: 0, diffId: null, diffAll: false, busy: "",
    fy: null, month: null, plMode: "month", showDetail: false,
    account: null, fcKey: null, fcMethod: "yoy", checks: {}, lastChecks: [],
    show: { prev: true, anom: true, fc: false, plan: true }, MT: null, plan: null, layout: {}, source: "tb", planTarget: null, planDetail: false, segAct: null, yj: null
  };
  try { Object.assign(state, JSON.parse(localStorage.getItem(LS) || "{}")); } catch (e) {}
  const persist = () => { try { localStorage.setItem(LS, JSON.stringify({ tab: state.tab, plMode: state.plMode, showDetail: state.showDetail, fcMethod: state.fcMethod, show: state.show, planDetail: state.planDetail })); } catch (e) {} };

  /* ---------- 読み込み ---------- */
  async function load() {
    try {
      state.led = await Ledger.readAll();
      state.loadError = "";
    } catch (e) {
      console.error(e);
      state.loadError = "ブックを読み込めませんでした：" + (e.message || e);
      if (!state.led) state.led = { exists: false, tb: [], history: [], master: {}, settings: { 期首月: 10 } };
    }
    // 取込の照合用（PDF科目ベース）と、表示用（月次PL・月次BSの行ベース）を分けて持つ
    state.MT = FinModel.build(state.led.tb, state.led.settings);
    state.M = state.MT; state.source = "tb"; state.layout = {};
    try {
      const sm = await MonthlySheets.readModel(state.led.tb, state.led.settings.期首月 || 10);
      state.segAct = null;
      if (sm.found && sm.rows.length) { state.M = FinModel.build(sm.rows, state.led.settings); state.source = "sheet"; state.layout = sm.layout; state.segAct = sm.seg || null; }
    } catch (e) { console.warn("readModel", e); }
    try { state.plan = await Plan.read(); } catch (e) { console.warn("plan", e); state.plan = null; }
    if (!state.yj) state.yj = Yojitsu.loadParams();
    const M = state.M;
    if (!state.fy || !M.fys.includes(state.fy)) state.fy = M.latest ? M.fyOf(M.latest) : null;
    if (!state.month || !M.has(state.month)) state.month = state.fy ? M.latestIn(state.fy) : null;
    if (!state.fcKey || !M.meta[state.fcKey]) state.fcKey = M.K.sales;
    if (!state.account || !M.meta[state.account]) state.account = M.K.sales;
    state.imports.forEach(evaluate);
    render();
    ensureAutoShow().catch(e => console.warn("autoShow", e));
  }


  /* ---------- ブックを開いたときに自動表示 ----------
   * manifest の TaskpaneId=Office.AutoShowTaskpaneWithDocument と、
   * ブック側の設定 Office.AutoShowTaskpaneWithDocument=true の両方で有効になる。
   * 台帳シート（TB_明細）があるブックだけ自動でONにする。OFFにしたブックは再度ONにしない。
   */
  const AUTO_KEY = "Office.AutoShowTaskpaneWithDocument";
  const AUTO_OFF_KEY = "finance2.autoShowOff";
  function docSettings() { try { return Office.context.document.settings; } catch (e) { return null; } }
  function autoShowState() { const s = docSettings(); return s ? !!s.get(AUTO_KEY) : null; }
  function setAutoShow(on, manual) {
    const s = docSettings(); if (!s) return Promise.resolve(false);
    s.set(AUTO_KEY, !!on);
    if (manual) s.set(AUTO_OFF_KEY, !on);
    return new Promise(res => s.saveAsync(r => res(r.status === "succeeded")));
  }
  async function ensureAutoShow() {
    const s = docSettings();
    if (!s || !state.led || !state.led.exists) return;
    if (s.get(AUTO_KEY) || s.get(AUTO_OFF_KEY)) return;
    await setAutoShow(true, false);
  }

  /* ---------- 取込：判定 ---------- */
  function ledgerRows(period) { return state.led.tb.filter(r => r.period === period); }

  function computeDiff(parsed) {
    const a = {}, b = {};
    ledgerRows(parsed.period).forEach(r => a[r.key] = r);
    parsed.rows.forEach(r => b[r.key] = r);
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])];
    const list = [];
    const counts = { chg: 0, pdf: 0, led: 0, same: 0 };
    keys.forEach(key => {
      const x = a[key], y = b[key];
      let type;
      if (x && y) type = (x.open === y.open && x.dr === y.dr && x.cr === y.cr && x.bal === y.bal) ? "same" : "chg";
      else type = y ? "pdf" : "led";
      counts[type]++;
      const ref = y || x;
      list.push({ key, type, sheet: ref.sheet, code: ref.code, name: ref.name, kind: ref.kind, order: ref.order + (ref.sheet === "PL" ? 1000 : 0), a: x, b: y, d: (y ? y.bal : 0) - (x ? x.bal : 0) });
    });
    list.sort((p, q) => p.order - q.order);
    const M = state.MT, imp = {};
    const pick = { 売上高: M.K.sales, 営業利益: M.K.op, 資産合計: M.K.assets, 負債合計: M.K.liab, 純資産: M.K.equity };
    Object.entries(pick).forEach(([label, key]) => {
      if (!key) return;
      const nm = M.meta[key] && norm(M.meta[key].name);
      const x = a[key] || Object.values(a).find(r => norm(r.name) === nm);
      const y = b[key] || parsed.rows.find(r => norm(r.name) === nm);
      if (x || y) imp[label] = (y ? y.bal : 0) - (x ? x.bal : 0);
    });
    return { list, counts, impact: imp, total: counts.chg + counts.pdf + counts.led };
  }

  function continuity(parsed) {
    const M = state.MT, warns = [];
    const prev = FinModel.addM(parsed.period, -1), next = FinModel.addM(parsed.period, 1);
    const fyStart = Number(parsed.period.slice(5)) === M.start;
    if (M.has(prev)) {
      const ng = parsed.rows.filter(r => {
        if (r.sheet === "PL" && fyStart) return false;
        const pb = M.bal(r.key, prev);
        return pb !== null && pb !== r.open;
      });
      if (ng.length) warns.push(`前月（${pLabel(prev)}）の残高と今月の繰越残高が ${ng.length}科目で一致しません（${ng.slice(0, 3).map(r => r.name).join("、")}など）`);
    }
    if (M.has(next)) {
      const nfy = Number(next.slice(5)) === M.start;
      const ng = parsed.rows.filter(r => {
        if (r.sheet === "PL" && nfy) return false;
        const no = M.open(r.key, next);
        return no !== null && no !== r.bal;
      });
      if (ng.length) warns.push(`翌月（${pLabel(next)}）の繰越残高と今月の残高が ${ng.length}科目で一致しません。翌月も取り込み直しが必要かもしれません`);
    }
    if (fyStart && parsed.rows.some(r => r.sheet === "PL" && r.open !== 0)) warns.push(`期首月（${M.start}月）なのに損益計算書の繰越残高が0ではありません。設定シートの期首月を確認してください`);
    return warns;
  }

  function evaluate(it) {
    if (["parsing", "done", "skip"].includes(it.status)) return;
    const p = it.parsed;
    it.warnings = p.warnings.slice();
    if (!p.rows.length || p.errors.some(e => !/期間/.test(e)) || (p.errors.length && p.period)) {
      it.status = "error"; it.msg = p.errors.join(" "); return;
    }
    if (!p.period) { it.status = "needPeriod"; it.msg = p.errors.join(" "); return; }
    const set = state.led.settings;
    it.companyWarn = !!(set.会社名 && p.company && norm(set.会社名) !== norm(p.company));
    if (it.companyWarn) it.warnings.unshift(`会社名が台帳（${set.会社名}）と異なります：${p.company}`);
    it.warnings.push(...continuity(p));
    const firstSame = state.imports.find(o => o !== it && o.parsed && o.parsed.period === p.period && ["new", "diff", "same"].includes(o.status) && state.imports.indexOf(o) < state.imports.indexOf(it));
    if (firstSame) { it.status = "dup"; it.msg = "同じ月のPDFが一覧にもう1件あります。先に並んでいる方を反映すると、こちらと比較できます。"; return; }
    if (!ledgerRows(p.period).length) { it.status = "new"; it.diff = null; return; }
    it.diff = computeDiff(p);
    it.status = it.diff.total ? "diff" : "same";
  }

  async function addFiles(files) {
    const pdfs = [...files].filter(f => /\.pdf$/i.test(f.name) || f.type === "application/pdf");
    if (!pdfs.length) { toast("PDFファイルを選んでください"); return; }
    for (const f of pdfs) {
      const it = { id: ++state.seq, file: f.name, status: "parsing", parsed: null, warnings: [] };
      state.imports.push(it);
      render();
      try {
        const buf = await f.arrayBuffer();
        it.parsed = await TBParser.parse(buf, f.name);
        it.status = "pending";
      } catch (e) {
        console.error(e);
        it.parsed = { fileName: f.name, rows: [], errors: ["読み取り中にエラーが発生しました：" + (e.message || e)], warnings: [] };
        it.status = "pending";
      }
      evaluate(it);
      render();
    }
  }

  async function commit(items, mode, reason) {
    if (state.busy) return;
    state.busy = "台帳に書き込んでいます…"; render();
    const done = [];
    try {
      for (const it of items) {
        const r = await Ledger.writePeriod(it.parsed, { mode, reason, master: state.led.master });
        it.status = "done"; it.impId = r.impId; it.ver = r.ver;
        done.push(it);
      }
    } catch (e) {
      console.error(e);
      toast("書き込みに失敗しました：" + (e.message || e), true);
    }
    state.busy = "";
    await load();
    if (done.length) {
      const res = await reflectMonthly(done.map(i => ({ period: i.parsed.period, impId: i.impId })));
      const ng = res.filter(r => !r.ok).length;
      toast(ng ? `台帳に反映しました。月次シートの検算で確認が必要な月があります（${ng}か月）` : (done.length === 1 ? `${pLabel(done[0].parsed.period)}を反映しました（${done[0].impId}）・検算OK` : `${done.length}か月分を反映しました・検算OK`), !!ng);
    }
    render();
  }

  /* ---------- 月次PL／月次BSへの反映と検算 ---------- */
  async function reflectMonthly(list) {
    state.busy = "月次PL・月次BSに反映しています…"; render();
    const out = [];
    for (const { period, impId } of list.sort((a, b) => (a.period < b.period ? -1 : 1))) {
      const rows = state.led.tb.filter(r => r.period === period);
      let res;
      try {
        res = await MonthlySheets.reflect(period, rows, { fyMonths: state.M.fyMonths(state.M.fyOf(period)) });
      } catch (e) {
        console.error(e);
        res = { period, ok: false, errors: ["月次シートへの反映に失敗しました：" + (e.message || e)], sheets: {}, inserted: [], ngCount: 0 };
      }
      state.checks[period] = res;
      out.push(res);
      const id = impId || (state.led.history.filter(h => h.period === period && h.status === "有効").pop() || {}).id;
      const text = res.errors.length ? "エラー" : res.ngCount ? `NG ${res.ngCount}件` : "OK";
      if (id) { try { await Ledger.setHistCheck(id, text); } catch (e) { console.warn(e); } }
    }
    state.lastChecks = out.map(r => r.period);
    state.busy = "";
    try { state.led = await Ledger.readAll(); } catch (e) {}
    return out;
  }

  /* ---------- 共通UI ---------- */
  let toastTimer = null;
  function toast(msg, isErr) {
    const t = $("#toast");
    t.textContent = msg; t.className = "toast" + (isErr ? " err" : "");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add("hidden"), isErr ? 6000 : 3000);
  }
  const chip = (t, kind = "navy") => `<span class="chip chip-${kind}">${esc(t)}</span>`;
  const fySelect = () => {
    const M = state.M;
    return `<label class="fy-sel">年度<select data-change="fy">${M.fys.slice().reverse().map(f => `<option value="${f}" ${f === state.fy ? "selected" : ""}>${esc(M.fyLabel(f))}</option>`).join("")}</select></label>`;
  };
  const emptyData = () => `<div class="empty"><p>まだ試算表が登録されていません。</p><p class="muted">取込タブで合計残高試算表のPDFを追加すると、ここにBS・PLが表示されます。</p><button type="button" class="btn primary" data-act="tab" data-tab="import">取込タブを開く</button></div>`;
  const dirCls = (d, key) => {
    if (d == null || d === 0) return "";
    const g = state.M.goodUp(key);
    if (g === null) return "";
    return (d > 0) === g ? "up-good" : "up-bad";
  };

  function render() {
    $("#tabs").innerHTML = TABS.map(([id, t]) => `<button type="button" class="tab ${state.tab === id ? "active" : ""}" data-act="tab" data-tab="${id}" aria-current="${state.tab === id}">${t}</button>`).join("");
    const v = $("#view");
    let html = "";
    if (state.loadError) html += `<div class="banner err">${esc(state.loadError)}</div>`;
    try {
      html += ({ import: renderImport, dash: renderDash, monthly: renderMonthly, account: renderAccount, forecast: renderForecast, plan: renderPlan, yojitsu: renderYojitsu }[state.tab] || renderImport)();
    } catch (e) {
      console.error(e);
      html += `<div class="banner err">画面の表示中にエラーが発生しました：${esc(e.message)}</div>`;
    }
    v.innerHTML = html;
    $("#busy").classList.toggle("hidden", !state.busy);
    $("#busy-msg").textContent = state.busy;
  }

  /* ---------- 取込タブ ---------- */
  function renderImport() {
    if (state.diffId) { const it = state.imports.find(i => i.id === state.diffId); if (it && it.diff) return renderDiff(it); state.diffId = null; }
    const M = state.MT, led = state.led;
    const newOnes = state.imports.filter(i => i.status === "new" && !i.companyWarn);
    const items = state.imports.map(renderItem).join("");
    return `
<section class="drop" id="drop">
  <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M12 17v-6"/><path d="M9 14l3-3 3 3"/></svg>
  <div class="drop-text"><strong>PDFをここにドロップ</strong><span>複数月をまとめて追加できます</span></div>
  <label class="btn">ファイルを選択<input type="file" id="file-input" accept=".pdf,application/pdf" multiple hidden></label>
</section>
${state.imports.length ? `
<section class="card">
  <div class="card-head"><h2>読み込んだPDF（${state.imports.length}件）</h2><button type="button" class="link" data-act="clear">一覧を空にする</button></div>
  <div class="items">${items}</div>
  <div class="bulk">
    <p class="muted">差分がある月は、差分確認で「差し替える」を選ぶまで反映しません。同じ内容のPDFは反映不要と表示します。</p>
    <button type="button" class="btn primary" data-act="commit-new" ${newOnes.length ? "" : "disabled"}>${newOnes.length ? `未登録の${newOnes.length}か月を反映` : "反映できる月はありません"}</button>
  </div>
</section>` : ""}
${renderChecks()}
<section class="card">
  <div class="card-head"><h2>登録状況</h2><span class="muted small">期首 ${M.start}月（設定シート）</span></div>
  ${renderGrid()}
</section>
<section class="card">
  <div class="card-head"><h2>取込履歴</h2>${led.exists ? `<button type="button" class="link" data-act="sheet" data-sheet="取込履歴">シートを開く</button>` : ""}</div>
  ${led.history.length ? `<div class="hist">${led.history.slice(-6).reverse().map(h => `<div class="hist-row"><span class="hist-p">${esc(pLabel(h.period))}</span><span class="muted small">${esc(h.at)}・第${h.ver}版${h.reason ? "・" + esc(h.reason) : ""}${h.check ? "・検算 " + esc(h.check) : ""}</span>${chip(h.status, h.status === "有効" ? "navy" : "gray")}</div>`).join("")}</div>` : `<p class="muted">まだ取り込んだPDFはありません。最初の取込で台帳用のシート（TB_明細・取込履歴・TB_退避・科目マスタ・設定・科目対応）を作り、月次PL・月次BSシートの該当月の列に書き込みます。</p>`}
  ${led.exists ? `<div class="sheet-links"><button type="button" class="btn" data-act="sheet" data-sheet="月次PL">月次PL</button><button type="button" class="btn" data-act="sheet" data-sheet="月次BS">月次BS</button><button type="button" class="btn" data-act="sheet" data-sheet="科目対応">科目対応</button><button type="button" class="btn" data-act="sheet" data-sheet="TB_明細">TB_明細</button></div>
  <div class="reflect">
    <label class="lbl" for="re-period">月次シートへ反映し直す</label>
    <div class="reflect-row"><select id="re-period" class="inp">${M.periods.slice().reverse().map(p => `<option value="${p}">${esc(pLabel(p))}</option>`).join("")}</select><button type="button" class="btn" data-act="re-reflect">反映</button></div>
    <p class="muted small">科目対応シートを直したときや、月の列を追加したあとに使います。台帳（TB_明細）の値で上書きします。</p>
  </div>` : ""}
</section>`;
  }

  function renderItem(it) {
    const p = it.parsed;
    const nRows = p ? p.rows.length : 0;
    const bal = p && p.checks ? (p.checks.balance === true ? "貸借一致" : p.checks.balance === false ? "貸借不一致" : "") : "";
    let st = "", act = "";
    switch (it.status) {
      case "parsing": st = chip("読み取り中", "gray"); break;
      case "new": st = chip("未登録の月"); act = it.companyWarn ? `<button type="button" class="btn" data-act="commit-one" data-id="${it.id}">会社名を確認して反映</button>` : `<span class="muted small">反映待ち</span>`; break;
      case "diff": st = chip(`登録済と差分 ${it.diff.total}件`, "amber"); act = `<button type="button" class="btn outline" data-act="diff" data-id="${it.id}">差分を確認</button>`; break;
      case "same": st = chip("登録済と同じ", "gray"); act = `<span class="muted small">反映不要</span>`; break;
      case "dup": st = chip("同じ月が重複", "amber"); break;
      case "done": st = chip(`反映済 ${it.impId}`, "navy"); break;
      case "skip": st = chip("反映しませんでした", "gray"); break;
      case "needPeriod": st = chip("要確認", "err"); act = `<span class="period-fix"><input type="month" id="pm-${it.id}" aria-label="期間"><button type="button" class="btn" data-act="set-period" data-id="${it.id}">この月で読む</button></span>`; break;
      default: st = chip("読み取り失敗", "err");
    }
    const warn = (it.warnings || []).map(w => `<li>${esc(w)}</li>`).join("");
    return `<div class="item">
  <div class="item-top"><span class="fname" title="${esc(it.file)}">${esc(it.file)}</span><button type="button" class="icon-x" data-act="remove" data-id="${it.id}" aria-label="一覧から外す">×</button></div>
  <div class="item-meta ${it.status === "error" || it.status === "needPeriod" ? "err-text" : ""}">${p ? esc(pLabel(p.period)) : ""}${nRows ? `・${nRows}行` : ""}${bal ? `・${bal}` : ""}${p && p.company ? `・${esc(p.company)}` : ""}</div>
  ${it.msg && ["error", "needPeriod", "dup"].includes(it.status) ? `<div class="item-msg">${esc(it.msg)}</div>` : ""}
  <div class="item-act">${st}${act}</div>
  ${warn ? `<ul class="warns">${warn}</ul>` : ""}
</div>`;
  }

  function renderChecks() {
    const list = state.lastChecks.map(p => state.checks[p]).filter(Boolean);
    if (!list.length) return "";
    return list.map(res => {
      const sheetBlock = (sh, title) => {
        const info = res.sheets[sh];
        if (!info) return "";
        if (info.error) return "";
        const rows = (info.checks || []).map(c => `<div class="crow ${c.status}">
          <span class="strong">${esc(c.label)}${c.note ? `<small class="muted block">${esc(c.note)}</small>` : ""}</span>
          <span class="num">${c.sheet == null ? "—" : yen(c.sheet)}</span>
          <span class="num muted">${c.pdf == null ? "—" : yen(c.pdf)}</span>
          <span class="num strong">${c.status === "ok" ? "一致" : c.status === "ng" ? sgn(c.diff) : "—"}</span></div>`).join("");
        return `<div class="csheet"><div class="card-head"><h3>${esc(info.name)}${info.created ? chip("新規作成", "navy") : ""}</h3><span class="muted small">${esc(info.col || "")}列・${info.written}行に書き込み</span></div>
          <div class="ctable"><div class="crow chead"><span>検算項目</span><span class="num">シート</span><span class="num">PDF</span><span class="num">差額</span></div>${rows}</div>
          ${info.noSource && info.noSource.length ? `<p class="muted small">取込元がない行（書き込みなし）：${esc(info.noSource.join("、"))}</p>` : ""}</div>`;
      };
      const head = res.errors.length ? chip("反映できませんでした", "err") : res.ngCount ? chip(`検算NG ${res.ngCount}件`, "amber") : chip("検算OK", "navy");
      return `<section class="card">
  <div class="card-head"><h2>${esc(pLabel(res.period))}の月次シート検算</h2>${head}</div>
  ${res.errors.map(e => `<div class="banner err">${esc(e)}</div>`).join("")}
  ${res.inserted.length ? `<div class="note">科目が増えたため行を追加しました：${res.inserted.map(x => `${esc(x.sheet)} ${x.row}行目「${esc(x.name)}」`).join("、")}。合計行の数式に含まれているか検算で確認してください。</div>` : ""}
  ${sheetBlock("PL")}${sheetBlock("BS")}
  ${res.ngCount ? `<p class="muted small">差額がある場合は「科目対応」シートで行の取込元を直してから「月次シートへ反映し直す」を実行してください。</p>` : ""}
</section>`;
    }).join("");
  }

  function renderGrid() {
    const M = state.MT;
    const pend = {};
    state.imports.forEach(i => { if (i.parsed && i.parsed.period && ["new", "diff"].includes(i.status)) pend[i.parsed.period] = i.status; });
    const fys = new Set(M.fys);
    Object.keys(pend).forEach(p => fys.add(M.fyOf(p)));
    if (!fys.size) fys.add(M.fyOf(new Date().getFullYear() + "-" + String(new Date().getMonth() + 1).padStart(2, "0")));
    const list = [...fys].sort().reverse().slice(0, 4);
    const rows = list.map(fy => {
      const cells = M.fyMonths(fy).map(p => {
        const s = pend[p] === "diff" ? "diff" : M.has(p) ? "done" : pend[p] === "new" ? "new" : "empty";
        const lbl = `${pLabel(p)}：${{ done: "登録済", diff: "差分確認待ち", new: "今回追加", empty: "未登録" }[s]}`;
        return s === "done" ? `<button type="button" class="cell done" data-act="goto-month" data-p="${p}" title="${lbl}" aria-label="${lbl}">${Number(p.slice(5))}</button>`
          : `<span class="cell ${s}" title="${lbl}">${Number(p.slice(5))}</span>`;
      }).join("");
      return `<div class="fy-row"><span class="fy-lbl">${esc(M.fyLabel(fy))}</span><div class="cells">${cells}</div></div>`;
    }).join("");
    return rows + `<div class="legend"><span><i class="lg lg-done"></i>登録済</span><span><i class="lg lg-diff"></i>差分確認待ち</span><span><i class="lg lg-new"></i>今回追加</span><span><i class="lg lg-empty"></i>未登録</span></div>`;
  }

  function renderDiff(it) {
    const d = it.diff, p = it.parsed;
    const cur = state.led.history.filter(h => h.period === p.period && h.status === "有効").pop();
    const rows = d.list.filter(r => state.diffAll || r.type !== "same");
    const typeNote = (r) => {
      if (r.type === "pdf") return "PDFにだけある科目";
      if (r.type === "led") return "台帳にだけある科目";
      if (r.type === "same") return "";
      const parts = [];
      if (r.a.dr !== r.b.dr) parts.push("借方 " + sgn(r.b.dr - r.a.dr));
      if (r.a.cr !== r.b.cr) parts.push("貸方 " + sgn(r.b.cr - r.a.cr));
      if (r.a.open !== r.b.open) parts.push("繰越 " + sgn(r.b.open - r.a.open));
      return parts.join("・");
    };
    const imp = Object.entries(d.impact).map(([l, v]) => `<div class="kv"><span>${l}</span><strong>${sgn(v)}</strong></div>`).join("");
    return `
<div class="diff-head">
  <button type="button" class="link" data-act="diff-close">← 読み込んだPDFへ戻る</button>
  <h2 class="title">${esc(pLabel(p.period))}の差分確認</h2>
  <p class="muted small">登録済（${cur ? esc(cur.at) + "・第" + cur.ver + "版" : "登録済データ"}）と ${esc(p.fileName)} を科目ごとに比較しています</p>
</div>
<section class="stats">
  <div><strong>${d.counts.chg}</strong><span>金額が変わった</span></div>
  <div><strong class="t-navy">${d.counts.pdf}</strong><span>PDFにだけある</span></div>
  <div><strong class="t-amber">${d.counts.led}</strong><span>台帳にだけある</span></div>
  <div><strong class="t-muted">${d.counts.same}</strong><span>一致</span></div>
</section>
<section class="card">
  <div class="card-head"><h2>差し替えた場合の影響</h2><span class="muted small">残高の変化（円）</span></div>
  ${imp || `<p class="muted">主要項目への影響はありません。</p>`}
  <div class="kv muted"><span>PDFの貸借</span><strong>${p.checks.balance === true ? "一致" : p.checks.balance === false ? "不一致" : "確認できません"}</strong></div>
</section>
<section class="card">
  <div class="card-head">
    <div class="seg" role="group" aria-label="表示する行"><button type="button" class="${!state.diffAll ? "on" : ""}" data-act="diff-all" data-v="0">差分のみ（${d.total}）</button><button type="button" class="${state.diffAll ? "on" : ""}" data-act="diff-all" data-v="1">すべて（${d.list.length}）</button></div>
  </div>
  <div class="dtable">
    <div class="drow dhead"><span>科目</span><span>台帳</span><span>PDF</span><span>差額</span></div>
    ${rows.map(r => `<div class="drow ${r.kind === "集計" ? "agg" : ""}">
      <span class="dname"><small>${r.sheet}${r.code ? " " + esc(r.code) : ""}</small>${esc(r.name)}${typeNote(r) ? `<em>${esc(typeNote(r))}</em>` : ""}</span>
      <span class="num muted">${r.a ? yen(r.a.bal) : "—"}</span><span class="num">${r.b ? yen(r.b.bal) : "—"}</span>
      <span class="num strong ${r.type === "pdf" ? "t-navy" : r.type === "led" ? "t-amber" : ""}">${r.type === "same" ? "" : sgn(r.d)}</span></div>`).join("")}
  </div>
</section>
<div class="note">差し替えは月単位です。${esc(pLabel(p.period))}の登録済の行を「TB_退避」シートに移してから、PDFの${p.rows.length}行を書き込みます。同じ月が二重に残ることはありません。一部の科目だけを採用すると貸借が合わなくなるため、行ごとの選択はできません。</div>
<section class="card">
  <label for="reason" class="lbl">差し替えの理由（任意）</label>
  <input type="text" id="reason" class="inp" placeholder="例：決算整理仕訳の反映">
  <p class="muted small">取込履歴に第${(cur ? cur.ver : 1) + 1}版として記録されます</p>
  <div class="btn-col">
    <button type="button" class="btn primary" data-act="replace" data-id="${it.id}">${esc(pLabel(p.period))}を差し替える</button>
    <button type="button" class="btn" data-act="skip" data-id="${it.id}">今回は反映しない</button>
  </div>
</section>`;
  }

  /* ---------- 表示の切り替え・計画・異常値 ---------- */
  const TOGGLES = [["prev", "前期と比較"], ["anom", "異常値"], ["fc", "予測"], ["plan", "計画"]];
  const toggles = (keys) => `<div class="toggles" role="group" aria-label="グラフに重ねる内容">${TOGGLES.filter(([k]) => keys.includes(k)).map(([k, t]) => `<button type="button" class="tgl ${state.show[k] ? "on" : ""}" data-act="toggle" data-k="${k}" aria-pressed="${!!state.show[k]}">${t}</button>`).join("")}</div>`;
  function planFor(fy) {
    const pl = state.plan;
    if (!pl || !pl.exists || !pl.rows) return null;
    const ms = state.M.fyMonths(fy);
    if (!ms.every(p => pl.periods.includes(p))) return null;
    const idx = ms.map(p => pl.periods.indexOf(p));
    const by = {};
    pl.rows.forEach(r => { by[r.key] = { name: r.name, kind: r.kind, values: idx.map(i => r.values[i]) }; });
    return by;
  }
  const minAbs = () => Number(state.led.settings["異常値の下限額"]) || 50000;
  let anomCache = { k: "", v: [] };
  function anomaliesOf(fy) {
    const key = fy + "|" + state.M.periods.length + "|" + state.M.latest + "|" + minAbs();
    if (anomCache.k !== key) anomCache = { k: key, v: FinModel.anomalies(state.M, fy, { minAbs: minAbs() }) };
    return anomCache.v;
  }
  const anomText = (a) => `${a.type === "yoy" ? "前年同月" : "平常"} ${yen(a.ref)}円 → ${yen(a.v)}円`;
  // 当期の途中なら残り月の予測を返す（当期が揃っていれば null）
  function midForecast(key, fy) {
    const M = state.M;
    if (!key || M.fyOf(M.latest) !== fy) return null;
    const f = FinModel.forecast(M, key, state.fcMethod);
    if (!f.ok || f.targetFy !== fy) return f.ok ? null : f;
    return f;
  }

  /* ---------- 全体タブ ---------- */
  function kpi(label, val, unit, cmp, cls) {
    return `<div class="kpi"><span class="kpi-l">${label}</span><span class="kpi-v">${val}<small>${unit}</small></span><span class="kpi-c ${cls || ""}">${cmp}</span></div>`;
  }
  function renderDash() {
    const M = state.M;
    if (!M.latest) return emptyData();
    const fy = state.fy, p = M.latestIn(fy), py = M.addM(p, -12), K = M.K;
    const prevHas = M.has(py);
    const months = M.fyMonths(fy);
    const reg = months.filter(q => M.has(q)).length;
    const plan = state.show.plan ? planFor(fy) : null;
    const planCum = (key) => { if (!plan || !plan[key]) return null; return months.filter(q => q <= p).reduce((t, q, i) => t + (plan[key].values[months.indexOf(q)] || 0), 0); };
    const fcS = state.show.fc ? midForecast(K.sales, fy) : null, fcO = state.show.fc ? midForecast(K.op, fy) : null;
    const salesYoY = prevHas ? pct(M.bal(K.sales, p), M.bal(K.sales, py)) : null;
    const opYoY = prevHas ? pct(M.bal(K.op, p), M.bal(K.op, py)) : null;
    const eq = (q) => (M.bal(K.equity, q) != null && M.bal(K.assets, q) ? M.bal(K.equity, q) / M.bal(K.assets, q) * 100 : null);
    const er = eq(p), erP = prevHas ? eq(py) : null;
    const cash = M.bal(K.cash, p), cashO = M.open(K.cash, p);
    const sub = (key, yoy, fc) => {
      const lines = [];
      lines.push(yoy != null ? `前年同期 <b class="${yoy > 0 ? "up-good" : yoy < 0 ? "up-bad" : ""}">${spct(yoy)}</b>` : `前年データなし`);
      const pc = planCum(key);
      if (pc != null) { const a = M.bal(key, p); const r = pc ? a / pc * 100 : null; lines.push(`計画比 <b class="${a >= pc ? "up-good" : "up-bad"}">${r == null ? "—" : r.toFixed(1) + "%"}</b>`); }
      if (fc && fc.ok) lines.push(`着地見込 <b>${yen(fc.total)}円</b>`);
      return lines.join("<br>");
    };
    const kpi2 = (label, val, unit, html) => `<div class="kpi"><span class="kpi-l">${label}</span><span class="kpi-v">${val}<small>${unit}</small></span><span class="kpi-c">${html}</span></div>`;
    // グラフ
    const cur = months.map(q => (M.has(q) ? M.month(K.sales, q) : null));
    const lines = [];
    if (state.show.prev) lines.push({ values: months.map(q => M.addM(q, -12)).map(q => (M.has(q) ? M.month(K.sales, q) : null)), color: FinCharts.C.gray, width: 2, dash: "4 3" });
    if (plan && plan[K.sales]) lines.push({ values: plan[K.sales].values, color: "#2E7D6B", width: 2, dash: "", dots: false });
    let band = null, marker = null;
    if (fcS && fcS.ok) {
      const fp = {}; fcS.points.forEach(x => fp[x.p] = x);
      lines.push({ values: months.map(q => (fp[q] ? fp[q].v : q === p ? M.month(K.sales, q) : null)), color: FinCharts.C.navy, width: 2, dash: "6 4" });
      band = { lo: months.map(q => (fp[q] ? fp[q].lo : null)), hi: months.map(q => (fp[q] ? fp[q].hi : null)) };
      marker = months.indexOf(p);
    }
    const anoms = state.show.anom ? anomaliesOf(fy) : [];
    const marks = anoms.filter(a => a.key === K.sales).map(a => ({ i: months.indexOf(a.p), v: a.v, title: anomText(a) }));
    const legend = [`<span><i class="lg lg-done"></i>${esc(M.fyLabel(fy))}</span>`];
    if (state.show.prev) legend.push(`<span><i class="lg lg-line"></i>前期</span>`);
    if (plan && plan[K.sales]) legend.push(`<span><i class="lg lg-line" style="border-top:2px solid #2E7D6B"></i>計画</span>`);
    if (fcS && fcS.ok) legend.push(`<span><i class="lg lg-line" style="border-top:2px dashed #1F3F6E"></i>予測</span>`);
    if (marks.length) legend.push(`<span><i class="lg lg-ring"></i>異常値</span>`);
    const bsA = [["流動資産", M.bal(K.curA, p), "#1F3F6E"], ["固定資産", (M.bal(K.fixA, p) || 0) + (M.bal(K.defA, p) || 0), "#5F7FA8"]];
    const bsL = K.curL ? [["流動負債", M.bal(K.curL, p), "#9A560A"], ["固定負債", M.bal(K.fixL, p), "#C28A45"], ["純資産", M.bal(K.equity, p), "#56606B"]]
      : [["負債", M.bal(K.liab, p), "#9A560A"], ["純資産", M.bal(K.equity, p), "#56606B"]];
    const bar = (label, segs) => {
      const tot = segs.reduce((t, x) => t + Math.max(0, x[1] || 0), 0) || 1;
      return `<div class="bsbar"><span class="muted small">${label} ${yen(segs.reduce((t, x) => t + (x[1] || 0), 0))}円</span><div class="bsbar-track">${segs.filter(x => (x[1] || 0) > 0).map(x => `<div style="width:${(x[1] / tot * 100).toFixed(1)}%;background:${x[2]}" title="${x[0]} ${yen(x[1])}円"><span>${x[0]}</span></div>`).join("")}</div></div>`;
    };
    const pl = M.keysOf("PL").filter(m => m.kind === "明細");
    let movers, moverNote;
    if (prevHas) {
      movers = pl.map(m => ({ m, v: M.bal(m.key, p), b: M.bal(m.key, py) })).filter(x => x.v != null && x.b != null && x.b !== 0)
        .map(x => ({ ...x, d: x.v - x.b, r: pct(x.v, x.b) })).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 5);
      moverNote = "累計・前年同期比";
    } else {
      movers = pl.map(m => ({ m, v: M.bal(m.key, p) })).filter(x => x.v).sort((a, b) => Math.abs(b.v) - Math.abs(a.v)).slice(0, 5);
      moverNote = "前年データがないため累計金額の大きい順";
    }
    const fcNote = state.show.fc && !(fcS && fcS.ok) ? `<p class="muted small">${fcS && fcS.reason ? esc(fcS.reason) : "当期は12か月の実績が揃っているため、予測は表示しません（翌期の見込は予測タブ）。"}</p>` : "";
    const prevNote = state.show.prev && !months.some(q => M.has(M.addM(q, -12))) ? `<p class="muted small">前期の月次データがありません。前期の試算表PDFを取り込むと重ねて表示します。</p>` : "";
    const planNote = state.show.plan && !plan ? `<p class="muted small">${esc(M.fyLabel(fy))}の計画シートがありません（計画タブで作成できます）。</p>` : "";
    return `
<div class="toolbar">${fySelect()}<span class="muted small">${mLabel(p)}まで（${reg}/12か月）${state.source === "sheet" ? "・月次シート" : ""}</span></div>
${toggles(["prev", "anom", "fc", "plan"])}
<div class="kpis">
  ${kpi2("売上高 累計", yen(M.bal(K.sales, p)), "円", sub(K.sales, salesYoY, fcS))}
  ${kpi2("営業利益 累計", yen(M.bal(K.op, p)), "円", sub(K.op, opYoY, fcO))}
  ${kpi2(`現預金 ${mLabel(p)}末`, yen(cash), "円", cashO != null && cash != null ? `前月末 <b class="${cash - cashO >= 0 ? "up-good" : "up-bad"}">${sgn(cash - cashO)}円</b>` : "")}
  ${kpi2("自己資本比率", er != null ? er.toFixed(1) : "—", "%", erP != null ? `前年同月 <b>${(er - erP >= 0 ? "+" : "−") + Math.abs(er - erP).toFixed(1)}pt</b>` : "前年データなし")}
</div>
<section class="card">
  <div class="card-head"><h2>月次売上高</h2><span class="muted small">単位：円</span></div>
  ${FinCharts.chart({ label: "月次売上高", labels: months.map(q => String(Number(q.slice(5)))), unit: "円", bars: [{ values: cur, color: FinCharts.C.navy }], lines, band, marker, marks })}
  <div class="legend">${legend.join("")}</div>
  ${prevNote}${planNote}${fcNote}
</section>
${state.show.anom ? renderAnomCard(anoms, fy) : ""}
<section class="card">
  <div class="card-head"><h2>${mLabel(p)}末の貸借対照表</h2></div>
  ${bar("資産", bsA)}${bar("負債・純資産", bsL)}
</section>
<section class="card">
  <div class="card-head"><h2>${prevHas ? "前年から大きく動いた科目" : "主な費用・収益"}</h2><span class="muted small">${moverNote}</span></div>
  <div class="movers">${movers.map(x => `<button type="button" class="mover" data-act="open-account" data-key="${esc(x.m.key)}"><span class="strong">${esc(x.m.name)}</span><span class="num muted">${yen(x.v)}円</span><span class="num strong ${x.r != null ? dirCls(x.d, x.m.key) : ""}">${x.r != null ? spct(x.r) : ""}</span></button>`).join("") || `<p class="muted">表示できる科目がありません。</p>`}</div>
</section>`;
  }

  function renderAnomCard(anoms, fy) {
    const item = (a) => `<button type="button" class="mover anom" data-act="open-account" data-key="${esc(a.key)}"><span><span class="strong">${esc(a.name)}</span><small class="muted block">${mLabel(a.p)}${a.n > 1 ? "〜" + mLabel(a.p2) + `（${a.n}か月）` : ""}・${anomText(a)}</small></span><span class="num strong ${a.dev > 0 ? "t-amber" : "t-navy"}">${sgn(a.dev)}円${a.n > 1 ? "/月" : ""}</span><span class="num small muted">${a.type === "yoy" ? "前年比" : "平常比"}</span></button>`;
    // 同じ科目・同じ向きで連続する月は1件にまとめる（例：役員報酬 7月〜9月）
    const merge = (list) => {
      const by = {};
      list.slice().sort((a, b) => (a.key + a.p < b.key + b.p ? -1 : 1)).forEach(a => {
        const g = by[a.key] = by[a.key] || [];
        const last = g[g.length - 1];
        if (last && FinModel.addM(last.p2, 1) === a.p && Math.sign(last.dev) === Math.sign(a.dev) && last.type === a.type) { last.p2 = a.p; last.n++; last.sum += a.dev; }
        else g.push(Object.assign({}, a, { p2: a.p, n: 1, sum: a.dev }));
      });
      return Object.values(by).flat().sort((a, b) => Math.abs(b.sum) - Math.abs(a.sum));
    };
    const pl = merge(anoms.filter(a => a.sheet === "PL")), bs = merge(anoms.filter(a => a.sheet === "BS"));
    const group = (title, list, n) => list.length ? `<h3 class="grp">${title}（${list.length}件）</h3><div class="movers">${list.slice(0, n).map(item).join("")}</div>${list.length > n ? `<p class="muted small">ほか${list.length - n}件は科目タブで確認できます。</p>` : ""}` : "";
    return `<section class="card">
  <div class="card-head"><h2>異常値</h2><span class="muted small">${esc(state.M.fyLabel(fy))}</span></div>
  ${anoms.length ? group("損益", pl, 6) + group("貸借（前月からの増減）", bs, 4) : `<p class="muted">大きく外れた月はありません。</p>`}
  <p class="muted small">平常＝その科目の全期間の中央値。PLは差が${yen(minAbs())}円以上、BSは${yen(minAbs() * 4)}円以上で、ふだんの振れ幅から大きく外れた月か、前年同月から±50%以上動いた月を出しています。下限額は設定シートの「異常値の下限額」で変えられます。</p>
</section>`;
  }

  /* ---------- 月別タブ ---------- */
  function renderMonthly() {
    const M = state.M;
    if (!M.latest) return emptyData();
    const fy = state.fy;
    if (!state.month || M.fyOf(state.month) !== fy) state.month = M.latestIn(fy);
    const p = state.month, py = M.addM(p, -12);
    const chips = M.fyMonths(fy).map(q => `<button type="button" class="mchip ${q === p ? "on" : ""}" data-act="month" data-p="${q}" ${M.has(q) ? "" : "disabled"}>${mLabel(q)}</button>`).join("");
    const filt = (sheet) => M.keysOf(sheet).filter(m => (state.showDetail || m.kind === "集計") && M.row(m.key, p));
    const cum = state.plMode === "cum";
    const an = {};
    if (state.show.anom) anomaliesOf(fy).filter(a => a.p === p).forEach(a => { an[a.key] = a; });
    const flag = (key) => an[key] ? `<span class="warn-ic" title="${esc(anomText(an[key]))}" aria-label="異常値">!</span>` : "";
    const plRows = filt("PL").map(m => {
      const a = cum ? M.bal(m.key, p) : M.month(m.key, p);
      const b = M.has(py) ? (cum ? M.bal(m.key, py) : M.month(m.key, py)) : null;
      const d = b == null || a == null ? null : a - b;
      return `<button type="button" class="trow ${m.kind === "集計" ? "agg" : "det"} ${an[m.key] ? "is-anom" : ""}" data-act="open-account" data-key="${esc(m.key)}"><span>${flag(m.key)}${esc(m.name)}</span><span class="num">${k(a)}</span><span class="num muted">${k(b)}</span><span class="num strong ${dirCls(d, m.key)}">${sk(d)}</span></button>`;
    }).join("");
    const bsRows = filt("BS").map(m => {
      const a = M.bal(m.key, p), b = M.open(m.key, p), d = b == null ? null : a - b;
      return `<button type="button" class="trow ${m.kind === "集計" ? "agg" : "det"} ${an[m.key] ? "is-anom" : ""}" data-act="open-account" data-key="${esc(m.key)}"><span>${flag(m.key)}${esc(m.name)}</span><span class="num">${k(a)}</span><span class="num muted">${k(b)}</span><span class="num strong">${sk(d)}</span></button>`;
    }).join("");
    return `
<div class="toolbar">${fySelect()}<label class="chk"><input type="checkbox" data-change="detail" ${state.showDetail ? "checked" : ""}>明細も表示</label></div>
${toggles(["anom"])}
<div class="mchips">${chips}</div>
<section class="card">
  <div class="card-head"><h2>${mLabel(p)}の損益計算書</h2>
    <div class="seg" role="group" aria-label="集計単位"><button type="button" class="${!cum ? "on" : ""}" data-act="plmode" data-v="month">当月</button><button type="button" class="${cum ? "on" : ""}" data-act="plmode" data-v="cum">累計</button></div></div>
  <div class="table">
    <div class="trow thead"><span>単位：円</span><span class="num">${cum ? "累計" : "当月"}</span><span class="num">前年同月</span><span class="num">増減</span></div>
    ${plRows || `<p class="muted pad">この月の損益データがありません。</p>`}
  </div>
  ${M.has(py) ? "" : `<p class="muted small">前年同月（${esc(pLabel(py))}）が未登録のため、前年の欄は空です。</p>`}
</section>
<section class="card">
  <div class="card-head"><h2>${mLabel(p)}末の貸借対照表</h2></div>
  <div class="table">
    <div class="trow thead"><span>単位：円</span><span class="num">当月末</span><span class="num">前月末</span><span class="num">増減</span></div>
    ${bsRows}
  </div>
  <p class="muted small">行を押すと科目の推移を開きます。${state.show.anom ? "「!」は異常値の月です（明細も表示にすると科目ごとに確認できます）。" : ""}</p>
</section>`;
  }

  /* ---------- 科目タブ ---------- */
  function accountOptions(sel, sheets = ["PL", "BS"]) {
    const M = state.M;
    return sheets.map(sh => `<optgroup label="${sh === "PL" ? "損益計算書" : "貸借対照表"}">${M.keysOf(sh).map(m => `<option value="${esc(m.key)}" ${m.key === sel ? "selected" : ""}>${m.kind === "集計" ? "" : "　"}${esc(m.name)}</option>`).join("")}</optgroup>`).join("");
  }
  function renderAccount() {
    const M = state.M;
    if (!M.latest) return emptyData();
    const key = state.account, m = M.meta[key];
    if (!m) return `<p class="muted pad">科目を選んでください。</p>`;
    const isPL = m.sheet === "PL";
    const fy = state.fy;
    const base = M.fyMonths(fy);
    const labels = base.map(q => String(Number(q.slice(5))));
    const lines = [];
    const legend = [];
    if (state.show.prev) {
      const fys = M.fys.filter(f => f < fy).slice(-2);
      const cols = [FinCharts.C.light, FinCharts.C.gray];
      fys.forEach((f, i) => {
        const off = (fy - f) * 12;
        const col = cols[cols.length - fys.length + i];
        lines.push({ values: base.map(q => { const r = M.addM(q, -off); return M.has(r) ? M.series(key, r) : null; }), color: col, width: 2, dash: i === 0 && fys.length === 2 ? "2 3" : "4 3", dots: true });
        legend.push(`<span><i class="lg lg-line" style="border-top:2px dashed ${col}"></i>${esc(M.fyLabel(f))}</span>`);
      });
    }
    const curVals = base.map(q => (M.has(q) ? M.series(key, q) : null));
    lines.push({ values: curVals, color: FinCharts.C.navy, width: 3, dots: true });
    legend.push(`<span><i class="lg lg-line" style="border-top:3px solid #1F3F6E"></i>${esc(M.fyLabel(fy))}</span>`);
    const plan = state.show.plan && isPL ? planFor(fy) : null;
    if (plan && plan[key]) { lines.push({ values: plan[key].values, color: "#2E7D6B", width: 2 }); legend.push(`<span><i class="lg lg-line" style="border-top:2px solid #2E7D6B"></i>計画</span>`); }
    let band = null, marker = null, fcMsg = "";
    if (state.show.fc && isPL) {
      const f = midForecast(key, fy);
      if (f && f.ok) {
        const fp = {}; f.points.forEach(x => fp[x.p] = x);
        const last = M.latestIn(fy);
        lines.push({ values: base.map(q => (fp[q] ? fp[q].v : q === last ? M.month(key, q) : null)), color: FinCharts.C.navy, width: 2, dash: "6 4" });
        band = { lo: base.map(q => (fp[q] ? fp[q].lo : null)), hi: base.map(q => (fp[q] ? fp[q].hi : null)) };
        marker = base.indexOf(last);
        legend.push(`<span><i class="lg lg-line" style="border-top:2px dashed #1F3F6E"></i>予測</span>`);
        fcMsg = `<p class="muted small">着地見込 ${yen(f.total)}円（幅 ${yen(f.lo)}〜${yen(f.hi)}円）</p>`;
      } else fcMsg = `<p class="muted small">${f && f.reason ? esc(f.reason) : "当期は12か月の実績が揃っているため、予測は表示しません。"}</p>`;
    }
    const anoms = state.show.anom ? anomaliesOf(fy).filter(a => a.key === key) : [];
    const anomAt = {}; anoms.forEach(a => anomAt[a.p] = a);
    const marks = isPL || m.sheet === "BS" ? anoms.map(a => ({ i: base.indexOf(a.p), v: isPL ? a.v : M.bal(key, a.p), title: anomText(a) })) : [];
    if (marks.length) legend.push(`<span><i class="lg lg-ring"></i>異常値</span>`);
    const p = M.latestIn(fy), py = M.addM(p, -12);
    const rows = base.filter(q => M.has(q)).map(q => {
      const a = M.series(key, q), b = M.has(M.addM(q, -12)) ? M.series(key, M.addM(q, -12)) : null;
      const pv = plan && plan[key] ? plan[key].values[base.indexOf(q)] : null;
      const an = anomAt[q];
      return `<div class="trow ${an ? "is-anom" : ""}"><span>${an ? `<span class="warn-ic" title="${esc(anomText(an))}">!</span>` : ""}${mLabel(q)}</span><span class="num">${yen(a)}</span><span class="num muted">${plan ? yen(pv) : yen(b)}</span><span class="num strong ${dirCls(plan ? (pv == null ? null : a - pv) : (b == null ? null : a - b), key)}">${plan ? (pv == null ? "—" : sgn(a - pv)) : (b == null ? "—" : sgn(a - b))}</span></div>`;
    }).join("");
    const cum = isPL ? M.bal(key, p) : null, cumP = isPL && M.has(py) ? M.bal(key, py) : null;
    return `
<div class="toolbar">${fySelect()}</div>
<label class="lbl" for="acc-sel">科目</label>
<select id="acc-sel" class="inp" data-change="account">${accountOptions(key)}</select>
<div class="acc-head">
  <span class="muted small">${isPL ? "損益計算書" : "貸借対照表"}${m.kind === "集計" ? "／集計行" : ""}</span>
  <h2 class="title">${esc(m.name)}</h2>
  ${isPL ? `<span class="muted small">${mLabel(p)}までの累計 ${yen(cum)}円${cumP != null ? "　前年同期 " + spct(pct(cum, cumP)) : ""}</span>` : `<span class="muted small">${mLabel(p)}末残高 ${yen(M.bal(key, p))}円</span>`}
</div>
${toggles(isPL ? ["prev", "anom", "fc", "plan"] : ["prev", "anom"])}
<section class="card">
  <div class="card-head"><h2>${isPL ? "月次発生額" : "月末残高"}</h2><span class="muted small">単位：円</span></div>
  ${FinCharts.chart({ label: m.name + "の推移", labels, unit: "円", lines, band, marker, marks, zero: isPL })}
  <div class="legend">${legend.join("")}</div>
  ${fcMsg}
  ${state.show.prev && !base.some(q => M.has(M.addM(q, -12))) ? `<p class="muted small">前期の月次データがありません。前期の試算表PDFを取り込むと重ねて表示します。</p>` : ""}
</section>
${anoms.length ? `<section class="card"><div class="card-head"><h2>異常値</h2><span class="muted small">${anoms.length}件</span></div>${anoms.map(a => `<div class="kv"><span>${mLabel(a.p)}</span><span>${anomText(a)}</span></div>`).join("")}</section>` : ""}
<section class="card">
  <div class="card-head"><h2>月別の比較</h2><span class="muted small">単位：円</span></div>
  <div class="table"><div class="trow thead"><span>月</span><span class="num">当期</span><span class="num">${plan ? "計画" : "前期"}</span><span class="num">${plan ? "計画差" : "前年差"}</span></div>${rows}</div>
</section>`;
  }

  /* ---------- 予測タブ ---------- */
  function renderForecast() {
    const M = state.M;
    if (!M.latest) return emptyData();
    const key = state.fcKey && M.meta[state.fcKey] && M.meta[state.fcKey].sheet === "PL" ? state.fcKey : M.K.sales;
    const f = FinModel.forecast(M, key, state.fcMethod);
    const methods = [["yoy", "前年同月比"], ["ma", "移動平均"], ["reg", "傾向線"]];
    const seg = `<div class="seg wide" role="group" aria-label="予測の方法">${methods.map(([v, t]) => `<button type="button" class="${state.fcMethod === v ? "on" : ""}" data-act="fcmethod" data-v="${v}">${t}</button>`).join("")}</div>`;
    const name = M.meta[key].name;
    let body = "";
    if (!f.ok) {
      body = `<div class="note">${esc(f.reason)}</div>`;
    } else {
      const months = M.fyMonths(f.targetFy);
      const act = months.map(q => (q <= M.latest && M.has(q) ? M.month(key, q) : null));
      const fpt = {}; f.points.forEach(x => fpt[x.p] = x);
      const lastAct = months.filter(q => q <= M.latest).pop();
      const fc = months.map(q => (fpt[q] ? fpt[q].v : q === lastAct ? M.month(key, q) : null));
      const lo = months.map(q => (fpt[q] ? fpt[q].lo : null)), hi = months.map(q => (fpt[q] ? fpt[q].hi : null));
      const prev = months.map(q => { const r = M.addM(q, -12); return M.has(r) ? M.month(key, r) : null; });
      const marker = lastAct ? months.indexOf(lastAct) : null;
      const prevTotal = (() => { const last = M.fyMonths(f.targetFy - 1).pop(); return M.has(last) ? M.bal(key, last) : null; })();
      const same = f.targetFy === M.fyOf(M.latest);
      const desc = {
        yoy: `直近${f.basis.length}か月の前年同月比（${spct((f.ratio - 1) * 100)}）を、${same ? "残りの月" : "翌期の各月"}の前年実績に掛けています。`,
        ma: `直近${f.basis.length}か月の平均（${yen(f.points[0].v)}円／月）を${same ? "残りの月" : "翌期の各月"}に置いています。`,
        reg: `直近${f.basis.length}か月の傾向（月あたり ${sk(f.slope)}円）を延ばしています。`
      }[state.fcMethod];
      const others = [M.K.sales, M.K.op, M.K.ord].filter(Boolean).map(kk => ({ kk, r: FinModel.forecast(M, kk, state.fcMethod) }));
      body = `
<section class="card">
  <div class="fc-head"><span class="muted small">${esc(M.fyLabel(f.targetFy))} ${esc(name)}の${same ? "着地見込" : "見込（翌期）"}</span>
  <span class="kpi-v big">${mil(f.total)}<small>円</small></span>
  <span class="muted small">幅 ${mil(f.lo)}〜${mil(f.hi)}${prevTotal != null ? "　前期 " + mil(prevTotal) : ""}${same ? `　うち実績 ${mil(f.actualCum)}（${f.actualMonths.length}か月）` : ""}</span></div>
  ${FinCharts.chart({ label: name + "の予測", labels: months.map(q => String(Number(q.slice(5)))), unit: "円", height: 190, lines: [{ values: prev, color: FinCharts.C.gray, width: 1.5, dash: "4 3" }, { values: act, color: FinCharts.C.navy, width: 3, dots: true }, { values: fc, color: FinCharts.C.navy, width: 2, dash: "6 4" }], band: { lo, hi }, marker: same ? marker : null })}
  <div class="legend">${same ? `<span><i class="lg lg-line" style="border-top:3px solid #1F3F6E"></i>実績</span>` : ""}<span><i class="lg lg-line" style="border-top:2px dashed #1F3F6E"></i>予測</span><span><i class="lg lg-band"></i>予測の幅</span><span><i class="lg lg-line"></i>前期</span></div>
</section>
<section class="card">
  <div class="card-head"><h2>主要項目の見込</h2><span class="muted small">単位：円</span></div>
  <div class="table">
    <div class="trow thead"><span>単位：円</span><span class="num">見込</span><span class="num">前期</span><span class="num">増減</span></div>
    ${others.map(({ kk, r }) => { const last = M.fyMonths(r.targetFy - 1).pop(); const pv = r.ok && M.has(last) ? M.bal(kk, last) : null; return `<div class="trow"><span><span class="strong">${esc(M.meta[kk].name)}</span>${r.ok ? `<small class="muted block">${mil(r.lo)}〜${mil(r.hi)}</small>` : ""}</span><span class="num strong">${r.ok ? mil(r.total) : "—"}</span><span class="num muted">${mil(pv)}</span><span class="num strong ${r.ok && pv != null ? dirCls(r.total - pv, kk) : ""}">${r.ok && pv != null ? spct(pct(r.total, pv)) : "—"}</span></div>`; }).join("")}
  </div>
</section>
<p class="muted small pad">${esc(desc)}幅は当てはめの誤差から求めた約80%の範囲です。試算表の実績だけで計算しており、受注残や予算は含みません。</p>`;
    }
    return `
<label class="lbl" for="fc-sel">予測する項目</label>
<select id="fc-sel" class="inp" data-change="fckey">${accountOptions(key, ["PL"])}</select>
<div class="lbl">予測の方法</div>${seg}
${body}`;
  }

  /* ---------- 計画タブ ---------- */
  function renderPlan() {
    const M = state.M;
    if (!M.latest) return emptyData();
    const pl = state.plan;
    const latestFy = M.fyOf(M.latest);
    const targets = [latestFy + 1, latestFy].filter(f => M.fys.includes(f - 1) || f - 1 === latestFy || M.fys.includes(f - 1));
    const tsel = state.planTarget && [latestFy + 1, latestFy].includes(state.planTarget) ? state.planTarget : latestFy + 1;
    const canCreate = !!(state.layout && state.layout.PL);
    const createCard = (again) => `<section class="card">
  <div class="card-head"><h2>${again ? "計画シートを作り直す" : "計画シートを作る"}</h2></div>
  <p class="muted small">前期の月次実績（月次PL）をそのまま各月の基準にして「計画」シートを作ります。シートのC列に調整率、D列に年間の調整額を入れると月の計画が変わります。月のセルを直接書き換えても構いません。</p>
  <label class="lbl" for="plan-fy">計画する期</label>
  <select id="plan-fy" class="inp" data-change="planfy">${[latestFy + 1, latestFy].map(f => `<option value="${f}" ${f === tsel ? "selected" : ""}>${esc(M.fyLabel(f))}（前期＝${esc(M.fyLabel(f - 1))}）</option>`).join("")}</select>
  ${M.fyMonths(tsel - 1).filter(q => M.has(q)).length < 12 ? `<p class="small t-amber">前期（${esc(M.fyLabel(tsel - 1))}）の月次実績は${M.fyMonths(tsel - 1).filter(q => M.has(q)).length}か月分です。ない月の基準は0になります。</p>` : ""}
  ${again ? `<label class="chk"><input type="checkbox" id="plan-confirm">今の計画シートを消して作り直す（入力した調整率・金額は消えます）</label>` : ""}
  <button type="button" class="btn primary" data-act="plan-create" ${canCreate ? "" : "disabled"}>${again ? "作り直す" : "計画シートを作成"}</button>
  ${canCreate ? "" : `<p class="muted small">月次PLシートがないため作れません。先に取込タブで試算表を反映してください。</p>`}
</section>`;
    if (!pl || !pl.exists) return createCard(false);
    if (pl.error) return `<div class="banner err">${esc(pl.error)}</div>` + createCard(true);
    const fy = M.fyOf(pl.periods[0]);
    const months = M.fyMonths(fy);
    const plan = planFor(fy) || {};
    const last = M.latestIn(fy);
    const done = last ? months.filter(q => q <= last) : [];
    const cumOf = (arr, n) => arr.slice(0, n).reduce((t, v) => t + (v || 0), 0);
    const K = M.K;
    const planYear = (key) => plan[key] ? cumOf(plan[key].values, 12) : null;
    const actCum = (key) => (last ? M.bal(key, last) : null);
    const planCumN = (key) => plan[key] ? cumOf(plan[key].values, done.length) : null;
    const kp = (label, key) => {
      const a = actCum(key), pc = planCumN(key), py = planYear(key);
      const r = a != null && pc ? a / pc * 100 : null;
      return `<div class="kpi"><span class="kpi-l">${label}</span><span class="kpi-v">${yen(py)}<small>円</small></span><span class="kpi-c">年間計画<br>${done.length ? `実績 ${yen(a)}円（${done.length}か月）<br>計画比 <b class="${r >= 100 ? "up-good" : "up-bad"}">${r == null ? "—" : r.toFixed(1) + "%"}</b>` : "実績はまだありません"}</span></div>`;
    };
    const salesPlan = plan[K.sales] ? plan[K.sales].values : months.map(() => null);
    const actual = months.map(q => (M.has(q) ? M.month(K.sales, q) : null));
    const planCumLine = []; const actCumLine = [];
    let t1 = 0, t2 = 0;
    months.forEach((q, i) => { t1 += salesPlan[i] || 0; planCumLine.push(t1); if (actual[i] != null) { t2 += actual[i]; actCumLine.push(t2); } else actCumLine.push(null); });
    const cumMode = state.plMode === "cum";
    const rowsSrc = pl.rows.filter(r => state.planDetail || r.kind === "集計" || r.key === K.sales);
    const tRows = rowsSrc.map(r => {
      const pv = cumOf(plan[r.key] ? plan[r.key].values : [], done.length || 0);
      const a = done.length && M.meta[r.key] ? M.bal(r.key, last) : null;
      const d = a == null ? null : a - pv;
      const rt = a != null && pv ? a / pv * 100 : null;
      const big = rt != null && Math.abs(rt - 100) >= 10 && Math.abs(d) >= minAbs();
      return `<button type="button" class="trow ${r.kind === "集計" ? "agg" : "det"} ${big ? "is-anom" : ""}" data-act="open-account" data-key="${esc(r.key)}"><span>${big ? `<span class="warn-ic" title="計画との差が10%以上">!</span>` : ""}${esc(r.name)}</span><span class="num">${done.length ? yen(pv) : yen(planYear(r.key))}</span><span class="num muted">${a == null ? "—" : yen(a)}</span><span class="num strong ${dirCls(d, r.key)}">${d == null ? "—" : sgn(d)}</span></button>`;
    }).join("");
    return `
<div class="toolbar"><span class="strong">${esc(M.fyLabel(fy))}の計画</span><button type="button" class="link" data-act="sheet" data-sheet="${esc(pl.sheet)}">計画シートを開く</button></div>
<div class="kpis">${kp("売上高", K.sales)}${kp("営業利益", K.op)}</div>
<section class="card">
  <div class="card-head"><h2>売上高 計画と実績</h2>
    <div class="seg" role="group" aria-label="表示"><button type="button" class="${!cumMode ? "on" : ""}" data-act="plmode" data-v="month">月次</button><button type="button" class="${cumMode ? "on" : ""}" data-act="plmode" data-v="cum">累計</button></div></div>
  ${FinCharts.chart({ label: "売上高の計画と実績", labels: months.map(q => String(Number(q.slice(5)))), unit: "円", bars: [{ values: cumMode ? actCumLine : actual, color: FinCharts.C.navy }], lines: [{ values: cumMode ? planCumLine : salesPlan, color: "#2E7D6B", width: 2, dots: true }] })}
  <div class="legend"><span><i class="lg lg-done"></i>実績</span><span><i class="lg lg-line" style="border-top:2px solid #2E7D6B"></i>計画</span></div>
</section>
<section class="card">
  <div class="card-head"><h2>${done.length ? `${mLabel(last)}までの計画と実績` : "年間計画"}</h2><label class="chk"><input type="checkbox" data-change="plandetail" ${state.planDetail ? "checked" : ""}>明細も表示</label></div>
  <div class="table"><div class="trow thead"><span>単位：円</span><span class="num">${done.length ? "計画累計" : "年間計画"}</span><span class="num">実績累計</span><span class="num">差額</span></div>${tRows}</div>
  <p class="muted small">「!」は計画との差が10%以上かつ${yen(minAbs())}円以上の科目です。</p>
</section>
${createCard(true)}`;
  }

  /* ---------- 予実タブ（js/yojitsu.js） ---------- */
  // 計画シートの売上内訳（保守・開発(既存)・開発(新規)）を対象期の12か月に並べる
  function planSegFor(fy) {
    const pl = state.plan;
    if (!pl || !pl.exists || !pl.seg) return null;
    const ms = state.M.fyMonths(fy);
    if (!ms.every(p => pl.periods.includes(p))) return null;
    const out = {};
    ["m", "e", "n"].forEach(k => { out[k] = ms.map(p => (pl.seg[k] ? pl.seg[k][pl.periods.indexOf(p)] : null)); });
    return out;
  }
  function renderYojitsu() {
    const M = state.M;
    if (!state.yj) state.yj = Yojitsu.loadParams();
    return Yojitsu.render({
      M, fy: state.fy, plan: M.latest ? planFor(state.fy) : null,
      segPlan: M.latest ? planSegFor(state.fy) : null, segAct: state.segAct, p: state.yj,
      fySelect: M.latest ? fySelect() : "", empty: emptyData(),
      toast, reload: load, rerender: render
    });
  }

  /* ---------- イベント ---------- */
  document.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-act]");
    if (!b || b.disabled) return;
    const act = b.dataset.act, id = Number(b.dataset.id);
    const it = state.imports.find(i => i.id === id);
    switch (act) {
      case "tab": state.tab = b.dataset.tab; state.diffId = state.tab === "import" ? state.diffId : null; persist(); render(); window.scrollTo(0, 0); break;
      case "clear": state.imports = []; state.diffId = null; render(); break;
      case "remove": state.imports = state.imports.filter(i => i.id !== id); state.imports.forEach(evaluate); render(); break;
      case "diff": state.diffId = id; state.diffAll = false; render(); window.scrollTo(0, 0); break;
      case "diff-close": state.diffId = null; render(); break;
      case "diff-all": state.diffAll = b.dataset.v === "1"; render(); break;
      case "commit-new": await commit(state.imports.filter(i => i.status === "new" && !i.companyWarn), "new"); break;
      case "commit-one": if (it) await commit([it], "new"); break;
      case "replace": if (it) { const reason = ($("#reason") || {}).value || ""; state.diffId = null; await commit([it], "replace", reason); } break;
      case "skip": if (it) { it.status = "skip"; state.diffId = null; render(); } break;
      case "set-period": {
        const v = ($("#pm-" + id) || {}).value;
        if (!/^\d{4}-\d{2}$/.test(v || "")) { toast("年月を選んでください"); break; }
        it.parsed.period = v; it.parsed.errors = it.parsed.errors.filter(x => !/期間/.test(x)); it.status = "pending"; it.msg = "";
        evaluate(it); render(); break;
      }
      case "goto-month": { const p = b.dataset.p; state.fy = state.M.fyOf(p); state.month = p; state.tab = "monthly"; persist(); render(); break; }
      case "month": state.month = b.dataset.p; render(); break;
      case "plmode": state.plMode = b.dataset.v; persist(); render(); break;
      case "open-account": state.account = b.dataset.key; state.tab = "account"; persist(); render(); window.scrollTo(0, 0); break;
      case "fcmethod": state.fcMethod = b.dataset.v; persist(); render(); break;
      case "toggle": state.show[b.dataset.k] = !state.show[b.dataset.k]; persist(); render(); break;
      case "plan-create": {
        const fy = Number(($("#plan-fy") || {}).value || state.planTarget);
        const cf = $("#plan-confirm");
        if (cf && !cf.checked) { toast("作り直す場合はチェックを入れてください"); break; }
        if (!state.layout.PL) { toast("月次PLシートが見つかりません", true); break; }
        state.busy = "計画シートを作成しています…"; render();
        try {
          const r = await Plan.create(state.layout.PL, state.M, fy);
          state.plan = await Plan.read();
          toast(`${state.M.fyLabel(fy)}の計画シートを作成しました` + (r.warn.length ? `（数式を写せなかった行：${r.warn.join("、")}）` : ""));
        } catch (err) { console.error(err); toast("計画シートを作れませんでした：" + (err.message || err), true); }
        state.busy = ""; render(); break;
      }
      case "sheet": try { const ok = await MonthlySheets.activateByPrefix(b.dataset.sheet); if (!ok) toast("シートが見つかりません"); } catch (err) { toast("シートを開けませんでした", true); } break;
      case "re-reflect": {
        const p = ($("#re-period") || {}).value;
        if (!p) break;
        const [res] = await reflectMonthly([{ period: p }]);
        toast(res.errors.length ? res.errors[0] : res.ngCount ? `検算NG ${res.ngCount}件` : `${pLabel(p)}を反映しました・検算OK`, !res.ok);
        render(); break;
      }
    }
  });
  document.addEventListener("change", (e) => {
    const t = e.target;
    if (t.id === "file-input") { addFiles(t.files); t.value = ""; return; }
    const c = t.dataset.change;
    if (c === "fy") { state.fy = Number(t.value); state.month = state.M.latestIn(state.fy); render(); }
    else if (c === "detail") { state.showDetail = t.checked; persist(); render(); }
    else if (c === "account") { state.account = t.value; render(); }
    else if (c === "fckey") { state.fcKey = t.value; render(); }
    else if (c === "planfy") { state.planTarget = Number(t.value); render(); }
    else if (c === "plandetail") { state.planDetail = t.checked; persist(); render(); }
  });
  ["dragenter", "dragover"].forEach(ev => document.addEventListener(ev, (e) => {
    if (!e.dataTransfer || ![...(e.dataTransfer.types || [])].includes("Files")) return;
    e.preventDefault();
    const d = $("#drop"); if (d) d.classList.add("over");
  }));
  ["dragleave", "drop"].forEach(ev => document.addEventListener(ev, (e) => {
    const d = $("#drop"); if (d && (ev === "drop" || !e.relatedTarget)) d.classList.remove("over");
  }));
  document.addEventListener("drop", (e) => {
    if (!e.dataTransfer || !e.dataTransfer.files.length) return;
    e.preventDefault();
    if (state.tab !== "import") { state.tab = "import"; state.diffId = null; }
    addFiles(e.dataTransfer.files);
  });

  /* ---------- 共通スライドメニュー ---------- */
  const COMMON_BASE = "https://ymatsuda-cmyk.github.io/tools/addin/common";
  let menuReady = null;
  function openMenu() {
    if (!menuReady) {
      menuReady = new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = COMMON_BASE + "/slide-menu.js";
        s.onload = () => {
          SlideMenu.init({
            appName: "finance2", version: APP_VERSION, position: "left", currentId: "finance2",
            menuUrl: COMMON_BASE + "/menu.json",
            localItems: [
              { label: "再読み込み", icon: "", onClick: () => load() },
              { label: "このブックを開いたら自動で表示：切り替え", icon: "", onClick: async () => { const on = !autoShowState(); const ok = await setAutoShow(on, true); toast(ok ? (on ? "このブックを開いたときに自動で表示します" : "このブックでは自動で表示しません") : "設定を保存できませんでした", !ok); } },
              { label: "画面設定をリセット", icon: "", onClick: () => { localStorage.removeItem(LS); location.reload(); } }
            ]
          });
          resolve();
        };
        s.onerror = (err) => { menuReady = null; reject(err); };
        document.head.appendChild(s);
      });
    }
    menuReady.then(() => SlideMenu.open()).catch(() => { menuReady = null; toast("メニューを読み込めませんでした", true); });
  }

  Office.onReady(() => {
    $("#version-label").textContent = APP_VERSION;
    $("#menu-btn").addEventListener("click", openMenu);
    $("#reload-btn").addEventListener("click", () => load());
    load();
  });
  window.Finance2 = { state, load, render };
})();
