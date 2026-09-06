/* ============================================================
 * 提案ナレッジ アドイン app.js（詳細UI版）
 * ------------------------------------------------------------
 * Excel操作・AI呼び出しの実体は roi-core.js に一本化した。
 * このファイルはUIのみを担当し、営業報告アドイン（簡単な操作の
 * みを提供）とは異なり、以下の詳細機能を持つ:
 *   ・AI抽出結果を保存前に確認・修正するレビュー画面
 *   ・課題×解決案の一覧・提案書への採用チェック
 *   ・提案書作成用プロンプトの組み立て・コピー
 * ROIマスタの追加・数式変更は、Excel上でシートを直接編集する
 * （このアドインにマスタ編集専用画面は設けていない）。
 * ============================================================ */

let demoMode = false;

if (window.Office) {
  Office.onReady(() => whenDomReady(init));
} else {
  window.addEventListener("DOMContentLoaded", () => init());
}

function whenDomReady(fn) {
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", fn, { once: true });
  else fn();
}

async function init() {
  bindStaticUI();
  const { demo } = await RoiCore.ensureAllSheets();
  demoMode = demo;
  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
  await renderCaseIdList();
}

function bindStaticUI() {
  document.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", () => switchTab(btn.dataset.tab)));
  document.getElementById("reload-btn").addEventListener("click", () => location.reload());
  document.getElementById("settings-btn").addEventListener("click", openSettings);
  document.getElementById("cfg-close-btn").addEventListener("click", closeSettings);
  document.getElementById("cfg-save-btn").addEventListener("click", saveSettings);

  ["case-select", "case-select-2", "case-select-3"].forEach(id => {
    document.getElementById(id).addEventListener("change", async (e) => {
      await selectCase(e.target.value);
      if (id !== "case-select" && e.target.value) await loadProposal();
    });
  });

  document.getElementById("add-log-btn").addEventListener("click", () => {
    const f = document.getElementById("log-form");
    f.style.display = f.style.display === "none" ? "" : "none";
  });
  document.getElementById("cancel-log-btn").addEventListener("click", () => {
    document.getElementById("log-form").style.display = "none";
  });
  document.getElementById("save-log-btn").addEventListener("click", saveHearingLog);
  document.getElementById("extract-btn").addEventListener("click", runExtraction);
  document.getElementById("apply-extract-btn").addEventListener("click", applyExtraction);
  document.getElementById("cancel-extract-btn").addEventListener("click", cancelExtraction);

  document.getElementById("build-prompt-btn").addEventListener("click", buildPrompt);

  // カテゴリ割り当てモーダル
  document.getElementById("assign-close-btn").addEventListener("click", closeAssignModal);
  document.getElementById("assign-apply-btn").addEventListener("click", applyCandidate);
  document.getElementById("assign-chat-btn").addEventListener("click", () => startCategoryChat(null));
  document.getElementById("assign-nocat-btn").addEventListener("click", () =>
    finishAssign("", RoiCore.ASSIGN_NOCAT));
  document.getElementById("chat-send-btn").addEventListener("click", async () => {
    const inp = document.getElementById("chat-input");
    const v = inp.value.trim();
    if (!v) return;
    assignState.messages.push({ role: "user", content: v });
    inp.value = "";
    renderChat();
    await sendCategoryChat();
  });
  document.getElementById("chat-save-common-btn").addEventListener("click", () =>
    saveCategory(RoiCore.SCOPE_COMMON));
  document.getElementById("chat-save-temp-btn").addEventListener("click", () =>
    saveCategory(RoiCore.SCOPE_TEMP));
}

function switchTab(tab) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".pane").forEach(p => p.classList.remove("active"));
  document.getElementById("pane-" + tab).classList.add("active");
}

/* ---------- 設定 ---------- */
function openSettings() {
  const cfg = RoiCore.getConfig();
  document.getElementById("cfg-webhook").value = cfg.webhookUrl || "";
  document.getElementById("cfg-token").value = cfg.token || "";
  document.getElementById("settings-modal").style.display = "flex";
}
function closeSettings() { document.getElementById("settings-modal").style.display = "none"; }
function saveSettings() {
  RoiCore.setConfig({
    webhookUrl: document.getElementById("cfg-webhook").value.trim(),
    token: document.getElementById("cfg-token").value.trim(),
  });
  closeSettings();
}

/* ---------- 候補一覧 ---------- */
let caseOptions = [];

async function renderCaseIdList() {
  caseOptions = demoMode
    ? [{ caseId: "KM-01", label: "KM-01 ／ kakimoto arms" }, { caseId: "OF-02", label: "OF-02 ／ 大石フーズ" }]
    : await RoiCore.listCaseIds();
  const opts = `<option value="">案件を選択してください</option>` +
    caseOptions.map(c => `<option value="${escAttr(c.caseId)}">${escHtml(c.label)}</option>`).join("");
  ["case-select", "case-select-2", "case-select-3"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = opts;
  });
  await renderCaseOverview();
}

function caseLabelOf(caseId) {
  const c = caseOptions.find(x => x.caseId === caseId);
  return c ? c.label : caseId;
}

/* 案件を選ぶ（3タブ間で共有） */
async function selectCase(caseId) {
  ["case-id", "case-id-2", "case-id-3"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = caseId;
  });
  ["case-select", "case-select-2", "case-select-3"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = caseId;
  });
  document.getElementById("hearing-area").style.display = caseId ? "" : "none";
  if (caseId) await loadHearingLogForCase();
}

/* ---------- 案件の状況（未抽出はグレー表示、タップで②タブへ） ---------- */
async function renderCaseOverview() {
  const el = document.getElementById("case-overview");
  if (!el) return;
  const stats = demoMode ? [] : await RoiCore.listCasesWithHearings();
  if (!stats.length) { el.innerHTML = `<div class="meta">まだ議事録が登録されていません</div>`; return; }
  stats.sort((a, b) => b.issueCount - a.issueCount);
  el.innerHTML = stats.map(s => {
    const done = s.issueCount > 0;
    return `
    <div class="case-row${done ? " done" : ""}" data-case="${escAttr(s.caseId)}">
      <div class="c-main">
        <div class="c-id">${escHtml(s.caseId)}</div>
        <div class="c-title">${escHtml(stripId(caseLabelOf(s.caseId)))}</div>
        <div class="c-sub">議事録 ${s.hearingCount}件</div>
      </div>
      <div class="c-count">
        ${done ? `<b>${s.issueCount}</b><span>課題</span>` : `<span class="c-none">未抽出</span>`}
      </div>
    </div>`;
  }).join("");

  el.querySelectorAll(".case-row").forEach(row => {
    row.addEventListener("click", async () => {
      const caseId = row.dataset.case;
      await selectCase(caseId);
      // 抽出済みなら②提案タブへ、未抽出なら①に留まり議事録を見せる
      if (row.classList.contains("done")) {
        switchTab("proposal");
        await loadProposal();
      }
    });
  });
}
function stripId(label) {
  const i = label.indexOf("／");
  return i >= 0 ? label.slice(i + 1).trim() : label;
}

/* ============================================================
   ① 議事録入力
   ============================================================ */
async function saveHearingLog() {
  const caseId = document.getElementById("case-id").value.trim();
  const title = document.getElementById("log-title").value.trim();
  const text = document.getElementById("hearing-text").value.trim();
  const url = document.getElementById("hearing-url").value.trim();
  if (!caseId) { setStatus("案件を選択してください"); return; }
  if (!text && !url) { setStatus("本文かURLのどちらかを入力してください"); return; }

  if (demoMode) { setStatus("デモモードのため保存はシミュレーションのみです"); }
  else {
    await RoiCore.appendHearingLog(caseId, title || "議事録", { text, url });
    setStatus("議事録に追加しました");
  }
  ["log-title", "hearing-text", "hearing-url"].forEach(id => { document.getElementById(id).value = ""; });
  document.getElementById("log-form").style.display = "none";
  await loadHearingLogForCase();
  await renderCaseOverview();
}

async function loadHearingLogForCase() {
  const caseId = document.getElementById("case-id").value.trim();
  const list = document.getElementById("hearing-log-list");
  if (!caseId) { list.innerHTML = ""; return; }
  const rows = demoMode ? [] : await RoiCore.listHearingLogs(caseId);
  document.getElementById("hearing-count").textContent = `この案件の議事録（${rows.length}件）`;
  document.getElementById("extract-btn").textContent =
    rows.length ? `${rows.length}件の議事録から課題を抽出` : "課題を抽出";
  list.innerHTML = rows.map(r => `
    <div class="log-item">
      <div class="log-top">
        <span class="log-title">${escHtml(r.title)}</span>
        <span class="log-date">${escHtml(fmtMonthDay(r.registeredAt))}</span>
      </div>
      ${r.url ? `<div class="log-url"><a href="${escAttr(r.url)}" target="_blank" rel="noopener">${escHtml(r.url)}</a></div>` : ""}
      ${r.text ? `<div class="log-text">${escHtml(r.text.slice(0, 60))}${r.text.length > 60 ? "…" : ""}</div>` : ""}
    </div>`).join("") || `<div class="meta">まだ議事録がありません</div>`;
}

/* ---------- 課題抽出（差分確認つき） ---------- */
let pendingExtraction = null;

async function runExtraction() {
  const caseId = document.getElementById("case-id").value.trim();
  if (!caseId) { setStatus("案件を選択してください"); return; }
  const cfg = RoiCore.getConfig();
  if (!cfg.webhookUrl) { setStatus("設定（⚙）でAI連携エンドポイントを登録してください"); return; }

  setStatus("議事録から課題を抽出中…");
  try {
    const preview = demoMode ? { hearingIds: [], results: [] } : await RoiCore.previewExtraction(caseId);
    if (!preview.results.length) { setStatus("課題は抽出されませんでした"); return; }
    pendingExtraction = { caseId, ...preview };
    renderExtractDiff();
    setStatus("");
  } catch (e) {
    console.warn(e);
    setStatus(e.message || "抽出に失敗しました。エンドポイントの設定を確認してください");
  }
}

function renderExtractDiff() {
  const box = document.getElementById("extract-review");
  const wrap = document.getElementById("extract-items");
  wrap.innerHTML = pendingExtraction.results.map((r, i) => {
    if (r.status === "新規") {
      return `
      <div class="diff-card new">
        <div class="diff-head"><span class="diff-cat">${escHtml(r.category)}</span>
          <span class="tag-new">新規</span></div>
        <div class="diff-title">${escHtml(r.newTitle)}</div>
        <div class="diff-text">${escHtml(r.newSummary)}</div>
      </div>`;
    }
    if (r.status === "未編集") {
      return `
      <div class="diff-card">
        <div class="diff-head"><span class="diff-cat">${escHtml(r.category)}</span>
          <span class="tag-plain">未編集</span></div>
        <div class="diff-title">${escHtml(r.newTitle)}</div>
        <div class="diff-text">${escHtml(r.newSummary)}</div>
        <div class="hint-sm">編集されていないため自動で更新します</div>
      </div>`;
    }
    return `
    <div class="diff-card">
      <div class="diff-head"><span class="diff-cat">${escHtml(r.category)}</span>
        <span class="tag-edited">編集済み</span></div>
      <div class="diff-block cur">
        <div class="diff-label">現在（営業が編集）</div>
        <div class="diff-title">${escHtml(r.current.title)}</div>
        <div class="diff-text">${escHtml(r.current.summary)}</div>
      </div>
      <div class="diff-block nw">
        <div class="diff-label">新しい抽出結果</div>
        <div class="diff-title">${escHtml(r.newTitle)}</div>
        <div class="diff-text">${escHtml(r.newSummary)}</div>
      </div>
      <div class="choice-row">
        <button class="btn btn-sm${r.keepText ? " btn-accent" : ""}" data-keep="${i}" data-v="1">残す</button>
        <button class="btn btn-sm${r.keepText ? "" : " btn-accent"}" data-keep="${i}" data-v="0">上書き</button>
      </div>
      <div class="hint-sm">数値（棚卸人数など）は選択に関わらず更新されます</div>
    </div>`;
  }).join("");

  wrap.querySelectorAll("[data-keep]").forEach(b => {
    b.addEventListener("click", () => {
      pendingExtraction.results[+b.dataset.keep].keepText = b.dataset.v === "1";
      renderExtractDiff();
    });
  });
  box.style.display = "";
}

function cancelExtraction() {
  pendingExtraction = null;
  document.getElementById("extract-review").style.display = "none";
}

async function applyExtraction() {
  if (!pendingExtraction) return;
  const { caseId, hearingIds, results } = pendingExtraction;
  setStatus("反映中…");
  if (!demoMode) await RoiCore.commitExtraction(caseId, hearingIds, results);
  cancelExtraction();
  setStatus(`${results.length}件の課題を反映しました`);
  await renderCaseOverview();
}

/* ============================================================
   ② 提案（一覧＋アコーディオン＋ステップウィザード）
   ------------------------------------------------------------
   ・課題一覧の右に損失コスト/年を表示
   ・クリックでアコーディオンが開き、試算→解決策→投資効果の3ステップ
   ・ステップ3で「提案に含める」と、上部の全体回収に反映される
   ・信頼度は「入力値の確定/推定/未確認」と「解決策の根拠区分」の両方から算出
   ============================================================ */
let propState = null; // { caseId, issues: [...] }

async function loadProposal() {
  const caseId = document.getElementById("case-id-2").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  const list = document.getElementById("proposal-list");
  list.innerHTML = `<div class="meta">読み込み中…</div>`;

  const [rows, solutions, decisions, issues] = await Promise.all([
    demoMode ? [] : RoiCore.getCalcRowsForCase(caseId, {}),
    RoiCore.getSolutions(),
    demoMode ? [] : RoiCore.getDecisions(caseId),
    demoMode ? [] : RoiCore.getIssues(caseId),
  ]);

  // ROI試算データが無くても、抽出課題だけある場合（数値が取れなかった課題）も表示する
  if (!rows.length && !issues.length) {
    list.innerHTML = `<div class="meta">この案件の課題がまだ抽出されていません。①でAI抽出を行ってください。</div>`;
    document.getElementById("totals").style.display = "none";
    return;
  }

  // ROI試算の行をカテゴリごとにまとめる
  const byCat = {};
  rows.forEach(r => {
    byCat[r.category] = byCat[r.category] || { inputs: [], outputs: [] };
    (r.kind === "出力" ? byCat[r.category].outputs : byCat[r.category].inputs).push(r);
  });

  // 課題（抽出課題）を主軸にする。カテゴリ未設定の課題も必ず一覧に載せる。
  propState = {
    caseId,
    issues: issues
      .filter(iss => iss.assignStatus !== RoiCore.ASSIGN_SKIP)
      .map(iss => {
        const cat = iss.category || "";
        const g = byCat[cat] || { inputs: [], outputs: [] };
        const plans = cat
          ? solutions.filter(s => s.category === cat)
              .sort((a, b) => ({ "低": 0, "中": 1, "高": 2 }[a.cost] ?? 9) - ({ "低": 0, "中": 1, "高": 2 }[b.cost] ?? 9))
          : [];
        const dec = cat ? decisions.find(d => d.category === cat) : null;
        let sel = 0;
        if (dec) {
          const i = plans.findIndex(p => p.name === dec.solutionName);
          if (i >= 0) sel = i;
        }
        const plan = plans[sel];
        return {
          issueId: iss.issueId,
          category: cat,
          assignStatus: iss.assignStatus,
          candidates: iss.candidates || [],
          title: iss.title || cat,
          summary: iss.summary || "",
          inputs: g.inputs,
          outputs: g.outputs,
          plans, sel, step: 1,
          cost: dec ? dec.cost : (plan ? plan.initialCost : 0),
          rate: dec ? dec.rate : (plan ? plan.rate : 0),
          basis: dec ? dec.basis : (plan ? plan.basis : ""),
          basisLevel: dec ? dec.basisLevel : (plan ? plan.basisLevel : "一般値"),
          included: dec ? dec.included : false,
        };
      }),
  };

  renderProposal();
}

/* 現状の年間損失額（出力行のうち削減額系を除いた「コスト」項目の合計）。
 * 削減額（_saving）は改善率を掛けた後の値なので、損失としては採らない。 */
function lossOfIssue(is) {
  const costRow = is.outputs.find(r => /_cost_yr$|_loss$/.test(String(r.itemId)))
    || is.outputs.find(r => !String(r.itemId).endsWith("_saving"));
  const v = costRow ? Number(costRow.value) : NaN;
  return isNaN(v) ? null : v;
}
function hoursOfIssue(is) {
  const r = is.outputs.find(x => /_hours_yr$/.test(String(x.itemId)));
  const v = r ? Number(r.value) : NaN;
  return isNaN(v) ? null : v;
}

function renderProposal() {
  const list = document.getElementById("proposal-list");
  list.innerHTML = propState.issues.map((is, ix) => {
    const loss = lossOfIssue(is);
    return `
    <div class="acc${!is.category ? " unmapped" : ""}" id="acc-${ix}">
      <div class="acc-head" data-acc="${ix}">
        <div class="a-main">
          ${catTag(is)}
          <div class="a-name">${escHtml(is.title || is.category)}</div>
          <div class="a-result" id="res-${ix}" style="display:none"></div>
        </div>
        <div class="a-loss">
          <b id="loss-${ix}">${loss != null ? fmtMan(loss) : "未算出"}</b>
          <span>損失/年</span>
        </div>
        <span id="chev-${ix}">▸</span>
      </div>
      <div class="acc-body">
        ${is.summary ? `<div class="issue-summary">${escHtml(is.summary)}</div>` : ""}
        ${unmappedBlock(is, ix)}
        ${!is.category ? "" : (!is.inputs.length ? `<div class="meta" style="margin:8px 0">この課題は数値化していません。解決案のみ提示します。</div>` : "")}
        ${!is.category ? "" : `
        <div class="wiz-dots">
          <span class="wiz-dot" id="dot-${ix}-1">1</span><div class="wiz-line" id="line-${ix}-1"></div>
          <span class="wiz-dot" id="dot-${ix}-2">2</span><div class="wiz-line" id="line-${ix}-2"></div>
          <span class="wiz-dot" id="dot-${ix}-3">3</span>
        </div>
        <div class="wiz-labels"><span>試算</span><span>解決策</span><span>投資効果</span></div>

        <div class="wiz-step" id="step-${ix}-1">
          <div class="mini-box" style="text-align:center;margin-bottom:10px">
            <div class="mini-label">3年放置した場合</div>
            <div id="v3-${ix}" style="font-size:22px;font-weight:600;color:#a32d2d">—</div>
          </div>
          ${is.inputs.map(inp => `
            <div class="sl-row">
              <label>${escHtml(inp.name)}</label>
              <input type="range" data-slider="${ix}" data-item="${escAttr(inp.itemId)}"
                min="${sliderMin(inp)}" max="${sliderMax(inp)}" step="${sliderStep(inp)}" value="${Number(inp.value) || 0}">
              <button class="cf-chip cf-${escAttr(inp.confidence || "未確認")}"
                data-cf="${ix}" data-item="${escAttr(inp.itemId)}"
                id="chip-${ix}-${escAttr(inp.itemId)}"></button>
            </div>`).join("")}
          <div class="hint-sm">数値をタップすると 未確認 → 推定 → 確定 が切り替わります</div>
          <div style="border-top:1px solid #e5e3da;margin-top:8px;padding-top:8px;display:flex;justify-content:space-between;font-size:11px">
            <span style="color:#5f5e5a">この課題の確からしさ</span><span id="cf-${ix}" style="font-weight:600">—</span>
          </div>
        </div>

        <div class="wiz-step" id="step-${ix}-2">
          <div id="plans-${ix}"></div>
          <div class="mini-box" style="margin-top:8px">
            <div class="mini-label">現在</div>
            <div style="height:13px;background:#e24b4a;border-radius:4px;margin-bottom:6px"></div>
            <div class="mini-label">導入後</div>
            <div style="background:#fff;border-radius:4px">
              <div id="after-${ix}" style="height:13px;background:#1d9e75;border-radius:4px;width:50%"></div>
            </div>
            <div id="after-t-${ix}" style="font-size:11px;color:#5f5e5a;margin-top:6px">—</div>
          </div>
        </div>

        <div class="wiz-step" id="step-${ix}-3">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
            <div><div class="mini-label">導入費（編集可）</div>
              <input type="number" data-cost="${ix}" id="cost-${ix}" value="${is.cost}" style="font-size:13px;padding:5px 7px"></div>
            <div><div class="mini-label">回収期間</div><div id="pay-${ix}" style="font-size:18px;font-weight:600">—</div></div>
          </div>
          <svg id="chart-${ix}" viewBox="0 0 300 100" style="width:100%;height:auto"></svg>
          <div style="border-top:1px solid #e5e3da;margin-top:8px;padding-top:8px;text-align:center">
            <div id="hrs-${ix}" style="font-size:15px;font-weight:600">—</div>
            <div style="font-size:10px;color:#5f5e5a">この業務の年間工数</div>
          </div>
        </div>

        <div class="btn-row">
          <button class="btn" data-back="${ix}" id="back-${ix}">戻る</button>
          <button class="btn btn-accent" data-next="${ix}" id="next-${ix}" style="flex:2">次へ</button>
        </div>`}
      </div>
    </div>`;
  }).join("");

  document.getElementById("totals").style.display = "";
  propState.issues.forEach((is, ix) => { if (is.category) { renderPlans(ix); navWizard(ix); } });
  updateProposal();
}

function sliderMin(inp) { return /時給|単価|額|費/.test(inp.name) ? 0 : 1; }
function sliderMax(inp) {
  const v = Number(inp.value) || 1;
  if (/時給|単価/.test(inp.name)) return 8000;
  if (/額|費|損失/.test(inp.name)) return Math.max(v * 3, 10000000);
  return Math.max(v * 3, 24);
}
function sliderStep(inp) {
  if (/時給|単価/.test(inp.name)) return 100;
  if (/額|費|損失/.test(inp.name)) return 100000;
  return 1;
}

function renderPlans(ix) {
  const is = propState.issues[ix];
  const el = document.getElementById(`plans-${ix}`);
  if (!el) return;
  if (!is.plans.length) { el.innerHTML = `<div class="meta">解決案が未登録です（ソリューションDBに追加してください）</div>`; return; }
  el.innerHTML = is.plans.map((p, i) => {
    const on = i === is.sel;
    return `
    <div class="plan-card${on ? " on" : ""}" data-plan="${ix}" data-pi="${i}">
      <div class="plan-head">
        <span class="plan-name">${escHtml(p.name)}</span>
        <span class="plan-tier cost-${escAttr(p.cost)}">${escHtml(p.cost)}</span>
      </div>
      <div class="plan-method">${escHtml(p.method)}</div>
      <div class="plan-rate">
        <span>改善率</span>
        ${on
          ? `<input type="number" data-rate="${ix}" id="rate-${ix}" value="${is.rate}" style="width:56px;padding:4px 6px;font-size:13px">%`
          : `<b>${p.rate}%</b>`}
        <span class="basis-tag basis-${escAttr(on ? is.basisLevel : p.basisLevel)}"
          ${on ? `data-basis="${ix}" style="cursor:pointer"` : ""}>${escHtml(on ? is.basisLevel : p.basisLevel)}</span>
      </div>
      ${on
        ? `<textarea data-basistext="${ix}" id="basis-${ix}" rows="3" style="font-size:11px;line-height:1.6">${escHtml(is.basis)}</textarea>
           <div class="hint-sm">根拠区分をタップすると 実績 → 推定 → 一般値 が切り替わります。変更はこの案件にのみ保存されます</div>`
        : `<div class="basis-text">${escHtml(p.basis)}</div>`}
    </div>`;
  }).join("");
}

function navWizard(ix) {
  const is = propState.issues[ix];
  [1, 2, 3].forEach(i => {
    const s = document.getElementById(`step-${ix}-${i}`);
    if (s) s.classList.toggle("on", i === is.step);
    const d = document.getElementById(`dot-${ix}-${i}`);
    if (d) d.classList.toggle("on", i <= is.step);
  });
  [1, 2].forEach(i => {
    const l = document.getElementById(`line-${ix}-${i}`);
    if (l) l.classList.toggle("on", is.step > i);
  });
  const back = document.getElementById(`back-${ix}`);
  if (back) back.style.visibility = is.step === 1 ? "hidden" : "visible";
  const next = document.getElementById(`next-${ix}`);
  if (next) next.textContent = is.step === 3 ? (is.included ? "提案から外す" : "提案に含める") : "次へ";
}

function updateProposal() {
  let sumLoss = 0, cfAll = [], payCost = 0, paySave = 0, payCf = [];

  propState.issues.forEach((is, ix) => {
    const loss = lossOfIssue(is), hrs = hoursOfIssue(is);
    const save = loss != null ? loss * (is.rate / 100) : 0;

    if (loss != null) { sumLoss += loss; }
    // 未設定の課題は試算していないので、信頼度の平均には含めない
    if (is.category) cfAll.push(RoiCore.confidenceOf(is.inputs, is.basisLevel));

    const lossEl = document.getElementById(`loss-${ix}`);
    if (lossEl) lossEl.textContent = loss != null ? fmtMan(loss) : "未算出";
    const v3 = document.getElementById(`v3-${ix}`);
    if (v3) v3.textContent = loss != null ? fmtMan(loss * 3) : "—";

    is.inputs.forEach(inp => {
      const chip = document.getElementById(`chip-${ix}-${inp.itemId}`);
      if (!chip) return;
      chip.className = `cf-chip cf-${inp.confidence || "未確認"}`;
      chip.textContent = fmtNum(inp.value) + (inp.unit || "");
    });

    const cfIssue = RoiCore.confidenceOf(is.inputs, is.basisLevel);
    const cfEl = document.getElementById(`cf-${ix}`);
    if (cfEl) cfEl.textContent = `${Math.round(cfIssue * 100)}% ・${RoiCore.confidenceLabel(cfIssue)}`;

    const after = document.getElementById(`after-${ix}`);
    if (after && loss != null && loss > 0) {
      after.style.width = Math.max((loss - save) / loss * 100, 2) + "%";
      document.getElementById(`after-t-${ix}`).textContent =
        `${fmtMan(loss)} → ${fmtMan(loss - save)}（年 ${fmtMan(save)}の削減）`;
    }

    const mo = save / 12;
    const payEl = document.getElementById(`pay-${ix}`);
    if (payEl) payEl.textContent = mo > 0 ? `${Math.ceil(is.cost / mo)}ヶ月` : "—";
    const hrsEl = document.getElementById(`hrs-${ix}`);
    if (hrsEl) hrsEl.textContent = hrs != null ? `${fmtNum(hrs)}時間 ＝ ${Math.round(hrs / 8)}人日` : "—";
    drawChart(ix, is.cost, mo);

    const res = document.getElementById(`res-${ix}`);
    if (res) {
      if (is.included) {
        res.style.display = "";
        res.textContent = `決定：${is.plans[is.sel] ? is.plans[is.sel].name : ""} ／ ${fmtMan(is.cost)}`;
      } else res.style.display = "none";
    }

    if (is.included) { payCost += is.cost; paySave += save; payCf.push(cfIssue); }
    if (is.category) navWizard(ix);
  });

  document.getElementById("t-loss").textContent = fmtMan(sumLoss);
  const total = propState.issues.length;
  const calced = propState.issues.filter(i => lossOfIssue(i) != null).length;
  const gc = cfAll.length ? cfAll.reduce((a, b) => a + b, 0) / cfAll.length : 0;
  // 全課題のうち何件が試算できているかを明示する。
  // 一部しか金額が出ていない状態で合計だけ見せると全体像と誤解されるため。
  document.getElementById("t-loss-cf").textContent = calced < total
    ? `全${total}件中 ${calced}件のみ試算済`
    : `${Math.round(gc * 100)}% ・${RoiCore.confidenceLabel(gc)}`;

  if (paySave > 0) {
    document.getElementById("t-pay").textContent = `${Math.ceil(payCost / (paySave / 12))}ヶ月`;
    const pc = payCf.reduce((a, b) => a + b, 0) / payCf.length;
    document.getElementById("t-pay-cf").textContent = `${Math.round(pc * 100)}% ・${RoiCore.confidenceLabel(pc)}`;
  } else {
    document.getElementById("t-pay").textContent = "—";
    document.getElementById("t-pay-cf").textContent = "課題を決定すると表示";
  }
}

function drawChart(ix, cost, mo) {
  const el = document.getElementById(`chart-${ix}`);
  if (!el) return;
  const x0 = 28, x1 = 292, y0 = 10, y1 = 84;
  const M = Math.min(Math.ceil((mo > 0 ? cost / mo : 12) * 1.6), 60) || 12;
  const mx = Math.max(mo * M - cost, cost) * 1.1;
  const px = m => x0 + (x1 - x0) * m / M, py = v => y1 - (y1 - y0) * ((v + cost) / (mx + cost));
  let pts = [];
  for (let m = 0; m <= M; m++) pts.push(px(m) + "," + py(mo * m - cost));
  const zx = mo > 0 ? px(Math.min(cost / mo, M)) : x1;
  el.innerHTML =
    `<line x1="${x0}" y1="${py(0)}" x2="${x1}" y2="${py(0)}" stroke="#b4b2a9"/>` +
    `<polyline points="${pts.join(" ")}" fill="none" stroke="#185fa5" stroke-width="2"/>` +
    `<line x1="${zx}" y1="${y0}" x2="${zx}" y2="${y1}" stroke="#378add" stroke-dasharray="3 3"/>` +
    `<circle cx="${zx}" cy="${py(0)}" r="4" fill="#185fa5"/>` +
    `<text x="${Math.min(zx + 5, 210)}" y="${y0 + 10}" font-size="10" fill="#5f5e5a">${mo > 0 ? Math.ceil(cost / mo) + "ヶ月" : "—"}</text>`;
}

/* ---- ②タブのイベント（委譲） ---- */
document.addEventListener("input", async (e) => {
  if (!propState) return;
  const t = e.target;
  if (t.dataset.slider !== undefined) {
    const is = propState.issues[+t.dataset.slider];
    const inp = is.inputs.find(i => i.itemId === t.dataset.item);
    if (inp) {
      inp.value = Number(t.value);
      recalcOutputs(is);
      updateProposal();
      if (!demoMode) await RoiCore.updateCalcInput(propState.caseId, inp.itemId, { value: inp.value });
    }
  }
  if (t.dataset.cost !== undefined) { propState.issues[+t.dataset.cost].cost = Number(t.value) || 0; updateProposal(); }
  if (t.dataset.rate !== undefined) { propState.issues[+t.dataset.rate].rate = Number(t.value) || 0; updateProposal(); }
  if (t.dataset.basistext !== undefined) { propState.issues[+t.dataset.basistext].basis = t.value; }
});

document.addEventListener("click", async (e) => {
  if (!propState) return;

  const asg = e.target.closest("[data-assign]");
  if (asg) { openAssignModal(+asg.dataset.assign); return; }

  const nc = e.target.closest("[data-nocat]");
  if (nc) {
    const is = propState.issues[+nc.dataset.nocat];
    if (!demoMode) await RoiCore.assignIssueCategory(propState.caseId, is.issueId,
      { category: "", assignStatus: RoiCore.ASSIGN_NOCAT });
    await loadProposal();
    return;
  }

  const chip = e.target.closest("[data-cf]");
  if (chip) {
    const is = propState.issues[+chip.dataset.cf];
    const inp = is.inputs.find(i => i.itemId === chip.dataset.item);
    if (inp) {
      const cyc = ["未確認", "推定", "確定"];
      inp.confidence = cyc[(cyc.indexOf(inp.confidence || "未確認") + 1) % 3];
      updateProposal();
      if (!demoMode) await RoiCore.updateCalcInput(propState.caseId, inp.itemId, { confidence: inp.confidence });
    }
    return;
  }

  const bt = e.target.closest("[data-basis]");
  if (bt) {
    const is = propState.issues[+bt.dataset.basis];
    const cyc = RoiCore.BASIS_LEVELS;
    is.basisLevel = cyc[(cyc.indexOf(is.basisLevel) + 1) % cyc.length];
    renderPlans(+bt.dataset.basis);
    updateProposal();
    return;
  }

  const pc = e.target.closest("[data-plan]");
  if (pc) {
    const ix = +pc.dataset.plan, is = propState.issues[ix];
    is.sel = +pc.dataset.pi;
    const p = is.plans[is.sel];
    is.cost = p.initialCost; is.rate = p.rate; is.basis = p.basis; is.basisLevel = p.basisLevel;
    const ci = document.getElementById(`cost-${ix}`);
    if (ci) ci.value = is.cost;
    renderPlans(ix);
    updateProposal();
    return;
  }

  const nx = e.target.closest("[data-next]");
  if (nx) {
    const ix = +nx.dataset.next, is = propState.issues[ix];
    if (is.step < 3) { is.step++; navWizard(ix); }
    else {
      is.included = !is.included;
      if (!demoMode) {
        await RoiCore.saveDecision(propState.caseId, is.category, {
          solutionName: is.plans[is.sel] ? is.plans[is.sel].name : "",
          cost: is.cost, rate: is.rate, basis: is.basis, basisLevel: is.basisLevel,
          included: is.included,
        });
        await RoiCore.toggleSelection(propState.caseId, is.category, is.included);
      }
      updateProposal();
    }
    return;
  }

  const bk = e.target.closest("[data-back]");
  if (bk) { const is = propState.issues[+bk.dataset.back]; if (is.step > 1) { is.step--; navWizard(+bk.dataset.back); } return; }

  const ah = e.target.closest("[data-acc]");
  if (ah) {
    const ix = +ah.dataset.acc, el = document.getElementById(`acc-${ix}`);
    el.classList.toggle("open");
    document.getElementById(`chev-${ix}`).textContent = el.classList.contains("open") ? "▾" : "▸";
  }
});

/* 出力行はExcelの数式で計算されるが、スライダー操作中は即座に画面へ返せないため、
 * ROIマスタの数式をJS側でも評価して暫定表示する。Excel側の値は書き込み後に
 * 数式が再計算するので、次回読み込み時には一致する。 */
function recalcOutputs(is) {
  const vals = {};
  is.inputs.forEach(i => { vals[i.itemId] = Number(i.value) || 0; });
  const master = RoiCore.getMasterItems().filter(m => m.category === is.category && m.kind === "出力");
  // 依存関係があるため数回まわして収束させる
  for (let pass = 0; pass < 3; pass++) {
    master.forEach(m => {
      if (!m.formula) return;
      try {
        const expr = String(m.formula).replace(/[a-zA-Z_][a-zA-Z0-9_]*/g, tok =>
          vals[tok] !== undefined ? vals[tok] : "0");
        vals[m.itemId] = Function(`"use strict";return (${expr})`)();
      } catch (err) { /* 数式が壊れている場合は無視 */ }
    });
  }
  is.outputs.forEach(o => { if (vals[o.itemId] !== undefined) o.value = vals[o.itemId]; });
}

/* カテゴリのタグ表示。未設定・カテゴリなし提案を区別する。 */
function catTag(is) {
  if (is.category) return `<div class="a-cat">${escHtml(is.category)}</div>`;
  if (is.assignStatus === RoiCore.ASSIGN_NOCAT) {
    return `<span class="tag-nocat">カテゴリなしで提案</span>`;
  }
  const sug = (is.candidates || [])[0];
  return `<span class="tag-unmapped">カテゴリ未設定</span>`
    + (sug ? ` <span style="font-size:10px;color:#888780">AI提案：${escHtml(sug.name)}</span>` : "");
}

/* カテゴリ未設定の課題に出す操作ブロック。 */
function unmappedBlock(is, ix) {
  if (is.category) return "";
  if (is.assignStatus === RoiCore.ASSIGN_NOCAT) {
    return `<div class="meta" style="margin:8px 0">カテゴリなしで提案に含めます。金額は出ませんが提案書には載ります。</div>
      <div class="btn-row"><button class="btn" data-assign="${ix}">カテゴリを割り当てる</button></div>`;
  }
  return `
    <div class="btn-row">
      <button class="btn btn-accent" data-assign="${ix}" style="flex:2">カテゴリを割り当てる</button>
      <button class="btn" data-nocat="${ix}">このまま提案に含める</button>
    </div>`;
}

/* ============================================================
   カテゴリ割り当てモーダル（候補選択 ＋ AIとの壁打ち）
   ============================================================ */
let assignState = null;

function openAssignModal(ix) {
  const is = propState.issues[ix];
  assignState = { ix, issue: is, selected: 0, messages: [], definition: null };
  document.getElementById("assign-issue").innerHTML = `
    <div class="mini-label">対象の課題</div>
    <div style="font-size:12px;font-weight:600">${escHtml(is.title)}</div>
    <div style="font-size:11px;color:#5f5e5a;line-height:1.6;margin-top:3px">${escHtml(is.summary)}</div>`;
  document.getElementById("assign-pick").style.display = "";
  document.getElementById("assign-chat").style.display = "none";
  renderCandidates();
  document.getElementById("assign-modal").style.display = "flex";
}

function renderCandidates() {
  const cands = assignState.issue.candidates || [];
  const el = document.getElementById("assign-candidates");
  if (!cands.length) {
    el.innerHTML = `<div class="meta">AIからの候補がありません。「AIと相談」でカテゴリを作れます。</div>`;
    document.getElementById("assign-apply-btn").style.display = "none";
    return;
  }
  document.getElementById("assign-apply-btn").style.display = "";
  el.innerHTML = cands.map((c, i) => `
    <div class="cand-card${i === assignState.selected ? " on" : ""}" data-cand="${i}">
      <span class="radio-dot"></span>
      <div class="c-body">
        <div class="cand-head">
          <span class="cand-name">${escHtml(c.name)}</span>
          <span class="${c.isNew ? "tag-new2" : "tag-exist"}">${c.isNew ? "新規" : "既存"}</span>
        </div>
        <div class="cand-reason">${escHtml(c.reason || "")}</div>
      </div>
    </div>`).join("");
  el.querySelectorAll("[data-cand]").forEach(e => {
    e.addEventListener("click", () => { assignState.selected = +e.dataset.cand; renderCandidates(); });
  });
}

async function applyCandidate() {
  const c = (assignState.issue.candidates || [])[assignState.selected];
  if (!c) return;
  if (c.isNew) {
    // 新規候補は定義が必要なので、そのまま壁打ちに入る
    await startCategoryChat(`「${c.name}」で作りたいです。`);
    return;
  }
  await finishAssign(c.name, RoiCore.ASSIGN_DONE);
}

async function finishAssign(category, status) {
  const { issue } = assignState;
  if (!demoMode) await RoiCore.assignIssueCategory(propState.caseId, issue.issueId, { category, assignStatus: status });
  closeAssignModal();
  await loadProposal();
}

function closeAssignModal() {
  assignState = null;
  document.getElementById("assign-modal").style.display = "none";
}

/* ---- AIとの壁打ち ---- */
async function startCategoryChat(initialMessage) {
  document.getElementById("assign-pick").style.display = "none";
  document.getElementById("assign-chat").style.display = "";
  if (initialMessage) assignState.messages.push({ role: "user", content: initialMessage });
  renderChat();
  await sendCategoryChat();
}

function renderChat() {
  document.getElementById("chat-log").innerHTML = assignState.messages.map(m =>
    `<div class="${m.role === "user" ? "chat-me" : "chat-ai"}">${escHtml(m.content)}</div>`).join("");
  const d = assignState.definition;
  const box = document.getElementById("chat-def");
  if (!d) { box.innerHTML = `<div class="def-label">案を作成中…</div>`; return; }
  box.innerHTML = `
    <div class="def-label">現在の案（編集できます）</div>
    <div class="mini-label">カテゴリ名</div>
    <input id="def-name" value="${escAttr(d.category || "")}" style="font-size:12px;margin-bottom:8px">
    <div class="mini-label">入力項目</div>
    ${(d.inputs || []).map(i => `
      <div class="def-item"><span class="di-name">${escHtml(i.name)}</span><span class="di-unit">${escHtml(i.unit || "")}</span></div>`).join("")}
    <div class="mini-label" style="margin-top:6px">計算式</div>
    <div class="def-formula">${(d.outputs || []).map(o => `${escHtml(o.name)} = ${escHtml(o.formula || "")}`).join("<br>")}</div>`;
}

async function sendCategoryChat() {
  const box = document.getElementById("chat-def");
  box.innerHTML = `<div class="def-label">AIが考えています…</div>`;
  try {
    const data = await RoiCore.proposeCategoryDefinition(
      propState.caseId, assignState.issue, assignState.messages);
    if (data.reply) assignState.messages.push({ role: "assistant", content: data.reply });
    if (data.definition) assignState.definition = data.definition;
    renderChat();
  } catch (e) {
    console.warn(e);
    box.innerHTML = `<div class="def-label">エラー: ${escHtml(e.message || "")}</div>`;
  }
}

async function saveCategory(scope) {
  const d = assignState.definition;
  if (!d) return;
  const nameInput = document.getElementById("def-name");
  if (nameInput) d.category = nameInput.value.trim() || d.category;
  if (!demoMode) {
    await RoiCore.addCategoryToMaster(d, {
      scope, caseId: scope === RoiCore.SCOPE_TEMP ? propState.caseId : "",
    });
  }
  await finishAssign(d.category, RoiCore.ASSIGN_DONE);
}

function fmtMan(v) {
  const n = Number(v);
  if (isNaN(n)) return "—";
  return Math.round(n / 10000).toLocaleString("ja-JP") + "万円";
}

/* ============================================================
   ③ プロンプト出力
   ============================================================ */
async function buildPrompt() {
  const caseId = document.getElementById("case-id-3").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  const rows = demoMode ? [] : await RoiCore.getCalcRowsForCase(caseId, { onlySelected: true });
  const allIssues = demoMode ? [] : await RoiCore.getIssues(caseId);
  // カテゴリなしで提案に含めると決めた課題も、金額なしで提案書に載せる
  const noCatIssues = allIssues.filter(i => i.assignStatus === RoiCore.ASSIGN_NOCAT);
  if (!rows.length && !noCatIssues.length) {
    document.getElementById("prompt-output").innerHTML = `<div class="meta">②で提案書に含める課題を選択してください。</div>`;
    return;
  }
  const solutions = demoMode ? [] : await RoiCore.getSolutions();
  const custRow = demoMode ? null : await RoiCore.getCustomerInfo(caseId);

  const byCategory = {};
  rows.forEach(r => {
    byCategory[r.category] = byCategory[r.category] || { inputs: [], outputs: [] };
    (r.kind === "出力" ? byCategory[r.category].outputs : byCategory[r.category].inputs).push(r);
  });
  const cats = Object.keys(byCategory);

  const custInputText = custRow ? `取引先：${custRow.name || ""}\n窓口：${custRow.contact || ""}` : `案件ID：${caseId}`;

  const roiText = cats.map(cat => {
    const lines = byCategory[cat].outputs.map(r => `${r.name} ${fmtNum(r.value)}${r.unit || ""}`).join("\n");
    return `[${cat}]\n${lines}`;
  }).join("\n\n");

  const decisions = demoMode ? [] : await RoiCore.getDecisions(caseId);

  const outlineText = cats.map((cat, i) => {
    const dec = decisions.find(d => d.category === cat);
    const cands = solutions.filter(s => s.category === cat);
    if (dec) {
      return `${i + 1}. ${cat}\n採用：${dec.solutionName}\n改善率：${dec.rate}%（根拠区分：${dec.basisLevel}）\n根拠：${dec.basis}\n概算導入費：${fmtNum(dec.cost)}円`;
    }
    const solLines = cands.length
      ? cands.map(s => `・${s.name}（コスト感：${s.cost}／改善率${s.rate}%・${s.basisLevel}）${s.method}`).join("\n")
      : "（解決案未登録）";
    return `${i + 1}. ${cat}\n解決案候補：\n${solLines}`;
  }).join("\n\n")
  + (noCatIssues.length
      ? "\n\n" + noCatIssues.map((iss, j) =>
          `${cats.length + j + 1}. ${iss.title}\n（ROI試算なし・定性課題として記載）\n${iss.summary}`).join("\n\n")
      : "");

  const promptText =
`中小企業向けの提案書を、現状課題→解決の型→ROI→費用・体制の4章構成で作成してください。
トーンは平易で、数値の根拠（信頼度）を明示してください。`;

  const blocks = [
    { label: "生成プロンプト", body: promptText },
    { label: "顧客インプット", body: custInputText },
    { label: "ROI数値", body: roiText },
    { label: "章立て骨子", body: outlineText },
  ];

  const out = document.getElementById("prompt-output");
  out.innerHTML = blocks.map((b, i) => `
    <div class="prompt-block" data-idx="${i}">
      <div class="pb-head"><span class="pb-label">${escHtml(b.label)}</span><button class="btn pb-copy">コピー</button></div>
      <div class="pb-body">${escHtml(b.body)}</div>
    </div>`).join("");
  out.querySelectorAll(".pb-copy").forEach((btn, i) => btn.addEventListener("click", () => copyToClipboard(blocks[i].body)));
}

function copyToClipboard(text) { if (navigator.clipboard) navigator.clipboard.writeText(text); }

/* ---------- ユーティリティ ---------- */
/* ---------- ユーティリティ ---------- */
/* Excelがセルを日付として自動認識すると、"YYYY-MM-DD HH:MM"のような文字列で
 * 書き込んでも、読み戻すとシリアル値（数値）で返ってくることがある。
 * どちらの形でも "MM-DD" 表示に変換できるようにする。 */
function fmtMonthDay(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") {
    // Excelのシリアル値（1900年1月1日を1とする日数）をJS Dateに変換
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (isNaN(d.getTime())) return "";
    const p = n => String(n).padStart(2, "0");
    return `${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }
  const s = String(v);
  return s.length >= 10 ? s.slice(5, 10) : s;
}
function setStatus(msg) { document.getElementById("extract-status").textContent = msg; }
function fmtNum(v) { const n = Number(v); return isNaN(n) ? String(v) : n.toLocaleString("ja-JP"); }
function escHtml(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function escAttr(s) { return escHtml(s); }
