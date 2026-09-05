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
let currentReview = null;

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
  renderCategorySelect();
  await renderCaseIdList();
}

function bindStaticUI() {
  document.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", () => switchTab(btn.dataset.tab)));
  document.getElementById("reload-btn").addEventListener("click", () => location.reload());
  document.getElementById("settings-btn").addEventListener("click", openSettings);
  document.getElementById("cfg-close-btn").addEventListener("click", closeSettings);
  document.getElementById("cfg-save-btn").addEventListener("click", saveSettings);

  document.getElementById("save-log-btn").addEventListener("click", saveHearingLog);
  document.getElementById("extract-btn").addEventListener("click", runExtraction);
  document.getElementById("apply-extract-btn").addEventListener("click", applyExtraction);
  document.getElementById("cancel-extract-btn").addEventListener("click", cancelExtraction);
  document.getElementById("case-id").addEventListener("change", loadHearingLogForCase);

  document.getElementById("load-proposal-btn").addEventListener("click", loadProposal);
  document.getElementById("build-prompt-btn").addEventListener("click", buildPrompt);
}

function switchTab(tab) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".pane").forEach(p => p.classList.remove("active"));
  document.getElementById("pane-" + tab).classList.add("active");
  const cur = document.getElementById("case-id").value;
  ["case-id-2", "case-id-3"].forEach(id => { document.getElementById(id).value = cur; });
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
async function renderCaseIdList() {
  const ids = demoMode ? ["KM-01", "OF-02"] : await RoiCore.listCaseIds();
  document.getElementById("case-id-list").innerHTML = ids.map(id => `<option value="${escAttr(id)}">`).join("");
}
function renderCategorySelect() {
  const cats = RoiCore.getCategories();
  document.getElementById("category-select").innerHTML = cats.map(c => `<option value="${escAttr(c)}">${escHtml(c)}</option>`).join("");
}

/* ============================================================
   ① 議事録入力
   ============================================================ */
async function saveHearingLog() {
  const caseId = document.getElementById("case-id").value.trim();
  const category = document.getElementById("category-select").value;
  const speaker = document.getElementById("speaker-select").value;
  const text = document.getElementById("hearing-text").value.trim();
  const url = document.getElementById("hearing-url").value.trim();
  if (!caseId) { setStatus("案件IDを入力してください"); return; }
  if (!text && !url) { setStatus("発言・メモかURLのどちらかを入力してください"); return; }

  if (demoMode) { setStatus("デモモードのため保存はシミュレーションのみです"); }
  else {
    await RoiCore.appendHearingLog(caseId, category, speaker, { text, url });
    setStatus("議事録に記録しました");
  }
  document.getElementById("hearing-text").value = "";
  document.getElementById("hearing-url").value = "";
  await loadHearingLogForCase();
}

async function loadHearingLogForCase() {
  const caseId = document.getElementById("case-id").value.trim();
  const list = document.getElementById("hearing-log-list");
  if (!caseId) { list.innerHTML = ""; return; }
  const rows = demoMode ? [] : await RoiCore.listHearingLogs(caseId);
  list.innerHTML = rows.map(r => `
    <div class="log-item">
      <div class="meta">${escHtml(r.category || "")} ／ ${escHtml(r.speaker || "")} ／ ${escHtml(r.registeredAt || "")}</div>
      ${r.url ? `<div class="meta"><a href="${escAttr(r.url)}" target="_blank" rel="noopener">${escHtml(r.url)}</a></div>` : ""}
      <div class="text">${escHtml(r.text || "")}</div>
    </div>`).join("") || `<div class="meta">まだ記録がありません</div>`;
}

/* ---------- AI抽出（レビューあり） ---------- */
async function runExtraction() {
  const caseId = document.getElementById("case-id").value.trim();
  const category = document.getElementById("category-select").value;
  const text = document.getElementById("hearing-text").value.trim();
  const url = document.getElementById("hearing-url").value.trim();
  if (!caseId) { setStatus("案件IDを入力してください"); return; }
  if (!text && !url) { setStatus("発言・メモかURLのどちらかを入力してください（記録前でも抽出だけ試せます）"); return; }

  const cfg = RoiCore.getConfig();
  if (!cfg.webhookUrl) { setStatus("設定（⚙）でAI連携エンドポイントを登録してください"); return; }

  setStatus("AIに抽出を依頼中…");
  try {
    const items = await RoiCore.runExtractionForReview(caseId, category, { text, url });
    currentReview = { caseId, category, items };
    renderExtractReview();
    setStatus("");
  } catch (e) {
    console.warn(e);
    setStatus(e.message || "抽出に失敗しました。エンドポイントの設定を確認してください");
  }
}

function renderExtractReview() {
  const box = document.getElementById("extract-review");
  const wrap = document.getElementById("extract-items");
  if (!currentReview || !currentReview.items.length) { box.style.display = "none"; return; }
  const master = RoiCore.getMasterItems().filter(m => m.category === currentReview.category && m.kind === "入力");
  wrap.innerHTML = currentReview.items.map((it, idx) => {
    const def = master.find(m => m.itemId === it.itemId);
    return `
    <div class="review-item" data-idx="${idx}">
      <div class="item-head"><span>${escHtml(def ? def.name : it.itemId)}</span><span>${escHtml(def ? def.unit : "")}</span></div>
      <div class="item-row">
        <input type="text" class="rv-value" value="${escAttr(it.value ?? "")}">
        <select class="rv-conf">${RoiCore.CONF_LEVELS.map(c => `<option value="${c}" ${c === it.confidence ? "selected" : ""}>${c}</option>`).join("")}</select>
      </div>
    </div>`;
  }).join("");
  box.style.display = "";
}

function cancelExtraction() { currentReview = null; document.getElementById("extract-review").style.display = "none"; }

async function applyExtraction() {
  if (!currentReview) return;
  document.querySelectorAll("#extract-items .review-item").forEach(el => {
    const idx = Number(el.dataset.idx);
    currentReview.items[idx].value = el.querySelector(".rv-value").value;
    currentReview.items[idx].confidence = el.querySelector(".rv-conf").value;
  });
  if (!demoMode) await RoiCore.applyReviewedItems(currentReview.caseId, currentReview.category, currentReview.items);
  setStatus("ROI試算シートに反映しました");
  cancelExtraction();
}

/* ============================================================
   ② 提案（課題×解決案）
   ============================================================ */
async function loadProposal() {
  const caseId = document.getElementById("case-id-2").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  const summary = demoMode ? [] : await RoiCore.getProposalSummaryForCase(caseId);
  const list = document.getElementById("proposal-list");
  if (!summary.length) {
    list.innerHTML = `<div class="meta">この案件のROI試算データがまだありません。①でAI抽出を行ってください。</div>`;
    document.getElementById("proposal-summary").style.display = "none";
    return;
  }

  list.innerHTML = summary.map(it => {
    const badgeClass = it.confidence === "確定" ? "badge-confirmed" : it.confidence === "推定" ? "badge-estimated" : "badge-unknown";
    return `
    <div class="proposal-card" data-cat="${escAttr(it.category)}">
      <div class="p-head">
        <input type="checkbox" class="p-check" ${it.selected ? "checked" : ""}>
        <div><div class="p-title">${escHtml(it.category)}</div></div>
      </div>
      ${it.solutionName ? `<div class="p-solution">→ 解決策：${escHtml(it.solutionName)}</div>` : `<div class="p-solution">解決策が未登録です（ソリューションDBに追加してください）</div>`}
      <div class="p-metrics">
        <span>削減額 <b>${it.saving != null ? fmtNum(it.saving) + (it.unit || "") + "/年" : "—"}</b></span>
        <span class="badge ${badgeClass}">${it.confidence}</span>
      </div>
    </div>`;
  }).join("");

  list.querySelectorAll(".p-check").forEach(cb => {
    cb.addEventListener("change", async (e) => {
      const cat = e.target.closest(".proposal-card").dataset.cat;
      if (!demoMode) await RoiCore.toggleSelection(caseId, cat, e.target.checked);
    });
  });

  document.getElementById("proposal-summary").style.display = "";
  const selectedCount = list.querySelectorAll(".p-check:checked").length;
  document.getElementById("proposal-summary-text").textContent = `選択中 ${selectedCount}件`;
}

/* ============================================================
   ③ プロンプト出力
   ============================================================ */
async function buildPrompt() {
  const caseId = document.getElementById("case-id-3").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  const rows = demoMode ? [] : await RoiCore.getCalcRowsForCase(caseId, { onlySelected: true });
  if (!rows.length) {
    document.getElementById("prompt-output").innerHTML = `<div class="meta">②で提案書に含める課題を選択（チェック）してください。</div>`;
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

  const outlineText = cats.map((cat, i) => {
    const sol = solutions.find(s => s.category === cat);
    return `${i + 1}. ${cat}\n解決の型：${sol ? sol.name : "（未設定）"}\n${sol ? sol.desc : ""}`;
  }).join("\n\n");

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
function setStatus(msg) { document.getElementById("extract-status").textContent = msg; }
function fmtNum(v) { const n = Number(v); return isNaN(n) ? String(v) : n.toLocaleString("ja-JP"); }
function escHtml(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function escAttr(s) { return escHtml(s); }
