/* ============================================================
 * eigyo-roi-panel.js（アコーディオン版）
 * ------------------------------------------------------------
 * 「営業報告」アドインの編集画面に埋め込む簡易パネル。
 *
 * 操作はこれだけ:
 *   ・「提案」ボタン　　　　→ 作成済みの課題をアコーディオンで一覧表示
 *   ・各行クリック　　　　　→ 詳細（ROI内訳・解決策・元データリンク）を展開
 *   ・チェックボックス　　　→ 提案書に含める課題を複数選択（即時保存）
 *   ・「作成」アイコン　　　→ 議事録＋メモをAIに渡し、課題を自動抽出して
 *                          一覧に追加（カテゴリの指定は不要。AIが判定する）
 *   ・「提案書を作成」　　　→ 選択済み課題の元データリンク一覧を表示
 *                          （実際のプロンプト組み立て・pptx化は提案ナレッジ側）
 *
 * 組み込み方法:
 *   <script src=".../roi-core.js"></script>
 *   <script src=".../eigyo-roi-panel.js"></script>
 *   <div id="roi-panel-mount"></div>
 *   RoiPanel.mount(el, caseId, { getMemo: () => 現在の備考欄の値 })
 *   getMemo を渡さない場合はパネル内に簡易メモ欄を表示する。
 * ============================================================ */

(function (global) {
  let inited = false;

  async function ensureInit() {
    if (inited) return;
    await RoiCore.ensureAllSheets();
    injectStyles();
    inited = true;
  }

  function injectStyles() {
    if (document.getElementById("roi-panel-style")) return;
    const style = document.createElement("style");
    style.id = "roi-panel-style";
    style.textContent = `
      .roi-panel { border: 1px solid #d3d1c7; border-radius: 8px; padding: 10px 12px; margin-top: 10px; font-size: 12px; background: #fafaf8; }
      .roi-panel .roi-h-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
      .roi-panel .roi-h { font-size: 12px; font-weight: 600; color: #44546A; }
      .roi-panel .roi-icon-btn {
        width: 26px; height: 26px; border-radius: 6px; border: 1px solid #b4b2a9; background: #fff;
        cursor: pointer; display: flex; align-items: center; justify-content: center; font-size: 14px;
      }
      .roi-panel .roi-icon-btn.accent { background: #185fa5; border-color: #185fa5; color: #fff; }
      .roi-panel textarea { width: 100%; font-size: 12px; padding: 5px 7px; border: 1px solid #d3d1c7; border-radius: 6px; margin-bottom: 6px; box-sizing: border-box; }
      .roi-panel button.roi-btn { font-size: 11px; padding: 6px 8px; border-radius: 6px; border: 1px solid #b4b2a9; background: #fff; cursor: pointer; width: 100%; }
      .roi-panel button.roi-btn.accent { background: #185fa5; border-color: #185fa5; color: #fff; }
      .roi-panel .roi-status { font-size: 11px; color: #854f0b; min-height: 14px; margin: 4px 0; }
      .roi-panel .roi-acc-item { background: #fff; border: 1px solid #e5e3da; border-radius: 6px; margin-bottom: 4px; overflow: hidden; }
      .roi-panel .roi-acc-head { display: flex; align-items: center; gap: 8px; padding: 7px 8px; cursor: pointer; }
      .roi-panel .roi-acc-head .roi-cat-name { flex: 1; font-weight: 600; }
      .roi-panel .roi-acc-body { display: none; padding: 0 8px 8px 30px; font-size: 11px; color: #5f5e5a; }
      .roi-panel .roi-acc-item.open .roi-acc-body { display: block; }
      .roi-panel .roi-acc-item.open .roi-chevron { transform: rotate(90deg); }
      .roi-panel .roi-chevron { transition: transform 0.1s; font-size: 12px; color: #888780; }
      .roi-panel .roi-badge { font-size: 10px; padding: 1px 6px; border-radius: 8px; }
      .roi-badge-confirmed { background: #e1f5ee; color: #0f6e56; }
      .roi-badge-estimated { background: #faeeda; color: #854f0b; }
      .roi-badge-unknown { background: #f1efe8; color: #444441; }
      .roi-panel .roi-src-item { margin-bottom: 3px; }
      .roi-panel .roi-empty { font-size: 11px; color: #888780; padding: 4px 2px; }
    `;
    document.head.appendChild(style);
  }

  function badgeClass(conf) {
    return conf === "確定" ? "roi-badge-confirmed" : conf === "推定" ? "roi-badge-estimated" : "roi-badge-unknown";
  }
  function escHtml(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function fmtNum(v) { const n = Number(v); return isNaN(n) ? String(v) : n.toLocaleString("ja-JP"); }

  async function mount(container, caseId, options = {}) {
    await ensureInit();
    const getMemo = options.getMemo || (() => container.querySelector(".roi-memo")?.value.trim() || "");

    container.innerHTML = `
      <div class="roi-panel">
        <div class="roi-h-row">
          <span class="roi-h">議事録・提案（${escHtml(caseId || "")}）</span>
          <div style="display:flex;gap:6px">
            <button class="roi-icon-btn roi-create" title="議事録・メモから課題を自動抽出"><i class="ti ti-sparkles"></i></button>
            <button class="roi-icon-btn accent roi-toggle" title="提案一覧"><i class="ti ti-list"></i></button>
          </div>
        </div>
        ${options.getMemo ? "" : `<textarea class="roi-memo" rows="2" placeholder="メモ（任意。議事録が無くてもここだけで抽出できます）"></textarea>`}
        <div class="roi-status"></div>
        <div class="roi-body" style="display:none">
          <div class="roi-acc-list"></div>
          <button class="roi-btn accent roi-build" style="display:none;margin-top:6px">提案書を作成（選択した課題の元データを表示）</button>
          <div class="roi-build-out"></div>
        </div>
      </div>
    `;

    const $ = sel => container.querySelector(sel);
    const setStatus = msg => { $(".roi-status").textContent = msg || ""; };

    $(".roi-toggle").addEventListener("click", async () => {
      const body = $(".roi-body");
      const showing = body.style.display !== "none";
      body.style.display = showing ? "none" : "";
      if (!showing) await refreshList();
    });

    $(".roi-create").addEventListener("click", async () => {
      const cfg = RoiCore.getConfig();
      if (!cfg.webhookUrl) { setStatus("AI連携エンドポイントが未設定です（提案ナレッジ側の設定で登録してください）"); return; }
      setStatus("議事録・メモから課題を判定中…");
      try {
        const cats = await RoiCore.autoExtractProposals(caseId, getMemo());
        setStatus(cats.length ? `${cats.length}件の課題を追加しました` : "該当する課題は見つかりませんでした");
        $(".roi-body").style.display = "";
        await refreshList();
      } catch (e) {
        console.warn(e);
        setStatus(e.message || "抽出に失敗しました");
      }
    });

    async function refreshList() {
      const listEl = $(".roi-acc-list");
      listEl.innerHTML = `<div class="roi-empty">読み込み中…</div>`;
      const summary = await RoiCore.getProposalSummaryForCase(caseId);
      if (!summary.length) {
        listEl.innerHTML = `<div class="roi-empty">まだ課題がありません。作成アイコンから抽出してください。</div>`;
        $(".roi-build").style.display = "none";
        return;
      }
      listEl.innerHTML = summary.map(it => `
        <div class="roi-acc-item" data-cat="${escHtml(it.category)}">
          <div class="roi-acc-head">
            <input type="checkbox" class="roi-check" ${it.selected ? "checked" : ""}>
            <span class="roi-cat-name">${escHtml(it.category)}</span>
            <span>${it.saving != null ? fmtNum(it.saving) + (it.unit || "") + "/年" : "—"}</span>
            <span class="roi-badge ${badgeClass(it.confidence)}">${escHtml(it.confidence)}</span>
            <i class="ti ti-chevron-right roi-chevron"></i>
          </div>
          <div class="roi-acc-body"><div class="roi-detail">詳細を読み込み中…</div></div>
        </div>`).join("");
      bindAccordionEvents();
      updateBuildButton();
    }

    function bindAccordionEvents() {
      container.querySelectorAll(".roi-acc-head").forEach(head => {
        head.addEventListener("click", async (e) => {
          if (e.target.classList.contains("roi-check")) return;
          const item = head.closest(".roi-acc-item");
          const willOpen = !item.classList.contains("open");
          item.classList.toggle("open");
          if (willOpen) await loadDetail(item);
        });
      });
      container.querySelectorAll(".roi-check").forEach(cb => {
        cb.addEventListener("click", e => e.stopPropagation());
        cb.addEventListener("change", async (e) => {
          const cat = e.target.closest(".roi-acc-item").dataset.cat;
          await RoiCore.toggleSelection(caseId, cat, e.target.checked);
          updateBuildButton();
        });
      });
    }

    async function loadDetail(item) {
      const cat = item.dataset.cat;
      const detailEl = item.querySelector(".roi-detail");
      const [rows, sources] = await Promise.all([
        RoiCore.getCalcRowsForCase(caseId, {}).then(all => all.filter(r => r.category === cat)),
        RoiCore.getSourceEntriesForCategory(caseId, cat),
      ]);
      const sol = (await RoiCore.getSolutions()).find(s => s.category === cat);
      const outputLines = rows.filter(r => r.kind === "出力").map(r => `${r.name} ${fmtNum(r.value)}${r.unit || ""}`).join("<br>");
      const srcHtml = sources.length
        ? sources.map(s => `<div class="roi-src-item">${s.url ? `<a href="${escHtml(s.url)}" target="_blank" rel="noopener">${escHtml(s.url)}</a>` : escHtml((s.text || "").slice(0, 40) + "…")}</div>`).join("")
        : `<div class="roi-src-item">元データのリンクは記録されていません</div>`;
      detailEl.innerHTML = `
        <div style="margin-bottom:4px">解決策：${escHtml(sol ? sol.name : "未設定")}</div>
        <div style="margin-bottom:6px">${outputLines || "（内訳なし）"}</div>
        <div style="font-weight:600;margin-bottom:2px">元データ</div>
        ${srcHtml}
      `;
    }

    function updateBuildButton() {
      const count = container.querySelectorAll(".roi-check:checked").length;
      $(".roi-build").style.display = count ? "" : "none";
      $(".roi-build").textContent = `提案書を作成（選択中 ${count}件の元データを表示）`;
    }

    $(".roi-build").addEventListener("click", async () => {
      const selectedCats = Array.from(container.querySelectorAll(".roi-acc-item"))
        .filter(item => item.querySelector(".roi-check").checked)
        .map(item => item.dataset.cat);
      const out = $(".roi-build-out");
      out.innerHTML = `<div class="roi-empty">読み込み中…</div>`;
      const blocks = await Promise.all(selectedCats.map(async cat => {
        const sources = await RoiCore.getSourceEntriesForCategory(caseId, cat);
        const list = sources.length
          ? sources.map(s => `<div class="roi-src-item">${s.url ? `<a href="${escHtml(s.url)}" target="_blank" rel="noopener">${escHtml(s.url)}</a>` : escHtml((s.text || "").slice(0, 40) + "…")}</div>`).join("")
          : `<div class="roi-src-item">元データのリンクは記録されていません</div>`;
        return `<div style="margin-top:6px"><b>${escHtml(cat)}</b>${list}</div>`;
      }));
      out.innerHTML = `<div style="font-weight:600;margin-top:8px">提案書データ・元データリンク一覧</div>${blocks.join("")}
        <div class="roi-empty" style="margin-top:6px">プロンプトの組み立て・提案書生成は提案ナレッジアドインで行ってください。</div>`;
    });
  }

  global.RoiPanel = { mount };
})(window);
