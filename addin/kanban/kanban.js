const APP_VERSION = "rev_20260427_1f1e933";

// window.APP_VERSIONも設定してindex.htmlから参照可能にする
window.APP_VERSION = APP_VERSION;

let allTasks = [];
let currentDraggedId = null;
let currentTask = null;

let selectedUser = null;
let selectedCategory = null;
let selectedPeriod = "all";
let showHeld = true; // 保留表示フラグ

// レーン高さ調整のデバウンス用変数
let heightAdjustTimeout = null;
let isAdjustingHeights = false;

Office.onReady(() => {
  // 保存されたサイズを復元
  restoreSavedSize();
  
  // 保存されたフィルター設定を復元
  restoreSavedFilters();
  
  // フィルターセクションの表示状態を復元
  restoreFilterSectionState();
  
  // 保留表示設定を復元
  restoreHeldDisplay();
  
  init();
});

// ===== サイズ記憶機能 =====
function restoreSavedSize() {
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    if (savedSize) {
      const size = JSON.parse(savedSize);
      
// 最小サイズの制限（デスクトップ版対応で極小に設定）
  const minWidth = 120;
  const minHeight = 300;
  const width = Math.max(size.width || 120, minWidth);
      const height = Math.max(size.height || 600, minHeight);
      
      // DOM要素のサイズを設定
      document.documentElement.style.minWidth = width + "px";
      document.body.style.minWidth = width + "px";
      
      // Office APIを使用してタスクペインのサイズを設定（可能な場合）
      if (Office.context.requirements.isSetSupported('TaskPaneApp', '1.1')) {
        try {
          Office.addin.setTaskpaneSize(width, height);
        } catch (e) {
          console.log("TaskPane resize not supported:", e);
        }
      }
      
      // 親ウィンドウへのサイズヒント
      if (window.parent && window.parent.postMessage) {
        window.parent.postMessage({
          type: 'resize',
          width: width,
          height: height
        }, '*');
      }
      
      console.log(`Restored size: ${width}x${height}`);
    } else {
      // デフォルトサイズを設定
      setDefaultSize();
    }
  } catch (e) {
    console.log("Size restoration error:", e);
    setDefaultSize();
  }
}

function setDefaultSize() {
  const defaultWidth = 120;
  const defaultHeight = 600;
  
  document.documentElement.style.minWidth = defaultWidth + "px";
  document.body.style.minWidth = defaultWidth + "px";
  
  if (window.parent && window.parent.postMessage) {
    window.parent.postMessage({
      type: 'resize',
      width: defaultWidth,
      height: defaultHeight
    }, '*');
  }
}

function setupSizeMonitoring() {
  let saveTimeout;
  let adjustTimeout;
  let resizeTimeout;
  let isPerformingAdjustment = false;
  
  // ペインサイズ追従のためのサイズ調整関数（デバウンス強化版）
  function performSizeAdjustment() {
    // 既に調整中の場合はスキップして無限ループを防ぐ
    if (isPerformingAdjustment) {
      return;
    }
    
    clearTimeout(adjustTimeout);
    adjustTimeout = setTimeout(() => {
      isPerformingAdjustment = true;
      try {
        // ペインの実際の幅と高さを取得
        const containerWidth = getActualPaneWidth();
        adjustLaneWidths(containerWidth);
        // 高さ調整は必要時のみ実行（チラつき防止）
        if (!isAdjustingHeights) {
          adjustLaneHeights();
        }
      } finally {
        // 調整完了後にフラグをリセット
        setTimeout(() => {
          isPerformingAdjustment = false;
        }, 100);
      }
    }, 100); // より長いディレイで安定化
  }
  
  // ResizeObserverでサイズ変更を監視（デバウンス強化）
  if (window.ResizeObserver) {
    const resizeObserver = new ResizeObserver(entries => {
      // ResizeObserver自体もデバウンスして頻繁な発動を制限
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(() => {
        for (const entry of entries) {
          const { width, height } = entry.contentRect;
          
          // サイズ調整を実行（内部でデバウンス済み）
          performSizeAdjustment();
          
          // ストレージ保存のデバウンス処理
          clearTimeout(saveTimeout);
          saveTimeout = setTimeout(() => {
            saveSizeToStorage(Math.round(width), Math.round(height));
          }, 800); // 保存はさらに遅延
        }
      }, 150); // ResizeObserver自体のデバウンス
    });
    
    // 複数の要素を監視して確実にサイズ変更を補捉
    resizeObserver.observe(document.documentElement); // メインコンテナ
    resizeObserver.observe(document.body); // body要素
    
    // board要素も監視（より正確な幅検知のため）
    const boardElement = document.getElementById('board');
    if (boardElement) {
      resizeObserver.observe(boardElement);
    }
  } else {
    // ResizeObserverが利用できない場合はwindowのresizeイベントを使用
    let lastWidth = getActualPaneWidth();
    let lastHeight = document.body.clientHeight || window.innerHeight;
    
    window.addEventListener('resize', () => {
      clearTimeout(saveTimeout);
      saveTimeout = setTimeout(() => {
        const currentWidth = getActualPaneWidth();
        const currentHeight = document.body.clientHeight || window.innerHeight;
        
        if (Math.abs(currentWidth - lastWidth) > 1 || Math.abs(currentHeight - lastHeight) > 1) {
          performSizeAdjustment();
          saveSizeToStorage(currentWidth, currentHeight);
          lastWidth = currentWidth;
          lastHeight = currentHeight;
        }
      }, 50); // より短いディレイ
    });
  }
  
  // MutationObserverでDOM変更も監視（カード追加時等）
  const mutationObserver = new MutationObserver(() => {
    performSizeAdjustment();
  });
  
  const boardElement = document.getElementById('board');
  if (boardElement) {
    mutationObserver.observe(boardElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['style']
    });
  }
  
  // 初期表示時にも調整を実行
  setTimeout(() => {
    performSizeAdjustment();
  }, 50);
  
  // ウィンドウフォーカス時にも再調整
  window.addEventListener('focus', () => {
    setTimeout(performSizeAdjustment, 100);
  });
}

// 実際のペイン幅を正確に取得する関数（デスクトップ版対応強化）
function getActualPaneWidth() {
  // 複数の方法で幅を取得し、最も正確なものを選択
  const bodyWidth = document.body.clientWidth;
  const bodyOffsetWidth = document.body.offsetWidth;
  const docElementWidth = document.documentElement.clientWidth;
  const windowWidth = window.innerWidth;
  
  // デスクトップ版とWeb版の差を考慮した幅取得
  let paneWidth = Math.max(bodyWidth || 0, docElementWidth || 0);
  
  // デスクトップ版で幅が正確に取得できない場合の対策
  if (paneWidth <= 0 || paneWidth > 2000) {
    paneWidth = Math.min(windowWidth || 300, bodyOffsetWidth || 300);
  }
  
  // 最低幅を保証（デスクトップ版対応で極小設定）
  return Math.max(paneWidth, 120);
}

// ===== レーン幅調整機能 =====
function adjustLaneWidths(containerWidth) {
  // 実際のコンテナ幅を正確に取得
  if (!containerWidth) {
    containerWidth = getActualPaneWidth();
  }
  
  // ペイン幅に完全追従するための計算（デスクトップ版対応でさらに極小設定）
  const margin = 1; // body marginを最小にして最大限活用
  const gap = 2; // レーン間ギャップを最小にしてデスクトップ版に対応
  const padding = 3; // レーン内paddingを最小にしてカード領域を確保
  
  // ボードの利用可能幅を最大限活用
  const boardTotalWidth = Math.max(containerWidth - (margin * 2), 120);
  
  // 保留レーンが表示されているか確認
  const heldLane = document.getElementById('held');
  const isHeldVisible = showHeld && heldLane && heldLane.style.display !== 'none';
  const laneCount = isHeldVisible ? 4 : 3;
  
  // ギャップの合計幅を算出
  const totalGapWidth = gap * (laneCount - 1);
  
  // レーンコンテンツの利用可能幅
  const availableWidth = boardTotalWidth - totalGapWidth;
  
  // 各レーンの幅を計算（等幅分割でペインをフル活用）
  let laneWidth = Math.floor(availableWidth / laneCount);
  
  // 最小幅の保証（カード48px + padding）をデスクトップ版対応で最小に
  const minLaneWidth = 48 + padding; 
  laneWidth = Math.max(laneWidth, minLaneWidth);
  
  // ボード全体をペイン幅に完全追従させる
  const boardElement = document.getElementById('board');
  if (boardElement) {
    boardElement.style.width = '100%'; // ペイン幅に完全追従
    boardElement.style.maxWidth = 'none';
    boardElement.style.minWidth = '120px';
    boardElement.style.boxSizing = 'border-box';
  }
  
  // CSSでレーン幅を動的に設定（ペイン全体を活用）
  const lanes = document.querySelectorAll('.lane');
  lanes.forEach(lane => {
    // 保留レーンが非表示の場合はスキップ
    if (lane.id === 'held' && !isHeldVisible) {
      lane.style.display = 'none';
      return;
    }
    
    // 保留レーンが表示される場合は表示
    if (lane.id === 'held' && isHeldVisible) {
      lane.style.display = 'flex';
    }
    
    // レーン幅をペインに完全追従させる
    lane.style.width = laneWidth + 'px';
    lane.style.minWidth = minLaneWidth + 'px';
    lane.style.maxWidth = laneWidth + 'px';
    lane.style.flex = `0 0 ${laneWidth}px`; // 明示的なflex-basisを指定
    lane.style.flexShrink = '0'; // 縮小を防ぐ
    lane.style.flexGrow = '0'; // 拡大を防ぐ
    lane.style.boxSizing = 'border-box';
  });
  
  // カードをレーン幅に完全追従させる
  adjustCardWidths(laneWidth - padding);
  
  // console.log(`Pane-following adjustment: Pane=${containerWidth}px, Lane=${laneWidth}px (${laneCount} lanes), Board=${boardTotalWidth}px`);
}

// カード幅の調整 - ペイン幅に完全追従
function adjustCardWidths(maxCardWidth) {
  const cards = document.querySelectorAll('.card');
  cards.forEach(card => {
    // カードをレーン幅に完全追従させる
    card.style.width = '100%'; 
    card.style.minWidth = '48px'; // 最小幅をデスクトップ版対応で最小に
    card.style.maxWidth = 'none'; // 最大幅制限を完全解除
    card.style.boxSizing = 'border-box';
    card.style.wordWrap = 'break-word'; // 長いテキストの折り返し
    card.style.overflowWrap = 'break-word';
    card.style.flexShrink = '0'; // カードの縮小を制限
    card.style.width = '100%'; // レーンの全幅を使用
  });
}

// 実際のペイン高さを正確に取得する関数
function getActualPaneHeight() {
  // 複数の方法で高さを取得し、最も正確なものを選択
  const bodyHeight = document.body.clientHeight;
  const bodyOffsetHeight = document.body.offsetHeight;
  const docElementHeight = document.documentElement.clientHeight;
  const windowHeight = window.innerHeight;
  
  // タスクペインの場合は通常bodyのclientHeightが最も適切
  let paneHeight = Math.max(bodyHeight || 0, docElementHeight || 0);
  
  // フォールバックとしてwindowHeightを使用
  if (paneHeight <= 0) {
    paneHeight = windowHeight || 600;
  }
  
  // 上部要素（フィルター等）の高さを動的に計算
  const boardElement = document.getElementById('board');
  let topBarHeight = 150; // デフォルト値
  
  if (boardElement) {
    // ボード要素の位置から上部要素の高さを計算
    const boardRect = boardElement.getBoundingClientRect();
    topBarHeight = Math.max(boardRect.top, 100); // 最小100pxを保証
  }
  
  // 利用可能高さを計算（下部余白も考慮）
  const availableHeight = Math.max(paneHeight - topBarHeight - 20, 200);
  
  return availableHeight;
}

// ===== レーン高さ調整機能（チラつき防止版） =====
function adjustLaneHeights() {
  // 既に調整中の場合はスキップしてチラつきを防ぐ
  if (isAdjustingHeights) {
    return;
  }
  
  // デバウンス処理を200msに延長してスクロール操作との競合を削減
  clearTimeout(heightAdjustTimeout);
  heightAdjustTimeout = setTimeout(() => {
    performHeightAdjustment();
  }, 200);
}

function performHeightAdjustment() {
  if (isAdjustingHeights) return;
  isAdjustingHeights = true;
  
  try {
    const lanes = document.querySelectorAll('.lane');
    const maxPaneHeight = getActualPaneHeight();
    
    const visibleLanes = Array.from(lanes).filter(lane => 
      lane.style.display !== 'none'
    );
    
    if (visibleLanes.length === 0) {
      isAdjustingHeights = false;
      return;
    }
    
    // 1. 一時的に全レーンの高さをリセット（非表示で測定）
    visibleLanes.forEach(lane => {
      lane.style.visibility = 'hidden'; // チラつき防止
      lane.style.height = 'auto';
      lane.style.minHeight = 'auto';
      lane.style.maxHeight = 'none';
    });
    
    // 2. DOM更新を強制実行してから測定（同期的）
    document.body.offsetHeight; // reflowを強制実行
    
    let maxNaturalHeight = 0;
    
    // 3. 各レーンの実際の高さを測定
    visibleLanes.forEach(lane => {
      const naturalHeight = lane.offsetHeight;
      if (naturalHeight > maxNaturalHeight) {
        maxNaturalHeight = naturalHeight;
      }
    });
    
    // 4. 最終的なレーン高さを決定
    const finalHeight = Math.min(
      Math.max(maxNaturalHeight, 150), 
      maxPaneHeight
    );
    
    // 5. 全レーンを同じ高さに統一（一括変更）
    visibleLanes.forEach(lane => {
      // スクロール位置を保存
      const cardList = lane.querySelector('.card-list');
      const scrollTop = cardList ? cardList.scrollTop : 0;
      
      lane.style.height = finalHeight + 'px';
      lane.style.minHeight = finalHeight + 'px';
      lane.style.maxHeight = finalHeight + 'px';
      lane.style.overflowY = 'auto'; // スクロール可能に変更
      lane.style.overflowX = 'hidden';
      lane.style.visibility = 'visible'; // 表示を復活
      
      // スクロール位置を復元
      if (cardList && scrollTop > 0) {
        cardList.scrollTop = scrollTop;
      }
    });
    
    // console.log(`Lane heights optimized: natural-max=${maxNaturalHeight}px, final=${finalHeight}px, pane-limit=${maxPaneHeight}px`);
    
  } finally {
    isAdjustingHeights = false;
  }
}

function saveSizeToStorage(width, height) {
  try {
    const sizeData = {
      width: width,
      height: height,
      timestamp: Date.now()
    };
    
    localStorage.setItem('kanban-taskpane-size', JSON.stringify(sizeData));
    console.log(`Saved size: ${width}x${height}`);
  } catch (e) {
    console.log("Size saving error:", e);
  }
}

// サイズ設定をリセットする関数
function resetSize() {
  try {
    localStorage.removeItem('kanban-taskpane-size');
    localStorage.removeItem('kanban-filters');
    console.log("Size and filter settings reset");
    
    // ページをリロードして新しい設定を適用
    window.location.reload();
  } catch (e) {
    console.log("Reset error:", e);
  }
}

// ===== フィルター記憶機能 =====
function restoreSavedFilters() {
  try {
    const savedFilters = localStorage.getItem('kanban-filters');
    if (savedFilters) {
      const filters = JSON.parse(savedFilters);
      
      selectedUser = filters.user || null;
      selectedCategory = filters.category || null;
      selectedPeriod = filters.period || "all";
      
      console.log('Restored filters:', filters);
    }
  } catch (e) {
    console.log("Filter restoration error:", e);
    // エラー時はデフォルト値を維持
    selectedUser = null;
    selectedCategory = null; 
    selectedPeriod = "all";
  }
}

function saveFilters() {
  try {
    const filterData = {
      user: selectedUser,
      category: selectedCategory,
      period: selectedPeriod,
      timestamp: Date.now()
    };
    
    localStorage.setItem('kanban-filters', JSON.stringify(filterData));
    console.log('Saved filters:', filterData);
  } catch (e) {
    console.log("Filter saving error:", e);
  }
}

async function init() {
  // 保存されたサイズまたはデフォルトサイズを適用
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    let minWidth = 200;
    
    if (savedSize) {
      const size = JSON.parse(savedSize);
      minWidth = Math.max(size.width || 200, 200);
    }
    
    document.documentElement.style.minWidth = minWidth + "px";
    document.body.style.minWidth = minWidth + "px";
  } catch (e) {
    // エラー時はデフォルトサイズ
    document.documentElement.style.minWidth = "200px";
    document.body.style.minWidth = "200px";
  }
  
  // ボードコンテナをペイン幅に完全追従させる
  const boardEl = document.getElementById("board");
  if (boardEl) {
    boardEl.style.width = "100%"; // ペイン幅に完全追従
    boardEl.style.maxWidth = "100%";
    boardEl.style.minWidth = "200px";
    boardEl.style.boxSizing = "border-box";
  }

  await loadExcelData();
  renderFilters();
  renderBoard();
  renderPeriodFilter();
  
  // サイズ監視を開始（初期化後に実行）
  setupSizeMonitoring();
  
  // 初期化完了後に一度だけレイアウト調整
  setTimeout(() => {
    const containerWidth = getActualPaneWidth();
    adjustLaneWidths(containerWidth);
    adjustLaneHeights();
  }, 500); // 遅延を延長してDOM安定化を確実に待つ
  
  // バージョン表示を更新
  if (typeof updateVersionDisplay === 'function') {
    updateVersionDisplay();
  }
}

// ===== Excel日付変換 =====
function excelDateToJS(value) {
  if (!value) return null;
  if (typeof value === "number") {
    return new Date((value - 25569) * 86400 * 1000);
  }
  return new Date(value);
}

function fmt(v) {
  const d = excelDateToJS(v);
  if (!d || isNaN(d)) return "";
  return `${d.getMonth()+1}/${d.getDate()}`;
}

// ===== データ取得 =====
async function loadExcelData() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    const rows = range.values;

    allTasks = rows.slice(10).map((row, i) => {
      if (!row[25] || row[19] === "-") return null;

      const t = {
        id: row[24],
        category: row[0],
        classification: row[1],  // B列の分類を追加
        title: row[25],
        user: row[13],
        start: row[15],
        end: row[16],
        actualStart: row[17],
        actualEnd: row[18],
        note: row[14],
        rowIndex: i + 11,

        isNoSchedule: !row[15] && !row[16],
        isStar: row[14] && row[14].toString().startsWith('★')  // 備考の先頭に★があるかチェック
      };

      t.status = getStatus(t);
      return t;
    }).filter(x => x);
  });
}

// ===== ステータス =====
function getStatus(t) {
  if (t.actualEnd) return "完了";
  if (t.actualStart) return "対応中";
  return "未着手";
}

// ===== フィルタ =====
function renderFilters() {
  renderUserFilter();
  renderCategoryFilter();
}

function renderUserFilter() {
  const users = [...new Set(
    allTasks
      .map(t => t.user)
      .filter(v => v && v !== "#")
  )];

  const el = document.getElementById("user-filters");
  el.innerHTML = "";

  users.forEach(u => {
    const b = document.createElement("button");
    b.textContent = u;

    if (selectedUser === u) b.classList.add("active");

    b.onclick = () => {
      selectedUser = (selectedUser === u) ? null : u;
      saveFilters(); // フィルタ設定を保存
      renderBoard();
      renderFilters();
    };

    el.appendChild(b);
  });
}

function renderCategoryFilter() {
  const cats = [...new Set(
    allTasks
      .map(t => t.category)
      .filter(v => v && v !== "#")
  )];

  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    const b = document.createElement("button");
    b.textContent = c;

    if (selectedCategory === c) b.classList.add("active");

    b.onclick = () => {
      selectedCategory = (selectedCategory === c) ? null : c;
      saveFilters(); // フィルタ設定を保存
      renderBoard();
      renderFilters();
    };

    el.appendChild(b);
  });
}

function setPeriod(p) {
  selectedPeriod = (selectedPeriod === p) ? "all" : p;
  saveFilters(); // フィルタ設定を保存
  renderBoard();
  renderPeriodFilter();
}

function renderPeriodFilter() {
  document.querySelectorAll("[data-period]").forEach(b => {
    b.classList.toggle("active", b.dataset.period === selectedPeriod);
  });
}

// ===== 保留表示切替 =====
function toggleHeldDisplay() {
  showHeld = document.getElementById('show-held').checked;
  localStorage.setItem('kanban-show-held', showHeld);
  renderBoard();
  
  // 保留レーンの表示切替後に必要最小限の調整
  setTimeout(() => {
    const containerWidth = getActualPaneWidth();
    adjustLaneWidths(containerWidth);
    // 高さは必要時のみ調整（チラつき防止）
  }, 100);
}

// 保留表示設定を復元
function restoreHeldDisplay() {
  const saved = localStorage.getItem('kanban-show-held');
  showHeld = saved !== null ? saved === 'true' : true;
  document.getElementById('show-held').checked = showHeld;
}

// ===== 描画 =====
function renderBoard() {
  ["todo","held","doing","done"].forEach(l => {
    const lane = document.querySelector(`#${l} .card-list`);
    if (lane) {
      lane.innerHTML = "";
    }
  });

  // 保留レーンの表示/非表示切替
  const heldLane = document.getElementById('held');
  if (heldLane) {
    heldLane.style.display = showHeld ? 'block' : 'none';
  }

  const filtered = allTasks.filter(isMatch);

  const normal = filtered
    .filter(t => t.status !== "完了")
    .sort((a, b) => {
      // スター付きを優先
      if (a.isStar && !b.isStar) return -1;
      if (!a.isStar && b.isStar) return 1;
      // 同じスター状態なら期限日順
      return excelDateToJS(a.end) - excelDateToJS(b.end);
    });

  const done = filtered
    .filter(t => t.status === "完了")
    .sort((a,b)=>excelDateToJS(b.actualEnd)-excelDateToJS(a.actualEnd));

  [...normal, ...done].forEach(t=>{
    const lane = getLane(t);  // タスク全体を渡すように変更
    document.querySelector(`#${lane} .card-list`).appendChild(createCard(t));
  });

  setupDnD();
  
  // カード描画後のレイアウト調整を最小限に抑制
  setTimeout(() => {
    const containerWidth = getActualPaneWidth();
    adjustLaneWidths(containerWidth);
    // 高さ調整は必要時のみ実行（スクロール体験を優先）
  }, 200);
}

// ===== カード =====
function createCard(t) {
  const d = document.createElement("div");
  d.className = "card";
  d.draggable = true;

  d.addEventListener("dragstart", (e) => {
    currentDraggedId = t.id;
    e.dataTransfer.setData("text/plain", t.id);
    d.classList.add("dragging");
  });

  d.addEventListener("dragend", () => {
    d.classList.remove("dragging");
  });

  d.addEventListener("click", (e) => {
    if (e.button !== 0) return;
    jumpToExcel(t.rowIndex);
  });

  d.addEventListener("contextmenu", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    await openModal(t);
  });

  const row1 = document.createElement("div");
  row1.className = "card-row1";

  const left = document.createElement("span");
  const rightGroup = document.createElement("span");
  rightGroup.className = "right-group";
  
  const user = document.createElement("span");
  user.className = "user-name";
  user.textContent = t.user || "";

  // 日付情報を設定
  if (t.isNoSchedule) {
    left.textContent = "TODO";
  } else if (t.status === "未着手" || t.status === "保留") {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)}`;
  } else if (t.status === "対応中") {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)} → ${fmt(t.actualStart)}～`;
  } else {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)} → ${fmt(t.actualStart)}～${fmt(t.actualEnd)}`;
  }

  // ユーザー名を右グループに追加
  rightGroup.appendChild(user);
  
  // 完了状態以外にのみスターアイコンを追加
  if (t.status !== "完了") {
    const star = document.createElement("span");
    star.className = "star-icon";
    star.textContent = t.isStar ? "★" : "☆";
    
    // ★の場合は金色クラスを追加
    if (t.isStar) {
      star.classList.add("filled");
    }
    
    star.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleStar(t);
    });
    rightGroup.appendChild(star);
  }

  // 日付情報と右グループ（ユーザー名+スター）を追加
  row1.appendChild(left);
  row1.appendChild(rightGroup);

  // タイトル行（タイトル + 分類）
  const row2 = document.createElement("div");
  row2.className = "card-title-row";
  
  const titleSpan = document.createElement("span");
  titleSpan.className = "card-title";
  titleSpan.textContent = t.title;
  
  const classificationSpan = document.createElement("span");
  classificationSpan.className = "card-classification";
  if (t.classification && t.classification.trim() !== "") {
    classificationSpan.textContent = `<${t.classification}>`;
  }
  
  row2.appendChild(titleSpan);
  row2.appendChild(classificationSpan);

  d.appendChild(row1);
  d.appendChild(row2);

  // スター状態に応じてカードスタイルを適用 
  if (t.isStar) {
    d.classList.add("starred");
  }

  applyColor(d, t);

  return d;
}

// ===== 色 =====
function applyColor(el, t) {
  if (t.status === "完了") {
    el.style.border = "2px solid #333";
    return;
  }

  const startRaw = excelDateToJS(t.start);
  const endRaw = excelDateToJS(t.end);

  if (!startRaw || !endRaw) return;

  // 時刻情報を除去して日付のみで比較
  const start = new Date(startRaw);
  start.setHours(0,0,0,0);
  
  const end = new Date(endRaw);
  end.setHours(0,0,0,0);
  
  const today = new Date();
  today.setHours(0,0,0,0);

  // ★ 遅延
  if (end < today) {
    el.style.border = "2px solid red";
    return;
  }

  // ★ 期間内（←ここが今回のポイント）
  if (start <= today && end >= today) {
    el.style.border = "2px solid green";
    return;
  }

  el.style.border = "1px solid #ccc";
}

// ===== DnD =====
function setupDnD() {
  ["todo","held","doing","done"].forEach(id=>{
    const lane = document.getElementById(id);

    lane.ondragover = (e)=>e.preventDefault();

    lane.ondrop = (e)=>{
      e.preventDefault();
      const t = allTasks.find(x=>x.id===currentDraggedId);
      if (t) updateStatus(t, id);
    };
  });
}

// ===== Excel =====
async function jumpToExcel(row){
  await Excel.run(async (ctx)=>{
    const s = ctx.workbook.worksheets.getItem("wbs");
    s.activate();
    s.getRange(`${row}:${row}`).select();
    await ctx.sync();
  });
}

// ===== util =====
function getLane(task){
  // 備考欄に▲がある場合は保留レーンに表示
  if (task.note && task.note.includes('▲')) {
    return "held";
  }
  
  // 通常のステータス判定
  const s = task.status;
  if(s==="未着手") return "todo";
  if(s==="保留") return "held";
  if(s==="対応中") return "doing";
  return "done";
}

function getMonday(d){
  const t=new Date(d);
  const day=t.getDay();
  const diff=t.getDate()-day+(day===0?-6:1);
  return new Date(t.setDate(diff));
}

function addDays(d,n){
  const t=new Date(d);
  t.setDate(t.getDate()+n);
  return t;
}

// JavaScriptのDateをExcelシリアル値に変換
function dateToExcelSerial(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return "";
  
  // Excel epoch: 1900年1月1日
  const excelEpoch = new Date(1900, 0, 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  
  // 日数差を計算
  const daysDiff = Math.floor((date - excelEpoch) / msPerDay);
  
  // Excelの1900年うるう年バグを考慮（1900年3月1日以降は+1）
  return daysDiff + (date >= new Date(1900, 2, 1) ? 2 : 1);
}

async function updateStatus(task, lane) {
  let actualStart = task.actualStart;
  let actualEnd = task.actualEnd;

  if (lane === "todo") {
    actualStart = "";
    actualEnd = "";
    
    // 保留レーンから移動した場合は▲→△に変更、ステータスも更新
    if (task.note && task.note.includes('▲')) {
      let newNote = task.note.replace(/▲/g, "△");
      task.note = newNote;
      await updateTaskStatus(task, "未着手");
    }
  }

  if (lane === "held") {
    // 保留状態：基本的には実績日時は変更しない
    // ただし、完了から保留に移動した場合のみ実績完了日を空にする
    if (task.status === "完了") {
      actualEnd = ""; // 完了からの移動時のみ実績完了日をクリア
    }
    // actualStart は常に維持
    
    // 保留レーンにドラッグ：△→▲に変更
    let newNote = ensureStatusSymbols(task.note || "");
    if (newNote.includes('△')) {
      newNote = newNote.replace(/△/g, "▲");
    } else if (!newNote.includes('▲')) {
      // 既に▲がない場合は追加
      const lines = newNote.split('\n');
      lines[0] = lines[0].replace(/△/, '') + '▲';
      newNote = lines.join('\n');
    }
    task.note = newNote;
    
    // ステータスも「保留」に変更
    await updateTaskStatus(task, "保留");
  }

  if (lane === "doing") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = "";
    
    // 保留レーンから移動した場合は▲→△に変更、ステータスも更新
    if (task.note && task.note.includes('▲')) {
      let newNote = task.note.replace(/▲/g, "△");
      task.note = newNote;
      await updateTaskStatus(task, "対応中");
    }
  }

  if (lane === "done") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = new Date();
    
    // 完了時は★→☆に変更
    if (task.isStar) {
      task.isStar = false;
    }
    
    // 保留レーンから移動した場合は▲→△に変更、ステータスも更新
    if (task.note && task.note.includes('▲')) {
      let newNote = task.note.replace(/▲/g, "△");
      task.note = newNote;
      await updateTaskStatus(task, "完了");
    }
  }

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;

    const startCell = sheet.getRange(`R${row}`);
    const endCell = sheet.getRange(`S${row}`);

    // Date型をExcelシリアル値に変換して設定
    startCell.values = [[dateToExcelSerial(actualStart)]];
    endCell.values = [[dateToExcelSerial(actualEnd)]];

    // 表示形式をm/d に設定
    startCell.numberFormat = [["m/d"]];
    endCell.numberFormat = [["m/d"]];

    // 備考の更新（保留状態変更やスター削除）
    if ((lane === "done" || lane === "doing" || lane === "held") && task.note !== undefined) {
      const noteCell = sheet.getRange(`O${row}`);
      noteCell.values = [[task.note]];
      noteCell.format.wrapText = false;
    }

    // 完了時に備考から★を削除
    if (lane === "done" && task.note && task.note.includes('★')) {
      let newNote = (task.note || "").replace(/★/g, "");
      const noteCell = sheet.getRange(`O${row}`);
      noteCell.values = [[newNote]];
      noteCell.format.wrapText = false;
      task.note = newNote; // タスクの備考も更新
    }
    
    await ctx.sync();
  });

  await init();
}

// ===== ステータス文字列更新 =====
async function updateTaskStatus(task, newStatus) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;
    const statusCell = sheet.getRange(`H${row}`);
    
    statusCell.values = [[newStatus]];
    
    await ctx.sync();
  });
  
  task.status = newStatus;
}

function isValidDate(v) {
  return v instanceof Date && !isNaN(v);
}

// ===== スター切り替え =====
async function toggleStar(task) {
  // スター状態を切り替え
  task.isStar = !task.isStar;
  
  // 備考を更新
  let newNote = task.note || "";
  
  if (task.isStar) {
    // ★を★に変更：先頭に★を付与
    if (!newNote.startsWith('★')) {
      newNote = '★' + newNote;
    }
  } else {
    // ★を☆に変更：備考から★を完全に削除（""に置換）
    newNote = newNote.replace(/★/g, "");
  }
  
  // Excelに備考を更新
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;
    const cell = sheet.getRange(`O${row}`);
    
    cell.values = [[newNote]];
    cell.format.wrapText = false;
    
    await ctx.sync();
  });
  
  // タスクの備考を更新
  task.note = newNote;
  
  // 画面を再描画
  renderBoard();
}

// ===== ステータス記号管理 =====
function ensureStatusSymbols(noteText) {
  if (!noteText) noteText = "";
  
  // 行で分割
  const lines = noteText.split('\n');
  let firstLine = lines[0] || "";
  
  // ★/☆がない場合、☆を追加
  if (!firstLine.includes('★') && !firstLine.includes('☆')) {
    firstLine = '☆' + firstLine;
  }
  
  // ▲/△がない場合、△を追加
  if (!firstLine.includes('▲') && !firstLine.includes('△')) {
    firstLine = firstLine + '△';
  }
  
  // 1行目を更新して復元
  lines[0] = firstLine;
  return lines.join('\n');
}

async function openModal(task) {
  currentTask = task;

  // O列から最新の備考内容を取得
  let originalNote = "";
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem("wbs");
      const noteCell = sheet.getRange(`O${task.rowIndex}`);
      noteCell.load("values");
      await ctx.sync();
      
      originalNote = (noteCell.values[0][0] || "").toString();
    });
  } catch (error) {
    console.log("O列の読み取りエラー:", error);
    // エラーの場合はタスクオブジェクトの値を使用
    originalNote = task.note || "";
  }
  
  // 備考欄のテンプレート処理
  let displayNote = originalNote;
  
  if (!displayNote.trim()) {
    // 完全に空の場合：ステータス記号とテンプレートを追加
    displayNote = "☆△\n＜タスク＞\n＜状況＞";
  } else {
    // ステータス記号を確認・追加
    displayNote = ensureStatusSymbols(displayNote);
    
    // 内容がある場合：行数をチェック
    const lines = displayNote.split('\n');
    
    // 2行目がない場合（1行のみまたは空行のみ）
    if (lines.length < 2 || (lines.length === 2 && !lines[1].trim())) {
      displayNote = displayNote.trimEnd() + "\n＜タスク＞\n＜状況＞";
    }
  }
  
  document.getElementById("modal-title").textContent = task.title;
  document.getElementById("modal-note").value = displayNote;

  const modal = document.getElementById("modal");
  modal.classList.remove("hidden");
  
  // Escキーでモーダルを閉じる
  const handleEscKey = (event) => {
    if (event.key === 'Escape') {
      closeModal();
    }
  };
  
  // モーダル外クリックで閉じる（変更がない場合のみ）
  const handleOverlayClick = (event) => {
    if (event.target === modal) {
      const currentNote = document.getElementById("modal-note").value;
      // 変更がない場合のみ閉じる（O列から取得した最新値と比較）
      if (currentNote === displayNote) {
        closeModal();
      }
    }
  };
  
  // モーダルコンテンツ内のクリックでイベント伝播を止める
  const modalContent = modal.querySelector('.modal-content');
  const handleContentClick = (event) => {
    event.stopPropagation();
  };
  
  // イベントリスナーを追加
  document.addEventListener('keydown', handleEscKey);
  modal.addEventListener('click', handleOverlayClick);
  modalContent.addEventListener('click', handleContentClick);
  
  // クリーンアップ関数をモーダルに保存
  modal._cleanup = () => {
    document.removeEventListener('keydown', handleEscKey);
    modal.removeEventListener('click', handleOverlayClick);
    modalContent.removeEventListener('click', handleContentClick);
  };
  
  // テキストエリアにフォーカス
  setTimeout(() => {
    document.getElementById("modal-note").focus();
  }, 100);
}

function closeModal() {
  const modal = document.getElementById("modal");
  modal.classList.add("hidden");
  
  // イベントリスナーをクリーンアップ
  if (modal._cleanup) {
    modal._cleanup();
    modal._cleanup = null;
  }
}

async function saveNote() {
  const note = document.getElementById("modal-note").value;

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = currentTask.rowIndex;

    const cell = sheet.getRange(`O${row}`);

    cell.values = [[note]];

    // ★これ追加
    cell.format.wrapText = false;

    // ★行高さ固定（例：20）
    const entireRow = sheet.getRange(`${row}:${row}`);
    entireRow.format.rowHeight = 20;

    await ctx.sync();
  });

  // タスクオブジェクトの備考を直接更新（Excel再読み込みを回避）
  if (currentTask) {
    currentTask.note = note;
    // スター状態も更新
    currentTask.isStar = note.startsWith('★');
  }

  closeModal();
  
  // 軽量な再描画（Excelデータ再読み込みなし）
  renderBoard();
}

function isMatch(t) {

  // 担当者
  if (selectedUser && t.user !== selectedUser) return false;

  // 分類
  if (selectedCategory && t.category !== selectedCategory) return false;

  // ★ 本日フィルタ：スター付きのみ表示
  if (selectedPeriod === "today") {
    return t.isStar;
  }

  // ★ 日付なし（TODO）
  if (t.isNoSchedule) {
    return selectedPeriod === "all" || selectedPeriod === "todo";
  }

  // ★ TODOフィルタが選択されている場合、日付ありのタスクは除外
  if (selectedPeriod === "todo") {
    return false;
  }

  const start = excelDateToJS(t.start);
  const end = excelDateToJS(t.end);

  if (!start || !end) return false;

  const today = new Date();
  today.setHours(0,0,0,0);

  const monday = getMonday(today);
  const sunday = addDays(monday, 6);
  const nextMonday = addDays(monday, 7);
  const nextSunday = addDays(monday, 13);

  switch (selectedPeriod) {

    case "past":
      return end < monday;

    case "week":
      return (start <= sunday && end >= monday);

    case "nextweek":
      return (start <= nextSunday && end >= nextMonday);

    case "future":
      return start > nextSunday;

    case "all":
    default:
      return true;
  }
}

// ===== フィルタエリア表示制御 =====
// フィルタセクションの表示状態を管理
let filterSectionState = {
  'user-filter': false,
  'category-filter': false,
  'period-filter': false
};

// フィルタエリアの表示/非表示を切り替え
function toggleFilterSection(sectionId) {
  // 現在の状態を切り替え
  filterSectionState[sectionId] = !filterSectionState[sectionId];
  
  // DOM要素を取得
  const section = document.getElementById(sectionId);
  const button = document.getElementById(`btn-${sectionId}`);
  
  if (section && button) {
    // 表示状態を更新
    if (filterSectionState[sectionId]) {
      section.classList.remove('hidden');
      button.classList.add('active');
    } else {
      section.classList.add('hidden');
      button.classList.remove('active');
    }
    
    // 状態をローカルストレージに保存
    saveFilterSectionState();
    
    // レイアウトを再調整
    setTimeout(() => {
      const containerWidth = getActualPaneWidth();
      adjustLaneWidths(containerWidth);
    }, 100);
  }
}

// フィルタセクションの状態をローカルストレージに保存
function saveFilterSectionState() {
  try {
    localStorage.setItem('kanban-filter-sections', JSON.stringify(filterSectionState));
  } catch (e) {
    console.log("Filter section state saving error:", e);
  }
}

// フィルタセクションの状態をローカルストレージから復元
function restoreFilterSectionState() {
  try {
    const saved = localStorage.getItem('kanban-filter-sections');
    if (saved) {
      filterSectionState = { ...filterSectionState, ...JSON.parse(saved) };
    }
    
    // DOM更新
    Object.keys(filterSectionState).forEach(sectionId => {
      const section = document.getElementById(sectionId);
      const button = document.getElementById(`btn-${sectionId}`);
      
      if (section && button) {
        if (filterSectionState[sectionId]) {
          section.classList.remove('hidden');
          button.classList.add('active');
        } else {
          section.classList.add('hidden');
          button.classList.remove('active');
        }
      }
    });
  } catch (e) {
    console.log("Filter section state restoration error:", e);
  }
}