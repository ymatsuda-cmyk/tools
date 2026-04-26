/* バグ管理 アドイン
   - バグシートの 2行目=項目名 / 3行目=入力例 / 4行目以降=データ
   - Excel.js があればExcel連携、なければデモデータでブラウザ単体動作
*/
(function () {
  'use strict';

  const SHEET_NAME = 'バグ';
  const HEADER_ROW = 2;
  const SAMPLE_ROW = 3;
  const DATA_START = 4;
  const COL_COUNT  = 28;

  // 動的にフィールド定義を生成する関数
  function getColumns() {
    return [
      { key: 'id',         letter: 'A', label: 'ID',           group: '基本情報', type: 'readonly' },
      { key: 'title',      letter: 'B', label: 'タイトル',      group: '基本情報', type: 'text' },
      { key: 'status',     letter: 'C', label: '状況',         group: '基本情報', type: 'select', options: ['新規','解析','修正','確認','再発','完了'] },
      { key: 'updated',    letter: 'D', label: '更新日',       group: '基本情報', type: 'date' },
      { key: 'assignee',   letter: 'E', label: '担当者',       group: '基本情報', type: 'select', options: ['', ...ASSIGNEE_ORDER.slice(1)] }, // 動的に設定
      { key: 'occurredOn', letter: 'F', label: '発生日',       group: '発生情報', type: 'date' },
      { key: 'reporter',   letter: 'G', label: '登録者',       group: '発生情報', type: 'select', options: ['', ...REPORTER_LIST] }, // 動的に設定
      { key: 'origin',     letter: 'H', label: '発生起因',     group: '発生情報', type: 'select', options: ['','定義(通常)','定義(電源断)','定義(通信断)'] },
      { key: 'originNumber', letter: 'I', label: '起因番号',   group: '発生情報', type: 'text' }, // 新規追加
      { key: 'steps',      letter: 'J', label: '再現手順',     group: '発生情報', type: 'textarea' },
      { key: 'expected',   letter: 'K', label: '期待する動作', group: '発生情報', type: 'textarea' },
      { key: 'actual',     letter: 'L', label: '実際の動作',   group: '発生情報', type: 'textarea' },
      { key: 'reproRate',  letter: 'M', label: '再現率',       group: '発生情報', type: 'select', options: ['','毎回','時々','1回のみ'] },
      { key: 'cause',      letter: 'N', label: '原因',         group: '対応情報', type: 'textarea' },
      { key: 'analyst',    letter: 'O', label: '解析者',       group: '対応情報', type: 'select', options: ['', ...REPORTER_LIST] }, // 動的に設定
      { key: 'analysisDate', letter: 'P', label: '解析日',     group: '対応情報', type: 'date' },
      { key: 'scope',      letter: 'Q', label: '影響範囲',     group: '対応情報', type: 'select', options: ['','定義(通常)','定義(電源断)','定義(通信断)','RPA','アプリ'] },
      { key: 'fix',        letter: 'R', label: '対応内容',     group: '対応情報', type: 'textarea' },
      { key: 'fixVer',     letter: 'S', label: '修正Ver',     group: '対応情報', type: 'text' },
      { key: 'fixer',      letter: 'T', label: '対応者',       group: '対応情報', type: 'select', options: ['', ...REPORTER_LIST] }, // 動的に設定
      { key: 'fixDate',    letter: 'U', label: '対応日',       group: '対応情報', type: 'date' },
      { key: 'verify',     letter: 'V', label: '確認内容',     group: '結果確認', type: 'textarea' },
      { key: 'reject',     letter: 'W', label: '差し戻し',     group: '結果確認', type: 'text' },
      { key: 'verifier',   letter: 'X', label: '確認者',       group: '結果確認', type: 'select', options: ['', ...REPORTER_LIST] }, // 動的に設定
      { key: 'verifyDate', letter: 'Y', label: '確認日',       group: '結果確認', type: 'date' },
      { key: 'tag',        letter: 'Z', label: 'タグ',         group: '管理',     type: 'text' },
      { key: 'priority',   letter: 'AA', label: '優先度',       group: '管理',     type: 'select', options: ['','高','中','低'] },
      { key: 'severity',   letter: 'AB', label: '影響度',      group: '管理',     type: 'select', options: ['','致命的','重大','警備'] }
    ];
  }

  const STATUS_ORDER = ['新規','解析','修正','確認','完了'];
  
  // 動的に設定されるマッピング（セルから取得）
  let STATUS_DISPLAY_NAMES = {};
  let ASSIGNEE_ORDER = ['(未割当)'];
  let REPORTER_LIST = [];
  const PRIORITY_RANK = { '高': 0, '中': 1, '低': 2, '': 3 };

  const state = {
    bugs: [],
    view: 'status',  // 'assignee' または 'status'
    filters: { text: '', priority: '', status: '' },
    inOffice: false,
    editingRow: null,
    presetTags: [], // プリセットタグを保存
    lastSelectedReporter: localStorage.getItem('bugTracker_lastReporter') || '' // 前回選択した登録者
  };

  function $(sel) { return document.querySelector(sel); }
  function el(tag, attrs, children) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === 'class') e.className = attrs[k];
      else if (k === 'text') e.textContent = attrs[k];
      else if (k === 'html') e.innerHTML = attrs[k];
      else e.setAttribute(k, attrs[k]);
    }
    if (children) children.forEach(c => c && e.appendChild(c));
    return e;
  }
  function setStatus(msg) { $('#status-msg').textContent = msg || ''; }

  function excelSerialToDateStr(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') {
      const ms = (v - 25569) * 86400 * 1000;
      const d = new Date(ms);
      if (isNaN(d.getTime())) return String(v);
      const yyyy = d.getUTCFullYear();
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(d.getUTCDate()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}`;
    }
    return String(v);
  }

  function ensureOffice(callback) {
    if (typeof Office === 'undefined') {
      state.inOffice = false;
      $('#env-label').textContent = 'モード: ブラウザ単体(デモ)';
      callback();
      return;
    }
    Office.onReady(info => {
      state.inOffice = (info && info.host === Office.HostType.Excel);
      $('#env-label').textContent = state.inOffice ? 'モード: Excelアドイン' : 'モード: ブラウザ単体(デモ)';
      callback();
    });
  }

  async function loadFromExcel() {
    if (!state.inOffice) {
      state.bugs = demoData();
      state.presetTags = ['UI', 'RPA', '通信', '電源', '設定', '認証', 'データ', 'パフォーマンス']; // デモ用プリセット
      // デモ用設定
      STATUS_DISPLAY_NAMES = {
        '新規': '新規',
        '解析': '解析',
        '修正': '修正',
        '確認': '確認',
        '再発': '再発',
        '完了': '完了'
      };
      // E3セルから担当者リストと登録者リストを取得（デモモード）
      const memberList = ['政次','高橋','伊藤','松田'];
      ASSIGNEE_ORDER = ['(未割当)', ...memberList];
      REPORTER_LIST = [...memberList];
      return;
    }
    setStatus('読み込み中...');
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      
      // 各種設定をセルから取得
      const configCells = sheet.getRangeByIndexes(SAMPLE_ROW - 1, 2, 1, 23); // C3:Y3
      configCells.load(['values']);
      await ctx.sync();
      
      const configValues = configCells.values[0];
      
      // C3セル：状態別表記設定（「元の状態:表示名/元の状態:表示名」形式）
      const statusDisplayConfig = configValues[0]; // C3 (C列は2番目なので0ベース)
      if (statusDisplayConfig && typeof statusDisplayConfig === 'string') {
        STATUS_DISPLAY_NAMES = {};
        statusDisplayConfig.split('/').forEach(pair => {
          const [original, display] = pair.split(':').map(s => s.trim());
          if (original && display) {
            STATUS_DISPLAY_NAMES[original] = display;
          }
        });
      } else {
        STATUS_DISPLAY_NAMES = {
          '新規': '新規',
          '解析': '解析',
          '修正': '修正',
          '確認': '確認',
          '再発': '再発',
          '完了': '完了'
        };
      }
      
      // E3セル：担当者・登録者リスト（/区切り）
      const memberConfig = configValues[2]; // E3 (E列は4番目なので2ベース)
      if (memberConfig && typeof memberConfig === 'string') {
        const memberList = memberConfig.split('/').map(s => s.trim()).filter(s => s);
        ASSIGNEE_ORDER = ['(未割当)', ...memberList];
        REPORTER_LIST = [...memberList];
      } else {
        const defaultMembers = ['政次','高橋','伊藤','松田'];
        ASSIGNEE_ORDER = ['(未割当)', ...defaultMembers];
        REPORTER_LIST = [...defaultMembers];
      }
      
      // D3セル：割り当て表記設定（廃止予定 - E3セルを使用）
      // const assigneeConfig = configValues[1];
      
      // G3セル：登録者リスト（廃止予定 - E3セルを使用）
      // const reporterConfig = configValues[4];
      
      // Y3セル：プリセットタグ
      const presetTagValue = configValues[23]; // Z3 (Z列は25番目なので23ベース)
      if (presetTagValue && typeof presetTagValue === 'string') {
        state.presetTags = presetTagValue.split('/').map(t => t.trim()).filter(t => t);
      } else {
        state.presetTags = ['UI', 'RPA', '通信', '電源', '設定', '認証', 'データ', 'パフォーマンス']; // デフォルト
      }
      

      
      const used = sheet.getUsedRange(true);
      used.load(['rowCount', 'columnCount']);
      await ctx.sync();

      const rowCount = used.rowCount || 0;
      if (rowCount < DATA_START) { state.bugs = []; return; }

      const dataRange = sheet.getRangeByIndexes(
        DATA_START - 1, 0, rowCount - (DATA_START - 1), COL_COUNT
      );
      dataRange.load(['values', 'numberFormat']);
      await ctx.sync();

      const values = dataRange.values;
      const bugs = [];
      const columns = getColumns(); // 一度取得して再利用
      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        if (row.every(v => v === '' || v === null)) continue;
        const obj = { rowIndex: DATA_START + r };
        for (let c = 0; c < COL_COUNT; c++) {
          const colDef = columns[c];
          let v = row[c];
          if (colDef.type === 'date') v = excelSerialToDateStr(v);
          if (v === null || v === undefined) v = '';
          v = String(v);
          obj[colDef.key] = v;
        }
        if (r === 130) {
          console.log('Excel row[130]:', row);
          console.log('Parsed bug obj[130]:', obj);
        }
        bugs.push(obj);
      }
      state.bugs = bugs;
    });
    setStatus('');
  }

  async function saveBugToExcel(bug) {
    if (!state.inOffice) {
      const idx = state.bugs.findIndex(b => b.rowIndex === bug.rowIndex);
      if (idx >= 0) state.bugs[idx] = bug;
      return;
    }
    setStatus('保存中...');
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const rowIndex0 = bug.rowIndex - 1;
      const writeRange = sheet.getRangeByIndexes(rowIndex0, 1, 1, COL_COUNT - 1);
      const columns = getColumns(); // 一度取得して再利用
      const rowVals = [];
      for (let c = 1; c < COL_COUNT; c++) {
        const colDef = columns[c];
        let v = bug[colDef.key];
        if (v === undefined || v === null) v = '';
        rowVals.push(v);
      }
      writeRange.values = [rowVals];
      
      // 文字列の折り返しを無効にする
      writeRange.format.wrapText = false;
      
      const today = new Date();
      const tStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
      sheet.getRangeByIndexes(rowIndex0, 3, 1, 1).values = [[tStr]];
      bug.updated = tStr;
      await ctx.sync();
    });
    setStatus('保存しました');
    setTimeout(() => setStatus(''), 2000);
  }

  function applyFilters(bugs) {
    const f = state.filters;
    return bugs.filter(b => {
      if (f.priority && b.priority !== f.priority) return false;
      if (f.status && b.status !== f.status) return false;
      if (f.text) {
        const t = f.text.toLowerCase();
        const hay = [b.title, b.assignee, b.tag, b.reporter].join(' ').toLowerCase();
        if (!hay.includes(t)) return false;
      }
      return true;
    });
  }
  function sortByPriority(bugs) {
    return bugs.slice().sort((a, b) => {
      const pa = PRIORITY_RANK[a.priority || ''] ?? 9;
      const pb = PRIORITY_RANK[b.priority || ''] ?? 9;
      if (pa !== pb) return pa - pb;
      return (a.rowIndex || 0) - (b.rowIndex || 0);
    });
  }



  function renderKanbanAssignee() {
    const board = $('#kanban-board-assignee');
    board.innerHTML = '';
    const order = ASSIGNEE_ORDER;

    // 担当者別表示では完了状態を除外
    const filteredBugs = applyFilters(state.bugs).filter(b => b.status !== '完了');
    const bugs = sortByPriority(filteredBugs);
    const groups = new Map();
    order.forEach(k => groups.set(k, []));
    bugs.forEach(b => {
      let key;
      // 新規の場合は必ず未割当レーンに表示
      if (b.status === '新規') {
        key = '(未割当)';
      } else {
        key = b.assignee || '(未割当)';
      }
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(b);
    });

    groups.forEach((items, key) => {
      const col = el('div', { class: 'kanban-col' });
      col.dataset.group = key;
      const header = el('div', { class: 'kanban-col-header' }, [
        el('span', { text: key || '(未設定)' }),
        el('span', { class: 'count', text: String(items.length) })
      ]);
      const body = el('div', { class: 'kanban-col-body' });
      items.forEach(b => body.appendChild(renderCard(b, true))); // drag enabled
      col.appendChild(header);
      col.appendChild(body);
      
      // ドロップイベントを追加（担当者別表示用）
      body.addEventListener('dragover', handleDragOver);
      body.addEventListener('dragleave', handleDragLeave);
      body.addEventListener('drop', handleDrop);
      
      board.appendChild(col);
    });

    $('#row-count').textContent = `${bugs.length} 件 / 全 ${state.bugs.length} 件`;
  }

  function renderKanbanStatus() {
    const board = $('#kanban-board-status');
    board.innerHTML = '';
    const order = STATUS_ORDER;

    const bugs = sortByPriority(applyFilters(state.bugs));
    const groups = new Map();
    order.forEach(k => groups.set(k, []));
    bugs.forEach(b => {
      let key = b.status || '';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(b);
    });

    groups.forEach((items, key) => {
      const col = el('div', { class: 'kanban-col' });
      col.dataset.group = key;
      const displayName = STATUS_DISPLAY_NAMES[key] || key || '(未設定)';
      const header = el('div', { class: 'kanban-col-header' }, [
        el('span', { text: displayName }),
        el('span', { class: 'count', text: String(items.length) })
      ]);
      const body = el('div', { class: 'kanban-col-body' });
      items.forEach(b => body.appendChild(renderCard(b, false))); // drag disabled
      col.appendChild(header);
      col.appendChild(body);
      
      board.appendChild(col);
    });

    $('#row-count').textContent = `${bugs.length} 件 / 全 ${state.bugs.length} 件`;
  }
  
  // ドラッグ&ドロップ機能（担当者別表示用）
  let draggedCard = null;
  
  function handleDragStart(e) {
    draggedCard = e.target;
    e.target.classList.add('dragging');
  }
  
  function handleDragOver(e) {
    e.preventDefault(); // ドロップを許可
    e.currentTarget.classList.add('drag-over');
  }
  
  function handleDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
  }
  
  async function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    
    if (!draggedCard) return;
    
    const targetCol = e.currentTarget.parentElement;
    const newAssignee = targetCol.dataset.group;
    const rowIndex = parseInt(draggedCard.dataset.row);
    const bug = state.bugs.find(b => b.rowIndex === rowIndex);
    
    if (!bug) return;
    
    // 実際の表示レーンを基準にoldAssigneeを判定
    let oldAssignee;
    if (bug.status === '新規') {
      oldAssignee = '(未割当)';  // 新規は必ず未割当レーンに表示されるため
    } else {
      oldAssignee = bug.assignee || '(未割当)';
    }
    
    // 担当者から未割当への移動を禁止
    if (oldAssignee !== '(未割当)' && newAssignee === '(未割当)') {
      setStatus('担当者から未割当には移動できません');
      setTimeout(() => setStatus(''), 3000);
      cleanupDrag();
      return;
    }
    
    // 同じ担当者への移動は何もしない
    if (oldAssignee === newAssignee) {
      cleanupDrag();
      return;
    }
    
    // 担当者を更新
    const newAssigneeValue = newAssignee === '(未割当)' ? '' : newAssignee;
    bug.assignee = newAssigneeValue;
    
    // 状態に応じて適切な担当者フィールドも更新
    let statusMessage = '';
    if (newAssignee !== '(未割当)') {
      switch(bug.status) {
        case '解析':
          bug.analyst = newAssignee;
          statusMessage = `担当者を「${newAssignee}」に変更し、解析者も更新しました`;
          break;
        case '修正':
          bug.fixer = newAssignee;
          statusMessage = `担当者を「${newAssignee}」に変更し、対応者も更新しました`;
          break;
        case '確認':
          bug.verifier = newAssignee;
          statusMessage = `担当者を「${newAssignee}」に変更し、確認者も更新しました`;
          break;
        default:
          statusMessage = `担当者を「${newAssignee}」に変更しました`;
          break;
      }
    }
    
    // 未割当から担当者への移動時は状態を「解析」に変更し、解析者を設定
    if (oldAssignee === '(未割当)' && newAssignee !== '(未割当)') {
      bug.status = '解析';
      bug.analyst = newAssignee; // 解析者に担当者を設定
      statusMessage = `担当者を「${newAssignee}」に変更し、状況を「解析」に変更しました`;
    }
    
    setStatus(statusMessage || `担当者を「${newAssignee}」に変更しました`);
    
    setTimeout(() => setStatus(''), 3000);
    
    // Excel保存
    try {
      // state.bugsも確実に更新
      const bugIndex = state.bugs.findIndex(b => b.rowIndex === bug.rowIndex);
      if (bugIndex >= 0) {
        state.bugs[bugIndex] = bug;
      }
      
      await saveBugToExcel(bug);
    } catch (e) {
      console.error('保存エラー:', e);
      setStatus('保存に失敗しました');
    }
    
    cleanupDrag();
    // 表示更新
    render();
  }
  
  function cleanupDrag() {
    if (draggedCard) {
      draggedCard.classList.remove('dragging');
      draggedCard = null;
    }
    // 全てのdrag-overクラスを削除
    document.querySelectorAll('.drag-over').forEach(el => {
      el.classList.remove('drag-over');
    });
  }
  
  function renderCard(b, dragEnabled = false) {
    const card = el('div', { class: 'kanban-card pri-' + (b.priority || '') });
    card.dataset.row = b.rowIndex;
    card.draggable = dragEnabled; // ドラッグ可否をパラメータで制御
    
    // 差し戻し状態の場合はカードにスタイルを追加
    if (b.reject === '○' && b.status === '修正') {
      card.style.borderLeft = '4px solid #f44336';
      card.style.backgroundColor = '#fff3f3';
    }
    
    // 1行目：左端にID、発生起因。右端に優先度と差し戻しマーク
    const row1 = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;' });
    
    // 発生起因と起因番号を組み合わせて表示
    const originText = (() => {
      if (!b.origin) return '';
      if (!b.originNumber) return b.origin;
      return `${b.origin}-${b.originNumber}`;
    })();
    
    const leftPart1 = el('div', { style: 'display:flex;gap:8px;' }, [
      el('span', { class: 'id', text: `#${b.id || ''}`, style: 'font-weight:bold;font-size:12px;' }),
      el('span', { text: originText, style: 'font-size:11px;color:#666;' })
    ]);
    const rightPart1 = el('div', { style: 'display:flex;gap:4px;align-items:center;' });
    
    // 差し戻しマーク
    if (b.reject === '○' && b.status === '修正') {
      rightPart1.appendChild(el('span', { 
        text: '⚠ 差し戻し', 
        style: 'background:#f44336;color:white;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:bold;' 
      }));
    }
    
    // 優先度バッジ
    if (b.priority) {
      rightPart1.appendChild(el('span', { class: `badge pri-${b.priority}`, text: b.priority, style: 'font-size:11px;' }));
    }
    row1.appendChild(leftPart1);
    row1.appendChild(rightPart1);
    
    // 2行目：タイトル
    const row2 = el('div', { 
      class: 'title', 
      text: b.title || '(無題)', 
      style: 'margin-bottom:4px;font-weight:500;line-height:1.3;' 
    });
    
    // 3行目：左端に状態、右端に名前
    const row3 = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;' });
    const leftPart3 = el('div', {});
    if (b.status) {
      leftPart3.appendChild(el('span', { class: `badge st-${b.status}`, text: STATUS_DISPLAY_NAMES[b.status] || b.status, style: 'font-size:11px;' }));
    }
    
    const rightPart3 = el('div', {});
    // 状態により表示する名前と日付を変更
    let nameText = '';
    let nameStyle = 'font-size:11px;';
    
    switch(b.status) {
      case '新規':
        nameText = '(未割当)';
        nameStyle += 'color:#999;';
        break;
      case '解析待ち':
      case '解析':
        nameText = b.analyst || '(未設定)';
        if (!b.analyst) nameStyle += 'color:#999;';
        // 解析日を表示（解析者と有効な解析日がある場合のみ）
        if (b.analyst && b.analysisDate && b.analysisDate.trim() !== '') {
          const date = new Date(b.analysisDate);
          if (!isNaN(date.getTime())) {
            const month = date.getMonth() + 1;
            const day = date.getDate();
            nameText += ` (${month}/${day})`;
          }
        }
        break;
      case '修正待ち':
      case '修正':
        nameText = b.fixer || '(未設定)';
        if (!b.fixer) nameStyle += 'color:#999;';
        // 対応日を表示（対応者と有効な対応日がある場合のみ）
        if (b.fixer && b.fixDate && b.fixDate.trim() !== '') {
          const date = new Date(b.fixDate);
          if (!isNaN(date.getTime())) {
            const month = date.getMonth() + 1;
            const day = date.getDate();
            nameText += ` (${month}/${day})`;
          }
        }
        break;
      case '確認待ち':
      case '確認':
        nameText = b.verifier || '(未設定)';
        if (!b.verifier) nameStyle += 'color:#999;';
        // 確認日を表示（確認者と有効な確認日がある場合のみ）
        if (b.verifier && b.verifyDate && b.verifyDate.trim() !== '') {
          const date = new Date(b.verifyDate);
          if (!isNaN(date.getTime())) {
            const month = date.getMonth() + 1;
            const day = date.getDate();
            nameText += ` (${month}/${day})`;
          }
        }
        break;
      default:
        if (b.assignee) {
          nameText = b.assignee;
        } else {
          nameText = '(未割当)';
          nameStyle += 'color:#999;';
        }
    }
    
    rightPart3.appendChild(el('span', { text: nameText, style: nameStyle }));
    row3.appendChild(leftPart3);
    row3.appendChild(rightPart3);
    
    card.appendChild(row1);
    card.appendChild(row2);
    card.appendChild(row3);
    
    card.addEventListener('click', () => openModal(b.rowIndex));
    
    // ドラッグ&ドロップイベントを必要時のみ追加
    if (dragEnabled) {
      card.addEventListener('dragstart', handleDragStart);
      card.addEventListener('dragend', cleanupDrag);
    }
    return card;
  }

  function openModal(rowIndex) {
    const bug = state.bugs.find(b => b.rowIndex === rowIndex);
    if (!bug) return;
    state.editingRow = rowIndex;
    $('#modal-title').textContent = `バグ詳細  #${bug.id || ''}  ${bug.title || ''}`;

    const body = $('#modal-body');
    body.innerHTML = '';

    // タブUI
    const tabNames = [
      { key: 'jisho', label: '事象' },
      { key: 'kaiseki', label: '解析' },
      { key: 'shochi', label: '処置' },
      { key: 'kekka', label: '結果確認' },
      { key: 'kanri', label: '管理' }
    ];
    // 状況に応じた初期タブ
    const statusTabMap = {
      '新規': 'jisho',
      '解析': 'kaiseki',
      '修正': 'shochi',
      '確認': 'kekka',
      '再発': 'kekka',
      '完了': 'jisho'
    };
    let activeTab = statusTabMap[bug.status] || 'jisho';

    // 常時表示エリア
    const alwaysArea = el('div', { class: 'always-area', style: 'margin-bottom:12px;padding:8px 0;border-bottom:1px solid #ccc;' });
    
    // 発生起因のみ表示
    const originDisplay = (() => {
      if (!bug.origin) return '';
      if (!bug.originNumber) return bug.origin;
      return `${bug.origin}-${bug.originNumber}`;
    })();
    
    const row1 = el('div', { style: 'display:flex;gap:16px;flex-wrap:wrap;margin-bottom:8px;' }, [
      el('div', {}, [el('b', { text: '発生起因: ' }), el('span', { text: originDisplay })])
    ]);
    alwaysArea.appendChild(row1);
    
    // 日付をm/d形式に変換する関数
    function formatToMD(dateStr) {
      if (!dateStr) return '';
      if (dateStr.includes('/')) return dateStr; // 既にm/d形式
      if (dateStr.includes('-')) {
        // yyyy-mm-dd形式からm/d形式に変換
        const parts = dateStr.split('-');
        if (parts.length === 3) {
          const month = parseInt(parts[1]);
          const day = parseInt(parts[2]);
          return `${month}/${day}`;
        }
      }
      return dateStr;
    }
    
    // ワークフローテーブル（流れるような表現）
    const workflowContainer = el('div', { style: 'width:100%;' });
    
    // ワークフローステップの定義
    const workflowSteps = [
      { key: 'new', label: '新規', person: bug.reporter, date: bug.occurredOn, status: '新規' },
      { key: 'analysis', label: '解析', person: bug.analyst, date: bug.analysisDate, status: '解析' },
      { key: 'fix', label: '修正', person: bug.fixer, date: bug.fixDate, status: '修正' },
      { key: 'verify', label: '確認', person: bug.verifier, date: bug.verifyDate, status: '確認' }
    ];
    
    // 現在のステータスのインデックスを取得
    const currentStatusIndex = workflowSteps.findIndex(step => step.status === bug.status);
    const completedStatusIndex = currentStatusIndex === -1 && bug.status === '完了' ? workflowSteps.length : currentStatusIndex;
    
    // フレックスコンテナ
    const flowContainer = el('div', { 
      style: 'display:flex;align-items:center;justify-content:space-between;padding:8px;' 
    });
    
    workflowSteps.forEach((step, index) => {
      // ステップの状態を判定
      let stepState = 'pending'; // 未完了
      if (index < completedStatusIndex || (bug.status === '完了' && index < workflowSteps.length)) {
        stepState = 'completed'; // 完了
      } else if (index === completedStatusIndex) {
        stepState = 'current'; // 現在
      }
      
      // ステップの色を決定
      let bgColor = '#f5f5f5';
      let textColor = '#999';
      let borderColor = '#ddd';
      
      if (stepState === 'completed') {
        bgColor = '#e8f5e8';
        textColor = '#2e7d2e';
        borderColor = '#4caf50';
      } else if (stepState === 'current') {
        bgColor = '#e3f2fd';
        textColor = '#1976d2';
        borderColor = '#2196f3';
      }
      
      // ステップボックス
      const stepBox = el('div', {
        style: `
          flex: 1;
          border: 2px solid ${borderColor};
          background: ${bgColor};
          border-radius: 8px;
          padding: 8px;
          text-align: center;
          position: relative;
          margin: 0 4px;
          transition: all 0.3s ease;
        `
      });
      
      // ステップラベル
      stepBox.appendChild(el('div', {
        text: step.label,
        style: `font-weight:bold;font-size:12px;color:${textColor};margin-bottom:4px;`
      }));
      
      // 担当者と日付
      const personText = step.person || '(未設定)';
      const dateText = formatToMD(step.date);
      const displayText = (dateText && dateText !== '') ? `${personText} (${dateText})` : personText;
      
      stepBox.appendChild(el('div', {
        text: stepState === 'pending' ? '-' : displayText,
        style: `font-size:11px;color:${stepState === 'pending' ? '#ccc' : textColor};`
      }));
      
      flowContainer.appendChild(stepBox);
      
      // 矢印を追加（最後のステップ以外）
      if (index < workflowSteps.length - 1) {
        const arrow = el('div', {
          text: '→',
          style: `
            font-size: 18px;
            color: ${index < completedStatusIndex ? '#4caf50' : '#ddd'};
            margin: 0 8px;
            font-weight: bold;
          `
        });
        flowContainer.appendChild(arrow);
      }
    });
    
    workflowContainer.appendChild(flowContainer);
    alwaysArea.appendChild(workflowContainer);
    
    body.appendChild(alwaysArea);

    const tabHeader = el('div', { class: 'tab-header' },
      tabNames.map(tab => {
        const btn = el('button', {
          class: 'tab-btn' + (activeTab === tab.key ? ' active' : ''),
          type: 'button'
        }, [el('span', { text: tab.label })]);
        btn.dataset.tab = tab.key;
        btn.addEventListener('click', () => {
          body.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          renderTab(tab.key);
        });
        return btn;
      })
    );
    body.appendChild(tabHeader);

    // タブ内容表示部
    const tabContent = el('div', { class: 'tab-content' });
    body.appendChild(tabContent);

    function renderTab(tabKey) {
      tabContent.innerHTML = '';
      
      // 状況に応じた入力制御フラグ
      const isDisabled = {
        kaiseki: bug.status === '新規' || bug.status === '確認',
        shochi: bug.status === '新規' || bug.status === '解析' || bug.status === '確認',
        kekka: bug.status === '新規' || bug.status === '解析' || bug.status === '修正'
      };
      
      if (tabKey === 'jisho') {
        // 事象タブ：タイトル、再現手順、期待する動作、実際の動作（編集可）
        console.log('steps:', bug.steps, '| expected:', bug.expected, '| actual:', bug.actual);
        tabContent.appendChild(el('div', {}, [
          el('label', { text: 'タイトル' }), el('br'),
          el('input', { 
            type: 'text', 
            style: 'width:98%;margin-bottom:16px;', 
            value: bug.title || '', 
            'data-key': 'title',
            placeholder: 'バグのタイトルを入力してください' 
          })
        ]));
        tabContent.appendChild(el('div', {}, [
          el('label', { text: '再現手順' }), el('br'),
          (() => {
            const ta = el('textarea', {
              rows: 10,
              style: 'width:98%;',
              placeholder: '再現手順を入力してください',
              'data-key': 'steps'
            });
            ta.textContent = (bug.steps !== undefined && bug.steps !== null) ? bug.steps : '';
            return ta;
          })()
        ]));
        tabContent.appendChild(el('div', {}, [
          el('label', { text: '期待する動作' }), el('br'),
          (() => {
            const ta = el('textarea', {
              rows: 2,
              style: 'width:98%;',
              placeholder: '期待する動作を入力してください',
              'data-key': 'expected'
            });
            ta.textContent = (bug.expected !== undefined && bug.expected !== null) ? bug.expected : '';
            return ta;
          })()
        ]));
        tabContent.appendChild(el('div', {}, [
          el('label', { text: '実際の動作' }), el('br'),
          (() => {
            const ta = el('textarea', {
              rows: 2,
              style: 'width:98%;',
              placeholder: '実際の動作を入力してください',
              'data-key': 'actual'
            });
            ta.textContent = (bug.actual !== undefined && bug.actual !== null) ? bug.actual : '';
            return ta;
          })()
        ]));
      } else if (tabKey === 'kaiseki') {
        // 解析タブ：原因（編集可）、解析完了チェック
        tabContent.appendChild(el('div', {}, [
          el('label', { text: '原因' }), el('br'),
          (() => {
            const ta = el('textarea', {
              rows: 15,
              style: 'width:98%;',
              placeholder: '原因を入力してください',
              'data-key': 'cause'
            });
            ta.textContent = (bug.cause !== undefined && bug.cause !== null) ? bug.cause : '';
            ta.disabled = isDisabled.kaiseki; // 入力制御
            return ta;
          })()
        ]));
        tabContent.appendChild(el('div', { style: 'margin-top:8px;' }, [
          el('label', {}, [
            (() => {
              const checkbox = el('input', { type: 'checkbox', 'data-key': 'kaisekikanryo' });
              // 解析日が入っていたらチェック
              if (bug.analysisDate && bug.analysisDate.trim() !== '') {
                checkbox.checked = true;
                checkbox.disabled = true; // 解析日が入っている場合は読み取り専用
              }
              // 対応日や確認日が入っている場合も読み取り専用
              if ((bug.fixDate && bug.fixDate.trim() !== '') || (bug.verifyDate && bug.verifyDate.trim() !== '')) {
                checkbox.disabled = true;
              }
              // 状況に応じた入力制御も適用
              if (isDisabled.kaiseki) {
                checkbox.disabled = true;
              }
              return checkbox;
            })(),
            el('span', { text: '解析完了（修正に変更）' })
          ])
        ]));
      } else if (tabKey === 'shochi') {
        // 処置タブ：影響範囲（チェックボックス表）と処置内容を横並び、修正Ver（編集可）、処置完了チェック
        
        // メインコンテナ（横並び）
        const mainContainer = el('div', { style: 'display:flex;gap:16px;margin-bottom:16px;' });
        
        // 左側：影響範囲のチェックボックス表
        const leftPanel = el('div', { style: 'flex:0 0 auto;min-width:fit-content;' });
        
        const scopeOptions = ['定義(通常)', '定義(電源断)', '定義(通信断)', 'RPA', 'アプリ'];
        const currentScope = bug.scope || '';
        
        // デバッグ情報を出力
        console.log('=== 修正完了抽出デバッグ ===');
        console.log('currentScope:', JSON.stringify(currentScope));
        
        // 影響範囲から（済）を除去して修正対象を抽出
        const selectedScopes = currentScope.split('/').map(s => s.trim().replace(/（済）$/, '')).filter(s => s);
        console.log('selectedScopes:', selectedScopes);
        
        // 修正完了状況を取得（scopeCompletedフィールド + 影響範囲の（済）から抽出）
        const currentCompleted = bug.scopeCompleted || '';
        let completedScopes = currentCompleted.split('/').map(s => s.trim()).filter(s => s);
        console.log('currentCompleted:', JSON.stringify(currentCompleted));
        console.log('completedScopes from field:', completedScopes);
        
        // 影響範囲から（済）が付いているアイテムも修正完了に含める
        const scopeItems = currentScope.split('/').map(s => s.trim()).filter(s => s);
        console.log('scopeItems:', scopeItems);
        
        const scopeWithCompleted = scopeItems.filter(s => s.includes('（済）'));
        console.log('scopeWithCompleted:', scopeWithCompleted);
        
        const completedFromScope = scopeWithCompleted.map(s => s.replace(/（済）$/, ''));
        console.log('completedFromScope:', completedFromScope);
        
        completedScopes = [...new Set([...completedScopes, ...completedFromScope])];
        console.log('final completedScopes:', completedScopes);
        console.log('=== デバッグ終了 ===');
        
        leftPanel.appendChild(el('div', { style: 'margin-bottom:16px;' }, [
          el('label', { text: '影響範囲', style: 'font-weight:bold;margin-bottom:8px;display:block;' }),
          (() => {
            const table = el('table', { style: 'border-collapse:collapse;width:100%;' });
            
            // ヘッダー行
            const headerRow = el('tr');
            headerRow.appendChild(el('th', { 
              text: '修正対象', 
              style: 'border:1px solid #ddd;padding:8px;background:#f5f5f5;text-align:left;white-space:nowrap;min-width:120px;' 
            }));
            headerRow.appendChild(el('th', { 
              text: '修正完了', 
              style: 'border:1px solid #ddd;padding:8px;background:#f5f5f5;text-align:center;white-space:nowrap;width:80px;' 
            }));
            table.appendChild(headerRow);
            
            // 各オプションの行
            scopeOptions.forEach(option => {
              const row = el('tr');
              
              // 修正対象列
              const targetCell = el('td', { style: 'border:1px solid #ddd;padding:6px;white-space:nowrap;' });
              const targetCheckbox = el('input', { 
                type: 'checkbox', 
                value: option,
                'data-scope-option': option,
                style: 'margin-right:8px;'
              });
              
              if (selectedScopes.includes(option)) {
                targetCheckbox.checked = true;
              }
              
              targetCheckbox.disabled = isDisabled.shochi; // 入力制御
              
              const targetLabel = el('label', { style: 'display:flex;align-items:center;cursor:pointer;white-space:nowrap;' }, [
                targetCheckbox,
                el('span', { text: option })
              ]);
              
              targetCell.appendChild(targetLabel);
              row.appendChild(targetCell);
              
              // 修正完了列
              const completedCell = el('td', { style: 'border:1px solid #ddd;padding:6px;text-align:center;white-space:nowrap;' });
              const completedCheckbox = el('input', { 
                type: 'checkbox', 
                value: option,
                'data-scope-completed': option,
                style: 'display:none;' // 初期は非表示
              });
              
              if (completedScopes.includes(option)) {
                completedCheckbox.checked = true;
              }
              
              completedCheckbox.disabled = isDisabled.shochi; // 入力制御
              
              // 修正対象がチェックされている場合は完了チェックボックスを表示
              if (selectedScopes.includes(option)) {
                completedCheckbox.style.display = 'inline-block';
              }
              
              // 修正対象チェックボックスの変更イベント
              targetCheckbox.addEventListener('change', (e) => {
                if (e.target.checked) {
                  completedCheckbox.style.display = 'inline-block';
                } else {
                  completedCheckbox.style.display = 'none';
                  completedCheckbox.checked = false;
                }
              });
              
              completedCell.appendChild(completedCheckbox);
              row.appendChild(completedCell);
              table.appendChild(row);
            });
            
            return table;
          })()
        ]));
        
        mainContainer.appendChild(leftPanel);
        
        // 右側：処置内容
        const rightPanel = el('div', { style: 'flex:1;margin-left:16px;' });
        rightPanel.appendChild(el('div', {}, [
          el('label', { text: '処置内容', style: 'font-weight:bold;margin-bottom:8px;display:block;' }),
          (() => {
            const ta = el('textarea', {
              rows: 15,
              style: 'width:98%;',
              placeholder: '処置内容を入力してください',
              'data-key': 'fix'
            });
            ta.textContent = (bug.fix !== undefined && bug.fix !== null) ? bug.fix : '';
            ta.disabled = isDisabled.shochi; // 入力制御
            return ta;
          })()
        ]));
        
        mainContainer.appendChild(rightPanel);
        tabContent.appendChild(mainContainer);
        
        // 下部：修正Ver、処置完了チェック
        tabContent.appendChild(el('div', {}, [el('label', { text: '修正Ver' }), el('br'),
          (() => {
            const input = el('input', { type: 'text', style: 'width:98%;', value: bug.fixVer || '', 'data-key': 'fixVer' });
            input.disabled = isDisabled.shochi; // 入力制御
            return input;
          })()]));
        tabContent.appendChild(el('div', { style: 'margin-top:8px;' }, [
          el('label', {}, [
            (() => {
              const checkbox = el('input', { type: 'checkbox', 'data-key': 'shochikanryo' });
              // 対応日が入っていたらチェック
              if (bug.fixDate && bug.fixDate.trim() !== '') {
                checkbox.checked = true;
                checkbox.disabled = true; // 対応日が入っている場合は読み取り専用
              }
              // 確認日が入っている場合も読み取り専用
              if (bug.verifyDate && bug.verifyDate.trim() !== '') {
                checkbox.disabled = true;
              }
              // 状況に応じた入力制御も適用
              if (isDisabled.shochi) {
                checkbox.disabled = true;
              }
              return checkbox;
            })(),
            el('span', { text: '処置完了（確認に変更）' })
          ])
        ]));
      } else if (tabKey === 'kekka') {
        // 結果確認タブ：確認内容（編集可）、確認完了・差し戻しラジオボタン
        tabContent.appendChild(el('div', {}, [
          el('label', { text: '確認内容' }), el('br'),
          (() => {
            const ta = el('textarea', {
              rows: 10,
              style: 'width:98%;',
              placeholder: '確認内容を入力してください',
              'data-key': 'verify'
            });
            ta.textContent = (bug.verify !== undefined && bug.verify !== null) ? bug.verify : '';
            ta.disabled = isDisabled.kekka; // 入力制御
            return ta;
          })()
        ]));
        
        // ラジオボタングループ
        const radioGroup = el('div', { style: 'margin-top:8px;' });
        const groupName = `result_${bug.rowIndex || 'new'}`;
        
        // 確認完了ラジオボタン
        const completeRadio = el('label', { style: 'margin-right:16px;' }, [
          (() => {
            const radio = el('input', { 
              type: 'radio', 
              name: groupName, 
              value: 'complete', 
              'data-key': 'kekkakanryo' 
            });
            // 確認日が入っていたらチェック
            if (bug.verifyDate && bug.verifyDate.trim() !== '') {
              radio.checked = true;
              radio.disabled = true; // 確認日が入っている場合は読み取り専用
            }
            // 状況に応じた入力制御も適用
            if (isDisabled.kekka) {
              radio.disabled = true;
            }
            return radio;
          })(),
          el('span', { text: '確認完了（完了に変更）' })
        ]);
        
        // 差し戻しラジオボタン
        const rejectRadio = el('label', {}, [
          (() => {
            const radio = el('input', { 
              type: 'radio', 
              name: groupName, 
              value: 'reject', 
              'data-key': 'sashimodoshi' 
            });
            // 差し戻しに○が入っていて、かつ確認日がない場合のみチェック
            if (bug.reject === '○' && (!bug.verifyDate || bug.verifyDate.trim() === '')) {
              radio.checked = true;
            }
            // 確認日が入っている場合は読み取り専用
            if (bug.verifyDate && bug.verifyDate.trim() !== '') {
              radio.disabled = true;
            }
            // 状況に応じた入力制御も適用
            if (isDisabled.kekka) {
              radio.disabled = true;
            }
            return radio;
          })(),
          el('span', { text: '差し戻し（修正に変更）' })
        ]);
        
        radioGroup.appendChild(completeRadio);
        radioGroup.appendChild(rejectRadio);
        tabContent.appendChild(radioGroup);
      } else if (tabKey === 'kanri') {
        // 管理タブ：優先度、影響度、タグ（編集可）
        
        // 1行目：優先度と影響度（横並び）
        const row1 = el('div', { style: 'display:flex;gap:16px;margin-bottom:16px;' }, [
          (() => {
            const field = el('div', { style: 'flex:1;' });
            field.appendChild(el('label', { text: '優先度' }));
            field.appendChild(el('br'));
            const select = el('select', { style: 'width:100%;', 'data-key': 'priority' });
            ['','高','中','低'].forEach(option => {
              const op = el('option', { value: option, text: option || '(選択)' });
              if (option === (bug.priority || '中')) { // 初期値は中、既存の値があればそれを使用
                op.selected = true;
              }
              select.appendChild(op);
            });
            field.appendChild(select);
            return field;
          })(),
          (() => {
            const field = el('div', { style: 'flex:1;' });
            field.appendChild(el('label', { text: '影響度' }));
            field.appendChild(el('br'));
            const select = el('select', { style: 'width:100%;', 'data-key': 'severity' });
            ['','致命的','重大','警備'].forEach(option => {
              const op = el('option', { value: option, text: option || '(選択)' });
              if (option === (bug.severity || '')) { // 初期値は空
                op.selected = true;
              }
              select.appendChild(op);
            });
            field.appendChild(select);
            return field;
          })()
        ]);
        
        // 2行目：タグ（チップ形式）
        const row2 = el('div', {}, [
          // タグヘッダー（ラベル + ＋ボタン）
          (() => {
            const headerContainer = el('div', { 
              style: 'display:flex;align-items:center;gap:8px;margin-bottom:8px;' 
            });
            
            headerContainer.appendChild(el('label', { 
              text: 'タグ', 
              style: 'font-weight:bold;' 
            }));
            
            const toggleBtn = el('button', {
              type: 'button',
              text: '＋',
              id: 'tag-toggle-btn',
              style: 'height:16px;width:16px;border:1px solid #007acc;background:#007acc;color:white;border-radius:8px;cursor:pointer;font-size:10px;line-height:1;display:flex;align-items:center;justify-content:center;transition:transform 0.3s ease;padding:0;'
            });
            
            toggleBtn.addEventListener('click', () => {
              const tagInputArea = document.querySelector('#tag-input-area');
              const btn = document.querySelector('#tag-toggle-btn');
              
              if (tagInputArea && btn) {
                const isVisible = tagInputArea.style.maxHeight !== '0px' && tagInputArea.style.maxHeight !== '';
                
                if (isVisible) {
                  // 閉じる
                  tagInputArea.style.maxHeight = '0px';
                  tagInputArea.style.opacity = '0';
                  btn.textContent = '＋';
                  btn.style.transform = 'rotate(0deg)';
                } else {
                  // 開く
                  tagInputArea.style.maxHeight = '300px';
                  tagInputArea.style.opacity = '1';
                  btn.textContent = '−';
                  btn.style.transform = 'rotate(180deg)';
                }
              }
            });
            
            headerContainer.appendChild(toggleBtn);
            return headerContainer;
          })(),
          
          // タグ入力エリア（スライド対応）
          (() => {
            const inputArea = el('div', {
              id: 'tag-input-area',
              style: 'max-height:0px;opacity:0;overflow:hidden;transition:max-height 0.3s ease, opacity 0.3s ease;margin-bottom:12px;'
            });
            
            // プリセットタグボタン群
            const presetContainer = el('div', { style: 'margin-bottom:12px;' });
            presetContainer.appendChild(el('div', { 
              text: 'プリセットタグ:', 
              style: 'font-size:12px;color:#666;margin-bottom:6px;' 
            }));
            
            const buttonContainer = el('div', { 
              style: 'display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;',
              id: 'preset-tags-container'
            });
            
            // 初期表示は後で行う
            
            presetContainer.appendChild(buttonContainer);
            inputArea.appendChild(presetContainer);
            
            // カスタムタグ入力
            const customContainer = el('div', { style: 'margin-bottom:12px;' });
            customContainer.appendChild(el('div', { 
              text: 'カスタムタグ:', 
              style: 'font-size:12px;color:#666;margin-bottom:6px;' 
            }));
            
            const inputContainer = el('div', { style: 'display:flex;gap:8px;' });
            const input = el('input', { 
              type: 'text', 
              placeholder: 'カスタムタグを入力',
              style: 'flex:1;padding:4px 8px;border:1px solid #ddd;border-radius:4px;',
              id: 'custom-tag-input'
            });
            
            const addBtn = el('button', {
              type: 'button',
              text: '追加',
              style: 'padding:4px 12px;border:1px solid #007acc;background:#007acc;color:white;border-radius:4px;cursor:pointer;'
            });
            
            addBtn.addEventListener('click', () => {
              const value = input.value.trim();
              if (value) {
                addTag(value);
                input.value = '';
              }
            });
            
            input.addEventListener('keydown', (e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                addBtn.click();
              }
            });
            
            inputContainer.appendChild(input);
            inputContainer.appendChild(addBtn);
            customContainer.appendChild(inputContainer);
            inputArea.appendChild(customContainer);
            
            return inputArea;
          })(),
          
          // 選択されたタグ表示エリア
          (() => {
            const tagContainer = el('div', { 
              style: 'border:1px solid #ddd;border-radius:4px;padding:8px;min-height:40px;background:#fafafa;',
              id: 'selected-tags-container'
            });
            tagContainer.appendChild(el('div', { 
              text: '選択されたタグ:', 
              style: 'font-size:12px;color:#666;margin-bottom:6px;' 
            }));
            
            const tagsDisplay = el('div', { 
              style: 'display:flex;flex-wrap:wrap;gap:6px;',
              id: 'tags-display'
            });
            
            tagContainer.appendChild(tagsDisplay);
            return tagContainer;
          })(),
          
          // 隠し入力フィールド（実際の値を保存）
          el('input', { 
            type: 'hidden', 
            'data-key': 'tag',
            id: 'tag-hidden-input',
            value: bug.tag || ''
          })
        ]);
        
        tabContent.appendChild(row1);
        tabContent.appendChild(row2);
        
        // プリセットボタン更新関数（グローバルスコープ）
        function renderPresetButtons() {
          const container = document.querySelector('#preset-tags-container');
          if (!container) return;
          
          container.innerHTML = '';
          const presetTags = state.presetTags || [];
          
          // 現在選択されているタグを取得
          const hiddenInput = document.querySelector('#tag-hidden-input');
          const currentTags = hiddenInput ? 
            (hiddenInput.value || '').split('/').map(t => t.trim()).filter(t => t) : [];
          
          presetTags.forEach(tag => {
            // 選択状態に応じて背景色を決定
            const isSelected = currentTags.includes(tag);
            const baseColor = isSelected ? '#007acc' : '#cccccc';
            const hoverColor = isSelected ? '#0056a3' : '#999999';
            
            const btn = el('button', {
              type: 'button',
              text: tag,
              style: `background:${baseColor};color:white;padding:4px 12px;border:none;border-radius:12px;cursor:pointer;font-size:12px;margin:2px;transition:background-color 0.2s ease;`
            });
            
            // ホバー効果
            btn.addEventListener('mouseenter', () => {
              btn.style.backgroundColor = hoverColor;
            });
            btn.addEventListener('mouseleave', () => {
              btn.style.backgroundColor = baseColor;
            });
            
            btn.addEventListener('click', () => addTag(tag));
            container.appendChild(btn);
          });
        }
        
        // タグ機能の初期化（DOM追加後に実行）
        setTimeout(() => {
          // プリセットボタンを初期表示
          renderPresetButtons();
          
          // 既存のタグを表示（/区切り）
          const currentTags = (bug.tag || '').split('/').map(t => t.trim()).filter(t => t);
          currentTags.forEach(tag => {
            if (tag.trim()) addTagChip(tag.trim());
          });
        }, 0);
        
        // プリセットタグをY3セルに保存する関数
        async function savePresetTags() {
          try {
            await Excel.run(async (ctx) => {
              const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
              const presetTagCell = sheet.getRangeByIndexes(SAMPLE_ROW - 1, 24, 1, 1); // Y列は24番目（0ベース）
              presetTagCell.values = [[state.presetTags.join('/')]]; // プリセットタグを/区切りで保存
              await ctx.sync();
            });
          } catch (error) {
            console.error('プリセットタグ保存エラー:', error);
          }
        }

        // タグ追加関数
        function addTag(tagName) {
          const tagsDisplay = document.querySelector('#tags-display');
          const hiddenInput = document.querySelector('#tag-hidden-input');
          
          if (!tagsDisplay || !hiddenInput) return;
          
          // 既存チェック（/区切り）
          const currentTags = (hiddenInput.value || '').split('/').map(t => t.trim()).filter(t => t);
          if (currentTags.includes(tagName.trim())) {
            return; // 既に存在する場合は追加しない
          }
          
          // プリセットタグにない場合は追加
          if (!state.presetTags.includes(tagName.trim())) {
            state.presetTags.push(tagName.trim());
            // プリセットボタンを再描画
            renderPresetButtons();
            // Y3セルに保存
            savePresetTags();
          }
          
          addTagChip(tagName.trim());
          updateHiddenInput();
        }
        
        // タグチップ追加関数
        function addTagChip(tagName) {
          const tagsDisplay = document.querySelector('#tags-display');
          if (!tagsDisplay) return;
          
          const chip = el('div', {
            style: 'display:inline-flex;align-items:center;background:#007acc;color:white;padding:2px 8px;border-radius:12px;font-size:12px;gap:4px;'
          });
          
          chip.appendChild(el('span', { text: tagName }));
          
          const removeBtn = el('button', {
            type: 'button',
            text: '×',
            style: 'background:none;border:none;color:white;cursor:pointer;font-size:14px;padding:0;margin-left:4px;'
          });
          
          removeBtn.addEventListener('click', () => {
            chip.remove();
            updateHiddenInput();
          });
          
          chip.appendChild(removeBtn);
          tagsDisplay.appendChild(chip);
        }
        
        // 隠し入力フィールド更新関数
        function updateHiddenInput() {
          const tagsDisplay = document.querySelector('#tags-display');
          const hiddenInput = document.querySelector('#tag-hidden-input');
          
          if (!tagsDisplay || !hiddenInput) return;
          
          const tags = [];
          tagsDisplay.querySelectorAll('div').forEach(chip => {
            const span = chip.querySelector('span');
            if (span) {
              tags.push(span.textContent.trim());
            }
          });
          
          hiddenInput.value = tags.join('/');
          // プリセットボタンの表示を更新（選択状態に応じて色分け）
          renderPresetButtons();
        }
      }
    }

    renderTab(activeTab);
    $('#modal').classList.remove('hidden');
  }

  function closeModal() {
    $('#modal').classList.add('hidden');
    state.editingRow = null;
  }

  function getNextId() {
    const maxId = Math.max(0, ...state.bugs.map(b => parseInt(b.id) || 0));
    return maxId + 1;
  }

  function getTodayString() {
    const today = new Date();
    return `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  }

  function openNewBugModal() {
    state.editingRow = 'new';
    $('#modal-title').textContent = '新規バグ登録';

    const body = $('#modal-body');
    body.innerHTML = '';

    // データ初期値
    const newBugData = {
      id: String(getNextId()).padStart(4, '0'),
      title: '', // タイトルを追加
      occurredOn: (() => {
        const d = new Date();
        return `${d.getMonth()+1}/${d.getDate()}`;
      })(),
      reporter: '',
      origin: '',
      originNumber: '', // 起因番号を追加
      reproRate: '',
      steps: '',
      expected: '',
      actual: ''
    };

    // 1行目: ID, 発生日, 登録者（横並び3列）
    const row1 = el('div', { class: 'form-row', style: 'display:flex;gap:16px;' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: 'ID *' }));
        const input = el('input', { type: 'text', readonly: 'readonly', style: 'width:100%;text-align:center;' });
        input.value = newBugData.id;
        input.dataset.key = 'id';
        fld.appendChild(input);
        return fld;
      })(),
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: '発生日 *' }));
        const input = el('input', { type: 'text', style: 'width:100%;text-align:center;' });
        input.value = newBugData.occurredOn;
        input.required = true;
        input.placeholder = 'm/d';
        input.dataset.key = 'occurredOn';
        fld.appendChild(input);
        return fld;
      })(),
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: '登録者 *' }));
        const input = el('select', { style: 'width:100%;' });
        
        // 動的に登録者リストを設定
        const reporterOptions = ['', ...REPORTER_LIST];
        reporterOptions.forEach(o => {
          const op = el('option', { value: o, text: o || '(選択)' });
          // 前回選択した登録者を初期選択
          if (o === state.lastSelectedReporter) {
            op.selected = true;
            newBugData.reporter = o;
          }
          input.appendChild(op);
        });
        
        // 登録者変更時にlocalStorageに保存
        input.addEventListener('change', (e) => {
          const selectedReporter = e.target.value;
          state.lastSelectedReporter = selectedReporter;
          localStorage.setItem('bugTracker_lastReporter', selectedReporter);
        });
        
        input.required = true;
        input.dataset.key = 'reporter';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 2行目: 発生起因, 起因番号, 再現率（横並び3列）
    const row2 = el('div', { class: 'form-row', style: 'display:flex;gap:16px;' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: '発生起因 *' }));
        const input = el('select', { style: 'width:100%;' });
        ['','定義(通常)','定義(電源断)','定義(通信断)'].forEach(o => {
          const op = el('option', { value: o, text: o || '(選択)' });
          input.appendChild(op);
        });
        input.required = true;
        input.dataset.key = 'origin';
        fld.appendChild(input);
        return fld;
      })(),
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: '起因番号' }));
        const input = el('input', { type: 'text', style: 'width:100%;' });
        input.value = newBugData.originNumber;
        input.placeholder = '起因番号を入力';
        input.dataset.key = 'originNumber';
        fld.appendChild(input);
        return fld;
      })(),
      (() => {
        const fld = el('div', { class: 'field', style: 'flex:1;' });
        fld.appendChild(el('label', { text: '再現率 *' }));
        const input = el('select', { style: 'width:100%;' });
        ['','毎回','時々','1回のみ'].forEach(o => {
          const op = el('option', { value: o, text: o || '(選択)' });
          input.appendChild(op);
        });
        input.required = true;
        input.dataset.key = 'reproRate';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 3行目: タイトル
    const row3 = el('div', { class: 'form-row' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'width:100%;' });
        fld.appendChild(el('label', { text: 'タイトル *' }));
        const input = el('input', { type: 'text', style: 'width:98%;' });
        input.value = newBugData.title;
        input.required = true;
        input.placeholder = 'バグのタイトルを入力してください';
        input.dataset.key = 'title';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 4行目: 再現手順（10行）
    const row4 = el('div', { class: 'form-row' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'width:100%;' });
        fld.appendChild(el('label', { text: '再現手順 *' }));
        const input = el('textarea', { rows: 10, style: 'width:98%;' });
        input.value = newBugData.steps;
        input.required = true;
        input.placeholder = '1. 再現手順を記載してください\n2. 詳細な操作手順\n3. 発生までの流れ';
        input.dataset.key = 'steps';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 5行目: 期待する動作
    const row5 = el('div', { class: 'form-row' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'width:100%;' });
        fld.appendChild(el('label', { text: '期待する動作 *' }));
        const input = el('textarea', { rows: 2, style: 'width:98%;' });
        input.value = newBugData.expected;
        input.required = true;
        input.dataset.key = 'expected';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 6行目: 実際の動作
    const row6 = el('div', { class: 'form-row' }, [
      (() => {
        const fld = el('div', { class: 'field', style: 'width:100%;' });
        fld.appendChild(el('label', { text: '実際の動作 *' }));
        const input = el('textarea', { rows: 2, style: 'width:98%;' });
        input.value = newBugData.actual;
        input.required = true;
        input.dataset.key = 'actual';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // レイアウトをbodyに追加
    body.appendChild(row1);
    body.appendChild(row2);
    body.appendChild(row3); // タイトル
    body.appendChild(row4); // 再現手順
    body.appendChild(row5); // 期待する動作
    body.appendChild(row6); // 実際の動作

    $('#modal').classList.remove('hidden');
  }

  async function saveNewBug(bugData) {
    if (!state.inOffice) {
      // デモモードの場合
      const maxRowIndex = Math.max(0, ...state.bugs.map(b => b.rowIndex || 0));
      bugData.rowIndex = maxRowIndex + 1;
      state.bugs.push(bugData);
      return;
    }

    setStatus('保存中...');
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const used = sheet.getUsedRange(true);
      used.load(['rowCount']);
      await ctx.sync();
      
      const newRowIndex = (used.rowCount || DATA_START - 1) + 1;
      bugData.rowIndex = newRowIndex;
      
      const columns = getColumns(); // 一度取得して再利用
      const rowVals = [];
      for (let c = 0; c < COL_COUNT; c++) {
        const colDef = columns[c];
        let v = bugData[colDef.key];
        if (v === undefined || v === null) v = '';
        rowVals.push(v);
      }
      
      const writeRange = sheet.getRangeByIndexes(newRowIndex - 1, 0, 1, COL_COUNT);
      writeRange.values = [rowVals];
      await ctx.sync();
      
      state.bugs.push(bugData);
    });
    setStatus('登録しました');
    setTimeout(() => setStatus(''), 2000);
  }

  function validateNewBugForm() {
    const errors = [];
    const formData = {};
    const modalBody = $('#modal-body');
    modalBody.querySelectorAll('[data-key]').forEach(inp => {
      const k = inp.dataset.key;
      let value = inp.value.trim();
      formData[k] = value;

      // 必須項目チェック
      if ([
        'id', 'title', 'occurredOn', 'reporter', 'origin', 'reproRate', 'steps', 'expected', 'actual'
      ].includes(k)) {
        if (!value) {
          errors.push(`${inp.previousSibling ? inp.previousSibling.textContent.replace('*','').trim() : k}は必須項目です`);
          inp.style.backgroundColor = '#ffebee';
        } else {
          inp.style.backgroundColor = '';
        }
      }

      // 発生日のm/d形式チェック
      if (k === 'occurredOn' && value) {
        if (!/^\d{1,2}\/\d{1,2}$/.test(value)) {
          errors.push('発生日は m/d 形式で入力してください');
          inp.style.backgroundColor = '#ffebee';
        } else {
          // 実在日付か判定
          const [m, d] = value.split('/').map(Number);
          const dt = new Date(2026, m - 1, d);
          if (dt.getMonth() + 1 !== m || dt.getDate() !== d) {
            errors.push('発生日が不正です');
            inp.style.backgroundColor = '#ffebee';
          }
        }
      }
    });
    return { isValid: errors.length === 0, errors, formData };
  }

  async function saveModal() {
    if (state.editingRow === 'new') {
      // 新規登録の場合
      const validation = validateNewBugForm();
      if (!validation.isValid) {
        alert('入力エラー:\n' + validation.errors.join('\n'));
        return;
      }
      
      const newBugData = validation.formData;
      // 自動設定項目
      newBugData.id = getNextId();
      newBugData.status = '新規';
      newBugData.updated = newBugData.occurredOn; // 更新日 = 発生日
      newBugData.assignee = newBugData.reporter; // 更新者 = 登録者（この場合担当者として設定）
      newBugData.priority = '中'; // 優先度の初期値
      newBugData.severity = ''; // 影響度の初期値
      newBugData.tag = ''; // タグの初期値
      
      try {
        await saveNewBug(newBugData);
        render();
        closeModal();
      } catch (e) {
        console.error(e);
        setStatus('登録失敗: ' + (e.message || e));
      }
    } else {
      // 既存のバグ編集の場合
      const bug = state.bugs.find(b => b.rowIndex === state.editingRow);
      if (!bug) { closeModal(); return; }
      
      // ラジオボタンの状態を確認
      const selectedRadio = $('#modal-body').querySelector('input[type="radio"]:checked');
      const isComplete = selectedRadio && selectedRadio.value === 'complete';
      const isReject = selectedRadio && selectedRadio.value === 'reject';
      
      $('#modal-body').querySelectorAll('[data-key]').forEach(inp => {
        const k = inp.dataset.key;
        const col = getColumns().find(c => c.key === k);
        if (!col || col.type === 'readonly') return;
        if (!['kekkakanryo', 'sashimodoshi'].includes(k)) { // ラジオボタンは除外
          bug[k] = inp.value;
        }
      });
      
      // 管理タブの変更を確実に反映
      const prioritySelect = $('#modal-body').querySelector('[data-key="priority"]');
      const severitySelect = $('#modal-body').querySelector('[data-key="severity"]');
      const tagInput = $('#modal-body').querySelector('[data-key="tag"]');
      if (prioritySelect) bug.priority = prioritySelect.value;
      if (severitySelect) bug.severity = severitySelect.value;
      if (tagInput) bug.tag = tagInput.value;
      
      // 処置完了チェックボックスの状態を確認
      const shochoKanryoCheck = $('#modal-body').querySelector('[data-key="shochikanryo"]');
      const isShochoKanryo = shochoKanryoCheck && shochoKanryoCheck.checked;
      
      // 解析完了チェックボックスの状態を確認
      const kaisekiKanryoCheck = $('#modal-body').querySelector('[data-key="kaisekikanryo"]');
      const isKaisekiKanryo = kaisekiKanryoCheck && kaisekiKanryoCheck.checked;
      
      // 現在アクティブなタブを確認
      const activeTab = $('#modal-body').querySelector('.tab-btn.active');
      const isShochoTab = activeTab && activeTab.dataset.tab === 'shochi';
      
      // 影響範囲の更新は処置タブでの操作時のみ実行（差し戻し時の影響範囲消失を防止）
      if (isShochoTab) {
        // 影響範囲のチェックボックスを集約
        const scopeCheckboxes = $('#modal-body').querySelectorAll('[data-scope-option]');
        const selectedScopes = [];
        scopeCheckboxes.forEach(cb => {
          if (cb.checked) {
            selectedScopes.push(cb.value);
          }
        });
        
        // 修正完了状況のチェックボックスを集約
        const completedCheckboxes = $('#modal-body').querySelectorAll('[data-scope-completed]');
        const completedScopes = [];
        completedCheckboxes.forEach(cb => {
          if (cb.checked) {
            completedScopes.push(cb.value);
          }
        });
        bug.scopeCompleted = completedScopes.join('/');
        
        // 修正対象と修正完了を統合して影響範囲に設定
        const scopeWithStatus = selectedScopes.map(scope => {
          return completedScopes.includes(scope) ? `${scope}（済）` : scope;
        });
        bug.scope = scopeWithStatus.join('/');
        
        // state.bugs配列も更新（Excel保存前に確実に反映）
        const bugIndex = state.bugs.findIndex(b => b.rowIndex === bug.rowIndex);
        if (bugIndex >= 0) {
          state.bugs[bugIndex].scope = bug.scope;
          state.bugs[bugIndex].scopeCompleted = bug.scopeCompleted;
        }
      }
      
      // 処置完了がチェックされている場合のバリデーションと更新（処置タブでの操作時のみ）
      if (isShochoKanryo && isShochoTab) {
        // 影響範囲のチェックボックス情報を再取得（処置タブの場合のみ）
        const scopeCheckboxes = $('#modal-body').querySelectorAll('[data-scope-option]');
        const selectedScopes = [];
        scopeCheckboxes.forEach(cb => {
          if (cb.checked) {
            selectedScopes.push(cb.value);
          }
        });
        
        const completedCheckboxes = $('#modal-body').querySelectorAll('[data-scope-completed]');
        const completedScopes = [];
        completedCheckboxes.forEach(cb => {
          if (cb.checked) {
            completedScopes.push(cb.value);
          }
        });
        
        // 修正対象にチェックが付いている場合のバリデーション
        if (selectedScopes.length > 0) {
          // 修正完了チェックの確認
          const incompleteScopes = selectedScopes.filter(scope => !completedScopes.includes(scope));
          if (incompleteScopes.length > 0) {
            alert(`処置完了にするには、以下の修正対象の修正完了にもチェックを付けてください：\n${incompleteScopes.join('\n')}`);
            return; // 保存を中止
          }
          
          // 修正Verと処置内容の入力確認
          if (!bug.fixVer || bug.fixVer.trim() === '') {
            alert('修正対象にチェックが付いている場合、修正Verの入力が必要です');
            return;
          }
          if (!bug.fix || bug.fix.trim() === '') {
            alert('修正対象にチェックが付いている場合、処置内容の入力が必要です');
            return;
          }
        }
        
        // 処置完了時の自動更新
        if (bug.status === '修正') {
          const today = new Date();
          const year = today.getFullYear();
          const month = String(today.getMonth() + 1).padStart(2, '0');
          const day = String(today.getDate()).padStart(2, '0');
          bug.fixDate = `${year}-${month}-${day}`;
          bug.status = '確認';
          bug.verifier = bug.reporter; // 確認者を登録者に設定
          bug.assignee = bug.reporter; // 担当者を登録者に設定
          bug.reject = ''; // 差し戻しをクリア
          setStatus('処置完了のため対応日を当日に設定し、状況を「確認」に変更、確認者と担当者を登録者で設定、差し戻しをクリアしました');
        }
      } else if (isShochoKanryo && !isShochoTab) {
        // 処置タブ以外で処置完了がチェックされた場合は警告
        alert('処置完了は処置タブで実行してください');
        return;
      }
      
      // 解析完了がチェックされている場合、解析日を当日に設定し、状況を修正に変更
      if (isKaisekiKanryo && bug.status === '解析') {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        bug.analysisDate = `${year}-${month}-${day}`;
        bug.status = '修正';
        bug.fixer = bug.analyst; // 対応者を解析者の名前で更新
        setStatus('解析完了のため解析日を当日に設定し、状況を「修正」に変更、対応者を解析者で設定しました');
      }
      
      // ラジオボタンの選択に応じて状態を変更
      if (isComplete && bug.status === '確認') {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        bug.verifyDate = `${year}-${month}-${day}`; // 確認日を当日に設定
        bug.status = '完了';
        setStatus('確認完了のため確認日を当日に設定し、状態を「完了」に変更しました');
      } else if (isReject && bug.status === '確認') {
        bug.status = '修正';
        bug.reject = '○'; // 差し戻し列を○で更新
        bug.fixDate = ''; // 対応日を空欄に
        bug.assignee = bug.fixer; // 担当者を対応者に設定
        setStatus('差し戻しのため状態を「修正」に変更し、対応日を空欄にして担当者を対応者に設定しました');
      }
      
      try {
        // 影響範囲の修正完了状態を含めてExcelに保存
        await saveBugToExcel(bug);
        render();
        closeModal();
      } catch (e) {
        console.error(e);
        setStatus('保存失敗: ' + (e.message || e));
      }
    }
  }

  function setView(v) {
    state.view = v;
    $('#btn-view-assignee').classList.toggle('active', v === 'assignee');
    $('#btn-view-status').classList.toggle('active',   v === 'status');
    $('#btn-view-trend').classList.toggle('active',   v === 'trend');
    $('#view-assignee').classList.toggle('active', v === 'assignee');
    $('#view-status').classList.toggle('active',   v === 'status');
    $('#view-trend').classList.toggle('active',   v === 'trend');
    render();
  }
  function render() {
    if (state.view === 'assignee') renderKanbanAssignee();
    else if (state.view === 'trend') renderTrend();
    else renderKanbanStatus();
  }

  function demoData() {
    return [
      { rowIndex: 4, id: 1, title: 'ログイン後に画面が真っ白', status: '解析', updated: '2025-04-10', assignee: '高橋',
        occurredOn: '2025-04-08', reporter: '政次', origin: '定義(通常)', steps: '1.ログイン\n2.TOPへ', expected: 'TOP表示', actual: '真っ白', reproRate: '毎回',
        cause: '', analyst: '政次', analysisDate: '2025-04-10', scope: 'アプリ', fix: '', fixVer: '', fixer: '', fixDate: '', verify: '', verifier: '', verifyDate: '', tag: 'UI', priority: '高', severity: '致命的' },
      { rowIndex: 5, id: 2, title: '通信断時にRPAが停止', status: '修正', updated: '2025-04-12', assignee: '伊藤',
        occurredOn: '2025-04-09', reporter: '松田', origin: '定義(通信断)', steps: '1.通信断発生', expected: '自動復旧', actual: '停止のまま', reproRate: '時々',
        cause: 'タイムアウト未設定', analyst: '伊藤', analysisDate: '2025-04-11', scope: 'RPA', fix: 'リトライ実装', fixVer: 'v1.2', fixer: '伊藤', fixDate: '2025-04-12', verify: '', verifier: '', verifyDate: '', tag: 'RPA', priority: '中', severity: '重大' },
      { rowIndex: 6, id: 3, title: '電源断後に設定が消える', status: '新規', updated: '', assignee: '',
        occurredOn: '2025-04-14', reporter: '高橋', origin: '定義(電源断)', steps: '1.電源断', expected: '保持', actual: '消失', reproRate: '1回のみ',
        cause: '', analyst: '', analysisDate: '', scope: '', fix: '', fixVer: '', fixer: '', fixDate: '', verify: '', verifier: '', verifyDate: '', tag: '', priority: '低', severity: '警備' },
      { rowIndex: 7, id: 4, title: 'タイトル文字化け', status: '完了', updated: '2025-04-13', assignee: '松田',
        occurredOn: '2025-04-05', reporter: '政次', origin: '定義(通常)', steps: '', expected: '', actual: '', reproRate: '毎回',
        cause: 'エンコード不一致', analyst: '松田', analysisDate: '2025-04-06', scope: 'アプリ', fix: 'UTF-8統一', fixVer: 'v1.1', fixer: '松田', fixDate: '2025-04-07', verify: '解消確認', verifier: '政次', verifyDate: '2025-04-13', tag: 'i18n', priority: '中', severity: '重大' }
    ];
  }

  function bindEvents() {
    $('#btn-view-assignee').addEventListener('click', () => setView('assignee'));
    $('#btn-view-status').addEventListener('click',   () => setView('status'));
    $('#btn-view-trend').addEventListener('click',   () => setView('trend'));
    $('#btn-add-new').addEventListener('click', () => openNewBugModal());
    $('#btn-reload').addEventListener('click', async () => {
      await loadFromExcel();
      render();
    });
    $('#filter-text').addEventListener('input', (e) => { state.filters.text = e.target.value; render(); });
    $('#filter-priority').addEventListener('change', (e) => { state.filters.priority = e.target.value; render(); });
    $('#filter-status').addEventListener('change',   (e) => { state.filters.status   = e.target.value; render(); });
    $('#modal-save').addEventListener('click', saveModal);
    $('#modal-cancel').addEventListener('click', closeModal);
    $('#modal-close').addEventListener('click', closeModal);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !$('#modal').classList.contains('hidden')) {
        closeModal();
      }
    });
  }

  function init() {
    bindEvents();
    ensureOffice(() => {
      loadFromExcel().then(render);
    });
  }

  init();
})();
