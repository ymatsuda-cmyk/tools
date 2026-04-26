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
  const COL_COUNT  = 26;

  const COLUMNS = [
    { key: 'id',         letter: 'A', label: 'ID',           group: '基本情報', type: 'readonly' },
    { key: 'title',      letter: 'B', label: 'タイトル',      group: '基本情報', type: 'text' },
    { key: 'status',     letter: 'C', label: '状況',         group: '基本情報', type: 'select', options: ['新規','解析待ち','修正待ち','確認待ち','再発','完了'] },
    { key: 'updated',    letter: 'D', label: '更新日',       group: '基本情報', type: 'date' },
    { key: 'assignee',   letter: 'E', label: '担当者',       group: '基本情報', type: 'select', options: ['','政次','高橋','伊藤','松田'] },
    { key: 'occurredOn', letter: 'F', label: '発生日',       group: '発生情報', type: 'date' },
    { key: 'reporter',   letter: 'G', label: '登録者',       group: '発生情報', type: 'select', options: ['','政次','高橋','伊藤','松田'] },
    { key: 'origin',     letter: 'H', label: '発生起因',     group: '発生情報', type: 'select', options: ['','定義(通常)','定義(電源断)','定義(通信断)'] },
    { key: 'steps',      letter: 'I', label: '再現手順',     group: '発生情報', type: 'textarea' },
    { key: 'expected',   letter: 'J', label: '期待する動作', group: '発生情報', type: 'textarea' },
    { key: 'actual',     letter: 'K', label: '実際の動作',   group: '発生情報', type: 'textarea' },
    { key: 'reproRate',  letter: 'L', label: '再現率',       group: '発生情報', type: 'select', options: ['','毎回','時々','1回のみ'] },
    { key: 'cause',      letter: 'M', label: '原因',         group: '対応情報', type: 'textarea' },
    { key: 'analyst',    letter: 'N', label: '解析者',       group: '対応情報', type: 'select', options: ['','政次','高橋','伊藤','松田'] },
    { key: 'analysisDate', letter: 'O', label: '解析日',     group: '対応情報', type: 'date' },
    { key: 'scope',      letter: 'P', label: '影響範囲',     group: '対応情報', type: 'select', options: ['','定義(通常)','定義(電源断)','定義(通信断)','RPA','アプリ'] },
    { key: 'fix',        letter: 'Q', label: '対応内容',     group: '対応情報', type: 'textarea' },
    { key: 'fixVer',     letter: 'R', label: '修正Ver',     group: '対応情報', type: 'text' },
    { key: 'fixer',      letter: 'S', label: '対応者',       group: '対応情報', type: 'select', options: ['','政次','高橋','伊藤','松田'] },
    { key: 'fixDate',    letter: 'T', label: '対応日',       group: '対応情報', type: 'date' },
    { key: 'verify',     letter: 'U', label: '確認内容',     group: '結果確認', type: 'textarea' },
    { key: 'verifier',   letter: 'V', label: '確認者',       group: '結果確認', type: 'select', options: ['','政次','高橋','伊藤','松田'] },
    { key: 'verifyDate', letter: 'W', label: '確認日',       group: '結果確認', type: 'date' },
    { key: 'tag',        letter: 'X', label: 'タグ',         group: '管理',     type: 'text' },
    { key: 'priority',   letter: 'Y', label: '優先度',       group: '管理',     type: 'select', options: ['','高','中','低'] },
    { key: 'severity',   letter: 'Z', label: '影響度',       group: '管理',     type: 'select', options: ['','致命的','重大','警備'] }
  ];

  const STATUS_ORDER = ['新規','解析待ち','修正待ち','確認待ち','完了'];
  const ASSIGNEE_ORDER = ['(未割当)','政次','高橋','伊藤','松田'];
  const PRIORITY_RANK = { '高': 0, '中': 1, '低': 2, '': 3 };

  const state = {
    bugs: [],
    view: 'assignee',  // 'assignee' または 'status'
    filters: { text: '', priority: '', status: '' },
    inOffice: false,
    editingRow: null
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
      return;
    }
    setStatus('読み込み中...');
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
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
      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        if (row.every(v => v === '' || v === null)) continue;
        const obj = { rowIndex: DATA_START + r };
        for (let c = 0; c < COL_COUNT; c++) {
          const colDef = COLUMNS[c];
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
      const rowVals = [];
      for (let c = 1; c < COL_COUNT; c++) {
        const colDef = COLUMNS[c];
        let v = bug[colDef.key];
        if (v === undefined || v === null) v = '';
        rowVals.push(v);
      }
      writeRange.values = [rowVals];
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

    const bugs = sortByPriority(applyFilters(state.bugs));
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
      const header = el('div', { class: 'kanban-col-header' }, [
        el('span', { text: key || '(未設定)' }),
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
    
    // 未割当から担当者への移動時は状態を「解析待ち」に変更し、解析者を設定
    if (oldAssignee === '(未割当)' && newAssignee !== '(未割当)') {
      bug.status = '解析待ち';
      bug.analyst = newAssignee; // 解析者に担当者を設定
      setStatus(`担当者を「${newAssignee}」に変更し、状況を「解析待ち」に変更しました`);
    } else {
      setStatus(`担当者を「${newAssignee}」に変更しました`);
    }
    
    setTimeout(() => setStatus(''), 3000);
    
    // Excel保存
    try {
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
    
    // 1行目：左端にID、発生起因。右端に優先度
    const row1 = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;' });
    const leftPart1 = el('div', { style: 'display:flex;gap:8px;' }, [
      el('span', { class: 'id', text: `#${b.id || ''}`, style: 'font-weight:bold;font-size:12px;' }),
      el('span', { text: b.origin || '', style: 'font-size:11px;color:#666;' })
    ]);
    const rightPart1 = el('div', {});
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
      leftPart3.appendChild(el('span', { class: `badge st-${b.status}`, text: b.status, style: 'font-size:11px;' }));
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
        nameText = b.analyst || '(未設定)';
        if (!b.analyst) nameStyle += 'color:#999;';
        // 解析日を表示
        if (b.analysisDate && b.analyst) {
          const date = new Date(b.analysisDate);
          const month = date.getMonth() + 1;
          const day = date.getDate();
          nameText += ` (${month}/${day})`;
        }
        break;
      case '修正待ち':
        nameText = b.fixer || '(未設定)';
        if (!b.fixer) nameStyle += 'color:#999;';
        // 対応日を表示
        if (b.fixDate && b.fixer) {
          const date = new Date(b.fixDate);
          const month = date.getMonth() + 1;
          const day = date.getDate();
          nameText += ` (${month}/${day})`;
        }
        break;
      case '確認待ち':
        nameText = b.verifier || '(未設定)';
        if (!b.verifier) nameStyle += 'color:#999;';
        // 確認日を表示
        if (b.verifyDate && b.verifier) {
          const date = new Date(b.verifyDate);
          const month = date.getMonth() + 1;
          const day = date.getDate();
          nameText += ` (${month}/${day})`;
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
      '解析待ち': 'kaiseki',
      '修正待ち': 'shochi',
      '確認待ち': 'kekka',
      '再発': 'kekka',
      '完了': 'jisho'
    };
    let activeTab = statusTabMap[bug.status] || 'jisho';

    // 常時表示エリア
    const alwaysArea = el('div', { class: 'always-area', style: 'margin-bottom:12px;padding:8px 0;border-bottom:1px solid #ccc;' });
    
    // 1行目：ID、状態、発生起因
    const row1 = el('div', { style: 'display:flex;gap:16px;flex-wrap:wrap;margin-bottom:8px;' }, [
      el('div', {}, [el('b', { text: 'ID: ' }), el('span', { text: bug.id || '' })]),
      el('div', {}, [el('b', { text: '状態: ' }), el('span', { text: bug.status || '' })]),
      el('div', {}, [el('b', { text: '発生起因: ' }), el('span', { text: bug.origin || '' })])
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
    
    // ワークフローテーブル
    const workflowTable = el('table', { style: 'width:100%;border-collapse:collapse;font-size:12px;' });
    
    // ヘッダー行
    const thead = el('thead');
    const headerRow = el('tr');
    ['新規', '解析', '修正', '確認'].forEach(header => {
      headerRow.appendChild(el('th', { 
        text: header, 
        style: 'border:1px solid #ddd;padding:6px;background:#f5f5f5;text-align:center;font-weight:bold;' 
      }));
    });
    thead.appendChild(headerRow);
    workflowTable.appendChild(thead);
    
    // データ行
    const tbody = el('tbody');
    const dataRow = el('tr');
    
    // 新規列
    const newCell = el('td', { style: 'border:1px solid #ddd;padding:6px;text-align:center;' });
    if (bug.reporter || bug.occurredOn) {
      const reporterText = bug.reporter || '(未設定)';
      const dateText = formatToMD(bug.occurredOn);
      const displayText = (dateText && dateText !== '') ? `${reporterText} (${dateText})` : reporterText;
      newCell.appendChild(el('div', { text: displayText, style: 'font-weight:bold;' }));
    } else {
      newCell.appendChild(el('div', { text: '-', style: 'color:#ccc;' }));
    }
    dataRow.appendChild(newCell);
    
    // 解析列
    const analysisCell = el('td', { style: 'border:1px solid #ddd;padding:6px;text-align:center;' });
    if (bug.status !== '新規') {
      const analystText = bug.analyst || '(未設定)';
      const analysisDate = formatToMD(bug.updated);
      const displayText = (analysisDate && analysisDate !== '') ? `${analystText} (${analysisDate})` : analystText;
      analysisCell.appendChild(el('div', { text: displayText, style: 'font-weight:bold;' }));
    } else {
      analysisCell.appendChild(el('div', { text: '-', style: 'color:#ccc;' }));
    }
    dataRow.appendChild(analysisCell);
    
    // 修正列
    const fixCell = el('td', { style: 'border:1px solid #ddd;padding:6px;text-align:center;' });
    if (['修正待ち', '確認待ち', '完了'].includes(bug.status)) {
      const fixerText = bug.fixer || '(未設定)';
      const fixDate = formatToMD(bug.updated);
      const displayText = (fixDate && fixDate !== '') ? `${fixerText} (${fixDate})` : fixerText;
      fixCell.appendChild(el('div', { text: displayText, style: 'font-weight:bold;' }));
    } else {
      fixCell.appendChild(el('div', { text: '-', style: 'color:#ccc;' }));
    }
    dataRow.appendChild(fixCell);
    
    // 確認列
    const verifyCell = el('td', { style: 'border:1px solid #ddd;padding:6px;text-align:center;' });
    if (['確認待ち', '完了'].includes(bug.status)) {
      const verifierText = bug.verifier || '(未設定)';
      const verifyDate = formatToMD(bug.updated);
      const displayText = (verifyDate && verifyDate !== '') ? `${verifierText} (${verifyDate})` : verifierText;
      verifyCell.appendChild(el('div', { text: displayText, style: 'font-weight:bold;' }));
    } else {
      verifyCell.appendChild(el('div', { text: '-', style: 'color:#ccc;' }));
    }
    dataRow.appendChild(verifyCell);
    
    tbody.appendChild(dataRow);
    workflowTable.appendChild(tbody);
    alwaysArea.appendChild(workflowTable);
    
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
      if (tabKey === 'jisho') {
        // 事象タブ：再現手順、期待する動作、実際の動作（編集可）
        console.log('steps:', bug.steps, '| expected:', bug.expected, '| actual:', bug.actual);
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
            return ta;
          })()
        ]));
        tabContent.appendChild(el('div', { style: 'margin-top:8px;' }, [
          el('label', {}, [
            el('input', { type: 'checkbox', 'data-key': 'kaisekikanryo' }),
            el('span', { text: '解析完了（修正待ちに変更）' })
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
        const selectedScopes = currentScope.split('/').map(s => s.trim()).filter(s => s);
        
        // 修正完了状況を取得（新しいフィールドとして追加）
        const currentCompleted = bug.scopeCompleted || '';
        const completedScopes = currentCompleted.split('/').map(s => s.trim()).filter(s => s);
        
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
            return ta;
          })()
        ]));
        
        mainContainer.appendChild(rightPanel);
        tabContent.appendChild(mainContainer);
        
        // 下部：修正Ver、処置完了チェック
        tabContent.appendChild(el('div', {}, [el('label', { text: '修正Ver' }), el('br'),
          el('input', { type: 'text', style: 'width:98%;', value: bug.fixVer || '', 'data-key': 'fixVer' })]));
        tabContent.appendChild(el('div', { style: 'margin-top:8px;' }, [
          el('label', {}, [
            el('input', { type: 'checkbox', 'data-key': 'shochikanryo' }),
            el('span', { text: '処置完了（確認待ちに変更）' })
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
            return ta;
          })()
        ]));
        
        // ラジオボタングループ
        const radioGroup = el('div', { style: 'margin-top:8px;' });
        const groupName = `result_${bug.rowIndex || 'new'}`;
        
        // 確認完了ラジオボタン
        const completeRadio = el('label', { style: 'margin-right:16px;' }, [
          el('input', { 
            type: 'radio', 
            name: groupName, 
            value: 'complete', 
            'data-key': 'kekkakanryo' 
          }),
          el('span', { text: '確認完了（完了に変更）' })
        ]);
        
        // 差し戻しラジオボタン
        const rejectRadio = el('label', {}, [
          el('input', { 
            type: 'radio', 
            name: groupName, 
            value: 'reject', 
            'data-key': 'sashimodoshi' 
          }),
          el('span', { text: '差し戻し（修正待ちに変更）' })
        ]);
        
        radioGroup.appendChild(completeRadio);
        radioGroup.appendChild(rejectRadio);
        tabContent.appendChild(radioGroup);
      } else if (tabKey === 'kanri') {
        // 管理タブ：タグ、優先度、影響度（編集可）
        tabContent.appendChild(el('div', {}, [el('label', { text: 'タグ' }), el('br'),
          el('input', { type: 'text', style: 'width:98%;', value: bug.tag || '', 'data-key': 'tag' })]));
        tabContent.appendChild(el('div', {}, [el('label', { text: '優先度' }), el('br'),
          el('input', { type: 'text', style: 'width:98%;', value: bug.priority || '', 'data-key': 'priority' })]));
        tabContent.appendChild(el('div', {}, [el('label', { text: '影響度' }), el('br'),
          el('input', { type: 'text', style: 'width:98%;', value: bug.severity || '', 'data-key': 'severity' })]));
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
      occurredOn: (() => {
        const d = new Date();
        return `${d.getMonth()+1}/${d.getDate()}`;
      })(),
      reporter: '',
      origin: '',
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
        ['', '政次', '高橋', '伊藤', '松田'].forEach(o => {
          const op = el('option', { value: o, text: o || '(選択)' });
          input.appendChild(op);
        });
        input.required = true;
        input.dataset.key = 'reporter';
        fld.appendChild(input);
        return fld;
      })()
    ]);

    // 2行目: 発生起因, 再現率（横並び2列）
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

    // 3行目: 再現手順（10行）
    const row3 = el('div', { class: 'form-row' }, [
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

    // 4行目: 期待する動作
    const row4 = el('div', { class: 'form-row' }, [
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

    // 5行目: 実際の動作
    const row5 = el('div', { class: 'form-row' }, [
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
    body.appendChild(row3);
    body.appendChild(row4);
    body.appendChild(row5);

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
      
      const rowVals = [];
      for (let c = 0; c < COL_COUNT; c++) {
        const colDef = COLUMNS[c];
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
        'id', 'occurredOn', 'reporter', 'origin', 'reproRate', 'steps', 'expected', 'actual'
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
        const col = COLUMNS.find(c => c.key === k);
        if (!col || col.type === 'readonly') return;
        if (!['kekkakanryo', 'sashimodoshi'].includes(k)) { // ラジオボタンは除外
          bug[k] = inp.value;
        }
      });
      
      // 処置完了チェックボックスの状態を確認
      const shochoKanryoCheck = $('#modal-body').querySelector('[data-key="shochikanryo"]');
      const isShochoKanryo = shochoKanryoCheck && shochoKanryoCheck.checked;
      
      // 解析完了チェックボックスの状態を確認
      const kaisekiKanryoCheck = $('#modal-body').querySelector('[data-key="kaisekikanryo"]');
      const isKaisekiKanryo = kaisekiKanryoCheck && kaisekiKanryoCheck.checked;
      
      // 影響範囲のチェックボックスを集約
      const scopeCheckboxes = $('#modal-body').querySelectorAll('[data-scope-option]');
      const selectedScopes = [];
      scopeCheckboxes.forEach(cb => {
        if (cb.checked) {
          selectedScopes.push(cb.value);
        }
      });
      bug.scope = selectedScopes.join('/');
      
      // 修正完了状況のチェックボックスを集約
      const completedCheckboxes = $('#modal-body').querySelectorAll('[data-scope-completed]');
      const completedScopes = [];
      completedCheckboxes.forEach(cb => {
        if (cb.checked) {
          completedScopes.push(cb.value);
        }
      });
      bug.scopeCompleted = completedScopes.join('/');
      
      // 処置完了がチェックされている場合、修正対象がすべて修正完了になっているかを確認
      if (isShochoKanryo && selectedScopes.length > 0) {
        const incompleteScopes = selectedScopes.filter(scope => !completedScopes.includes(scope));
        if (incompleteScopes.length > 0) {
          alert(`処置完了にするには、以下の修正対象の修正完了にもチェックを付けてください：\n${incompleteScopes.join('\n')}`);
          return; // 保存を中止
        }
      }
      
      // 解析完了がチェックされている場合、解析日を当日に設定し、状況を修正待ちに変更
      if (isKaisekiKanryo && bug.status === '解析待ち') {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        bug.analysisDate = `${year}-${month}-${day}`;
        bug.status = '修正待ち';
        setStatus('解析完了のため解析日を当日に設定し、状況を「修正待ち」に変更しました');
      }
      
      // ラジオボタンの選択に応じて状態を変更
      if (isComplete && bug.status === '確認待ち') {
        bug.status = '完了';
        setStatus('確認完了のため状態を「完了」に変更しました');
      } else if (isReject && bug.status === '確認待ち') {
        bug.status = '修正待ち';
        setStatus('差し戻しのため状態を「修正待ち」に変更しました');
      }
      
      try {
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
    $('#view-assignee').classList.toggle('active', v === 'assignee');
    $('#view-status').classList.toggle('active',   v === 'status');
    render();
  }
  function render() {
    if (state.view === 'assignee') renderKanbanAssignee();
    else renderKanbanStatus();
  }

  function demoData() {
    return [
      { rowIndex: 4, id: 1, title: 'ログイン後に画面が真っ白', status: '解析待ち', updated: '2025-04-10', assignee: '高橋',
        occurredOn: '2025-04-08', reporter: '政次', origin: '定義(通常)', steps: '1.ログイン\n2.TOPへ', expected: 'TOP表示', actual: '真っ白', reproRate: '毎回',
        cause: '', scope: 'アプリ', fix: '', fixVer: '', fixer: '', verify: '', verifier: '', tag: 'UI', priority: '高', severity: '致命的' },
      { rowIndex: 5, id: 2, title: '通信断時にRPAが停止', status: '修正待ち', updated: '2025-04-12', assignee: '伊藤',
        occurredOn: '2025-04-09', reporter: '松田', origin: '定義(通信断)', steps: '1.通信断発生', expected: '自動復旧', actual: '停止のまま', reproRate: '時々',
        cause: 'タイムアウト未設定', scope: 'RPA', fix: 'リトライ実装', fixVer: 'v1.2', fixer: '伊藤', verify: '', verifier: '', tag: 'RPA', priority: '中', severity: '重大' },
      { rowIndex: 6, id: 3, title: '電源断後に設定が消える', status: '新規', updated: '', assignee: '',
        occurredOn: '2025-04-14', reporter: '高橋', origin: '定義(電源断)', steps: '1.電源断', expected: '保持', actual: '消失', reproRate: '1回のみ',
        cause: '', scope: '', fix: '', fixVer: '', fixer: '', verify: '', verifier: '', tag: '', priority: '低', severity: '警備' },
      { rowIndex: 7, id: 4, title: 'タイトル文字化け', status: '完了', updated: '2025-04-13', assignee: '松田',
        occurredOn: '2025-04-05', reporter: '政次', origin: '定義(通常)', steps: '', expected: '', actual: '', reproRate: '毎回',
        cause: 'エンコード不一致', scope: 'アプリ', fix: 'UTF-8統一', fixVer: 'v1.1', fixer: '松田', verify: '解消確認', verifier: '政次', tag: 'i18n', priority: '中', severity: '重大' }
    ];
  }

  function bindEvents() {
    $('#btn-view-assignee').addEventListener('click', () => setView('assignee'));
    $('#btn-view-status').addEventListener('click',   () => setView('status'));
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
