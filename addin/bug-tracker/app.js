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
  const COL_COUNT  = 22;

  const COLUMNS = [
    { key: 'id',         letter: 'A', label: 'ID',           group: '基本情報', type: 'readonly' },
    { key: 'title',      letter: 'B', label: 'タイトル',      group: '基本情報', type: 'text' },
    { key: 'status',     letter: 'C', label: '状況',         group: '基本情報', type: 'select',
      options: ['新規','解析中','修正待ち','確認待ち','再発','完了'] },
    { key: 'updated',    letter: 'D', label: '更新日',       group: '基本情報', type: 'date' },
    { key: 'assignee',   letter: 'E', label: '担当者',       group: '基本情報', type: 'select',
      options: ['','政次','高橋','伊藤','松田'] },
    { key: 'occurredOn', letter: 'F', label: '発生日',       group: '発生情報', type: 'date' },
    { key: 'reporter',   letter: 'G', label: '登録者',       group: '発生情報', type: 'select',
      options: ['','政次','高橋','伊藤','松田'] },
    { key: 'origin',     letter: 'H', label: '発生起因',     group: '発生情報', type: 'select',
      options: ['','定義(通常)','定義(電源断)','定義(通信断)'] },
    { key: 'steps',      letter: 'I', label: '再現手順',     group: '発生情報', type: 'textarea' },
    { key: 'expected',   letter: 'J', label: '期待する動作', group: '発生情報', type: 'textarea' },
    { key: 'actual',     letter: 'K', label: '実際の動作',   group: '発生情報', type: 'textarea' },
    { key: 'reproRate',  letter: 'L', label: '再現率',       group: '発生情報', type: 'select',
      options: ['','毎回','時々','1回のみ'] },
    { key: 'cause',      letter: 'M', label: '原因',         group: '対応情報', type: 'textarea' },
    { key: 'scope',      letter: 'N', label: '影響範囲',     group: '対応情報', type: 'select',
      options: ['','定義(通常)','定義(電源断)','定義(通信断)','RPA','アプリ'] },
    { key: 'fix',        letter: 'O', label: '対応内容',     group: '対応情報', type: 'textarea' },
    { key: 'fixVer',     letter: 'P', label: '修正Ver',     group: '対応情報', type: 'text' },
    { key: 'fixer',      letter: 'Q', label: '対応者',       group: '対応情報', type: 'select',
      options: ['','政次','高橋','伊藤','松田'] },
    { key: 'verify',     letter: 'R', label: '確認内容',     group: '結果確認', type: 'textarea' },
    { key: 'verifier',   letter: 'S', label: '確認者',       group: '結果確認', type: 'select',
      options: ['','政次','高橋','伊藤','松田'] },
    { key: 'tag',        letter: 'T', label: 'タグ',         group: '管理',     type: 'text' },
    { key: 'priority',   letter: 'U', label: '優先度',       group: '管理',     type: 'select',
      options: ['','高','中','低'] },
    { key: 'severity',   letter: 'V', label: '影響度',       group: '管理',     type: 'select',
      options: ['','致命的','重大','警備'] }
  ];

  const STATUS_ORDER = ['新規','解析中','修正待ち','確認待ち','再発','完了'];
  const ASSIGNEE_ORDER = ['政次','高橋','伊藤','松田','(未割当)'];
  const PRIORITY_RANK = { '高': 0, '中': 1, '低': 2, '': 3 };

  const state = {
    bugs: [],
    view: 'list',
    kanbanGroup: 'status',
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
          else if (v === null) v = '';
          obj[colDef.key] = v;
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

  function renderList() {
    const thead = $('#bug-thead');
    thead.innerHTML = '';
    COLUMNS.forEach(c => thead.appendChild(el('th', { text: c.label })));

    const tbody = $('#bug-tbody');
    tbody.innerHTML = '';

    const bugs = sortByPriority(applyFilters(state.bugs));
    bugs.forEach(b => {
      const tr = el('tr');
      tr.dataset.row = b.rowIndex;
      COLUMNS.forEach(c => {
        const td = el('td');
        const v = b[c.key] ?? '';
        if (c.key === 'priority' && v) {
          td.appendChild(el('span', { class: `badge pri-${v}`, text: v }));
        } else if (c.key === 'status' && v) {
          td.appendChild(el('span', { class: `badge st-${v}`, text: v }));
        } else {
          td.textContent = String(v).replace(/\n/g, ' / ');
          td.title = String(v);
        }
        tr.appendChild(td);
      });
      tr.addEventListener('click', () => openModal(b.rowIndex));
      tbody.appendChild(tr);
    });

    $('#row-count').textContent = `${bugs.length} 件 / 全 ${state.bugs.length} 件`;
  }

  function renderKanban() {
    const board = $('#kanban-board');
    board.innerHTML = '';
    const groupKey = state.kanbanGroup;
    const order = groupKey === 'status' ? STATUS_ORDER : ASSIGNEE_ORDER;

    const bugs = sortByPriority(applyFilters(state.bugs));
    const groups = new Map();
    order.forEach(k => groups.set(k, []));
    bugs.forEach(b => {
      let key = b[groupKey] || '';
      if (groupKey === 'assignee' && !key) key = '(未割当)';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(b);
    });

    groups.forEach((items, key) => {
      const col = el('div', { class: 'kanban-col' });
      const header = el('div', { class: 'kanban-col-header' }, [
        el('span', { text: key || '(未設定)' }),
        el('span', { class: 'count', text: String(items.length) })
      ]);
      const body = el('div', { class: 'kanban-col-body' });
      items.forEach(b => body.appendChild(renderCard(b)));
      col.appendChild(header);
      col.appendChild(body);
      board.appendChild(col);
    });

    $('#row-count').textContent = `${bugs.length} 件 / 全 ${state.bugs.length} 件`;
  }
  function renderCard(b) {
    const card = el('div', { class: 'kanban-card pri-' + (b.priority || '') });
    card.dataset.row = b.rowIndex;
    card.appendChild(el('div', { class: 'id', text: `#${b.id || ''}` }));
    card.appendChild(el('div', { class: 'title', text: b.title || '(無題)' }));
    const meta = el('div', { class: 'meta' });
    if (b.priority) meta.appendChild(el('span', { class: `badge pri-${b.priority}`, text: '優:' + b.priority }));
    if (b.status)   meta.appendChild(el('span', { class: `badge st-${b.status}`,   text: b.status }));
    if (b.assignee) meta.appendChild(el('span', { text: '担当:' + b.assignee }));
    if (b.tag)      meta.appendChild(el('span', { text: 'タグ:' + b.tag }));
    card.appendChild(meta);
    card.addEventListener('click', () => openModal(b.rowIndex));
    return card;
  }

  function openModal(rowIndex) {
    const bug = state.bugs.find(b => b.rowIndex === rowIndex);
    if (!bug) return;
    state.editingRow = rowIndex;
    $('#modal-title').textContent = `バグ詳細  #${bug.id || ''}  ${bug.title || ''}`;

    const body = $('#modal-body');
    body.innerHTML = '';
    const groups = {};
    COLUMNS.forEach(c => { (groups[c.group] = groups[c.group] || []).push(c); });

    Object.keys(groups).forEach(gname => {
      const wrap = el('div', { class: 'field-group' });
      wrap.appendChild(el('h3', { text: gname }));
      groups[gname].forEach(c => {
        const fld = el('div', { class: 'field' });
        fld.appendChild(el('label', { text: c.label }));
        let input;
        const v = bug[c.key] ?? '';
        if (c.type === 'select') {
          input = el('select');
          c.options.forEach(o => {
            const op = el('option', { value: o, text: o || '(なし)' });
            if (o === v) op.selected = true;
            input.appendChild(op);
          });
        } else if (c.type === 'textarea') {
          input = el('textarea', { rows: '3' });
          input.value = String(v);
        } else if (c.type === 'date') {
          input = el('input', { type: 'date' });
          input.value = String(v).slice(0, 10);
        } else if (c.type === 'readonly') {
          input = el('input', { type: 'text', readonly: 'readonly' });
          input.value = String(v);
        } else {
          input = el('input', { type: 'text' });
          input.value = String(v);
        }
        input.dataset.key = c.key;
        fld.appendChild(input);
        wrap.appendChild(fld);
      });
      body.appendChild(wrap);
    });

    $('#modal').classList.remove('hidden');
  }
  function closeModal() {
    $('#modal').classList.add('hidden');
    state.editingRow = null;
  }
  async function saveModal() {
    const bug = state.bugs.find(b => b.rowIndex === state.editingRow);
    if (!bug) { closeModal(); return; }
    $('#modal-body').querySelectorAll('[data-key]').forEach(inp => {
      const k = inp.dataset.key;
      const col = COLUMNS.find(c => c.key === k);
      if (!col || col.type === 'readonly') return;
      bug[k] = inp.value;
    });
    try {
      await saveBugToExcel(bug);
      render();
      closeModal();
    } catch (e) {
      console.error(e);
      setStatus('保存失敗: ' + (e.message || e));
    }
  }

  function setView(v) {
    state.view = v;
    $('#btn-view-list').classList.toggle('active',   v === 'list');
    $('#btn-view-kanban').classList.toggle('active', v === 'kanban');
    $('#view-list').classList.toggle('active',   v === 'list');
    $('#view-kanban').classList.toggle('active', v === 'kanban');
    $('#kanban-controls').classList.toggle('hidden', v !== 'kanban');
    render();
  }
  function render() {
    if (state.view === 'list') renderList();
    else renderKanban();
  }

  function demoData() {
    return [
      { rowIndex: 4, id: 1, title: 'ログイン後に画面が真っ白', status: '解析中', updated: '2025-04-10', assignee: '高橋',
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
    $('#btn-view-list').addEventListener('click',   () => setView('list'));
    $('#btn-view-kanban').addEventListener('click', () => setView('kanban'));
    $('#btn-reload').addEventListener('click', async () => {
      await loadFromExcel();
      render();
    });
    $('#kanban-group').addEventListener('change', (e) => {
      state.kanbanGroup = e.target.value;
      render();
    });
    $('#filter-text').addEventListener('input', (e) => { state.filters.text = e.target.value; render(); });
    $('#filter-priority').addEventListener('change', (e) => { state.filters.priority = e.target.value; render(); });
    $('#filter-status').addEventListener('change',   (e) => { state.filters.status   = e.target.value; render(); });

    $('#modal-close').addEventListener('click', closeModal);
    $('#modal-cancel').addEventListener('click', closeModal);
    $('#modal-save').addEventListener('click', saveModal);
    $('#modal').addEventListener('click', (e) => { if (e.target.id === 'modal') closeModal(); });
  }

  ensureOffice(async () => {
    bindEvents();
    try {
      await loadFromExcel();
    } catch (e) {
      console.error(e);
      setStatus('読み込み失敗: ' + (e.message || e));
      state.bugs = demoData();
    }
    render();
  });
})();