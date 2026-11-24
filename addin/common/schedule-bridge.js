/*!
 * schedule-bridge.js
 * 期スケジュール（HTML）⇄ Excel「計画シート」連携のホスト側コード。
 *
 * ・ホスト＝作業ウィンドウ、またはリボンのコマンド関数のページ（Excel.run を呼べる側）。
 * ・ポップアップ（ダイアログ）側のHTMLは Excel を直接触れないため、メッセージでこのコードに依頼する。
 * ・保存は「変更されたセルだけ」を書き込む（3-way マージ）。Excelの共同編集で他の人が変えたセルは上書きしない。
 * ・他の人／自分がExcelのセルを変えたら、読み直して差分があればダイアログへ送る。
 *
 * 使い方（ホスト側のページから）:
 *   ScheduleBridge.open({ url: 'https://<アドインと同じドメイン>/ki-schedule.html' });
 */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.ScheduleBridge = api;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this), function () {
  'use strict';

  /* ================= 定数・シート構成 ================= */
  const TASK_COLS = ['ID', '種別', '親ID', '取引先等', '予定名', '開始', '終了', '継続', '入力単位', '入力値', '投入人数', '担当者', 'メモ', '順', 'コメント'];
  const MASTER_COLS = ['種別', 'ID', '値1', '値2', '値3'];
  const DATE_COLS = ['開始', '終了'];
  const UNIT_LABEL = { total: '人日（合計）', day: '人日/日', month: '人日/月' };
  const LABEL_UNIT = { '人日（合計）': 'total', '人日(合計)': 'total', '人日/日': 'day', '人日/月': 'month' };
  const BASIS_KEY = { total: 'pd', day: 'rate', month: 'mrate' };
  const ID_KINDS = ['取引先等', '要員', '決定事項', '宿題'];

  const uid = () => Math.random().toString(36).slice(2, 10);
  const pad = (n) => String(n).padStart(2, '0');
  const r3 = (v) => Math.round((Number(v) || 0) * 1000) / 1000;
  const norm = (v) => (v == null ? '' : typeof v === 'number' ? String(Math.round(v * 1000) / 1000) : typeof v === 'boolean' ? (v ? '1' : '0') : String(v).trim());
  /* コメント：「日時｜名前｜日付｜本文」を改行でつなぐ（アプリのCSVと同じ形式） */
  const hid = (s) => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return 'c' + (h >>> 0).toString(36); };
  const isDay = (v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v || ''));
  const cmLine = (c) => [c.at || '', c.by || '', c.date || '', String(c.text || '').replace(/\r?\n/g, '⏎').replace(/｜/g, '|')].join('｜');
  const cmStr = (list) => (list || []).map(cmLine).join('\n');
  function cmParse(str) {
    return String(str == null ? '' : str).split(/\r?\n/).map((line) => {
      if (!line.trim()) return null;
      const p = line.split('｜');
      if (p.length < 4) return { id: hid(line), at: '', by: '', date: '', text: line.trim() };
      const text = p.slice(3).join('｜').replace(/⏎/g, '\n');
      return text.trim() ? { id: hid(line), at: p[0], by: p[1], date: isDay(p[2]) ? p[2] : '', text } : null;
    }).filter(Boolean);
  }
  const truthy = (v) => { const s = norm(v); return s === '○' || s === '◯' || s === '1' || /^(true|yes|y)$/i.test(s); };

  /* ================= 日付（Excelのシリアル値 ⇄ ISO） ================= */
  function serialToIso(n) {
    const ms = Math.round((Math.floor(n) - 25569) * 86400000), d = new Date(ms);
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }
  function isoToSerial(iso) { const p = String(iso).split('-').map(Number); return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000) + 25569; }
  function cellToIso(v, fy) {
    if (v === '' || v == null) return '';
    if (typeof v === 'number') return v > 20000 ? serialToIso(v) : '';
    const t = String(v).trim();
    let m = /^(\d{4})[-\/.年](\d{1,2})[-\/.月](\d{1,2})/.exec(t);
    if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    m = /^(\d{1,2})[\/月](\d{1,2})/.exec(t);
    if (m) { const mo = +m[1]; return (mo >= 10 ? fy : fy + 1) + '-' + pad(mo) + '-' + pad(+m[2]); }
    return '';
  }

  /* ================= アプリのplan ⇄ シート行 ================= */
  function planToTables(plan) {
    const laneName = new Map(plan.lanes.map((l) => [l.id, l.name])), memName = new Map((plan.members || []).map((m) => [m.id, m.name]));
    const tasks = plan.tasks.map((t) => {
      const basis = t.ongoing && t.basis === 'total' ? 'day' : (t.basis || 'total');
      return {
        'ID': t.id, '種別': t.parentId ? '詳細予定' : '予定', '親ID': t.parentId || '', '取引先等': laneName.get(t.laneId) || '', '予定名': t.name,
        '開始': t.start, '終了': t.ongoing ? '' : t.end, '継続': t.ongoing ? '○' : '', '入力単位': UNIT_LABEL[basis] || UNIT_LABEL.total,
        '入力値': r3(t[BASIS_KEY[basis]]), '投入人数': r3(t.people), '担当者': (t.assignees || []).map((a) => memName.get(a)).filter(Boolean).join('・'), 'メモ': t.note || '', '順': t.seq != null ? r3(t.seq) : '', 'コメント': cmStr(t.comments)
      };
    });
    const master = [];
    const m = (kind, id, v1, v2, v3) => master.push({ '種別': kind, 'ID': id, '値1': v1 == null ? '' : v1, '値2': v2 == null ? '' : v2, '値3': v3 == null ? '' : v3 });
    m('期', 'fy', plan.fy);
    m('計画名', 'name', plan.name);
    plan.lanes.forEach((l) => m('取引先等', l.id, l.name));
    (plan.members || []).forEach((x) => m('要員', x.id, x.name));
    const st = plan.settings || {};
    m('設定', 'excludeHolidays', st.excludeHolidays === false ? 0 : 1);
    m('設定', 'yearEnd', st.yearEnd === false ? 0 : 1);
    (plan.decisions || []).forEach((d) => m('決定事項', d.id, d.text, d.taskId || ''));
    (plan.todos || []).forEach((t) => m('宿題', t.id, t.text, t.done ? 1 : 0));
    return { tasks, master };
  }

  function tablesToPlan(tbl, fyHint) {
    const g = (kind, id) => tbl.master.find((r) => norm(r['種別']) === kind && norm(r['ID']) === id);
    const fy = Number(norm((g('期', 'fy') || {})['値1'])) || fyHint;
    const plan = { fy, name: norm((g('計画名', 'name') || {})['値1']) || ((fy - 1988) + '期'), lanes: [], members: [], tasks: [], decisions: [], todos: [], settings: { excludeHolidays: true, yearEnd: true } };
    tbl.master.forEach((r) => {
      const k = norm(r['種別']), id = norm(r['ID']), v1 = norm(r['値1']), v2 = norm(r['値2']);
      if (k === '取引先等') plan.lanes.push({ id, name: v1 || '取引先' });
      else if (k === '要員') plan.members.push({ id, name: v1 || '要員' });
      else if (k === '設定') { if (id === 'excludeHolidays') plan.settings.excludeHolidays = v1 !== '0'; if (id === 'yearEnd') plan.settings.yearEnd = v1 !== '0'; }
      else if (k === '決定事項') plan.decisions.push({ id, text: v1, taskId: v2 });
      else if (k === '宿題') plan.todos.push({ id, text: v1, done: truthy(r['値2']) });
    });
    const laneId = (name) => { name = name || '取引先'; let l = plan.lanes.find((x) => x.name === name); if (!l) { l = { id: uid(), name }; plan.lanes.push(l); } return l.id; };
    const memberId = (name) => { let x = plan.members.find((y) => y.name === name); if (!x) { x = { id: uid(), name }; plan.members.push(x); } return x.id; };
    tbl.tasks.forEach((r) => {
      const name = norm(r['予定名']), start = cellToIso(r['開始'], fy);
      if (!name || !start) return;
      const ongoing = truthy(r['継続']);
      let end = cellToIso(r['終了'], fy);
      if (ongoing) end = (fy + 1) + '-09-30'; else if (!end) end = start;
      let basis = LABEL_UNIT[norm(r['入力単位'])] || (ongoing ? 'day' : 'total');
      if (ongoing && basis === 'total') basis = 'day';
      const val = Number(norm(r['入力値'])), ppl = norm(r['投入人数']);
      const t = {
        id: norm(r['ID']), name, laneId: laneId(norm(r['取引先等'])), parentId: norm(r['親ID']), start, end, ongoing, basis,
        people: ppl !== '' && isFinite(Number(ppl)) ? Number(ppl) : 1,
        assignees: norm(r['担当者']).split(/[・,、;；]/).map((x) => x.trim()).filter(Boolean).map(memberId), note: norm(r['メモ']),
        seq: (norm(r['順']) !== '' && isFinite(Number(norm(r['順'])))) ? Number(norm(r['順'])) : null,
        comments: cmParse(r['コメント'])
      };
      const v = isFinite(val) ? val : 0;
      if (basis === 'total') t.pd = v; else if (basis === 'day') t.rate = v; else t.mrate = v;
      plan.tasks.push(t);
    });
    return plan;
  }

  /* ================= 差分（セル単位） ================= */
  const keyTask = (r) => norm(r['ID']);
  const keyMaster = (r) => norm(r['種別']) + '|' + norm(r['ID']);
  function diffOne(base, next, keyOf, cols) {
    const bm = new Map(base.map((r) => [keyOf(r), r])), nm = new Map(next.map((r) => [keyOf(r), r]));
    const added = [], removed = [], changed = [];
    nm.forEach((r, k) => {
      const b = bm.get(k);
      if (!b) added.push(r);
      else cols.forEach((c) => { if (norm(b[c]) !== norm(r[c])) changed.push({ key: k, col: c, value: r[c] }); });
    });
    bm.forEach((r, k) => { if (!nm.has(k)) removed.push(k); });
    return { added, removed, changed };
  }
  function diffTables(base, next) { return { tasks: diffOne(base.tasks, next.tasks, keyTask, TASK_COLS), master: diffOne(base.master, next.master, keyMaster, MASTER_COLS) }; }
  function isEmptyDelta(d) { return ['tasks', 'master'].every((k) => !d[k].added.length && !d[k].removed.length && !d[k].changed.length); }
  function tablesEqual(a, b) { return isEmptyDelta(diffTables(a, b)); }

  /* ================= 同期エンジン（アダプタ経由でシートを読み書き） ================= */
  class Engine {
    constructor(adapter, send, opts) {
      opts = opts || {};
      this.ad = adapter; this.send = send; this.fyHint = Number(opts.fy) || new Date().getFullYear();
      this.onError = opts.onError || function () {};
      this.rev = 0; this.pending = new Map(); this.dialogBase = null; this.sheetBase = null; this.q = Promise.resolve();
    }
    _enqueue(fn) {
      this.q = this.q.then(fn).catch((e) => { this.onError(e); this.send({ t: 'error', message: (e && e.message) || String(e) }); });
      return this.q;
    }
    init(fy) {
      return this._enqueue(async () => {
        if (fy) this.fyHint = Number(fy) || this.fyHint;
        await this.ad.ensure();
        const tbl = await this._read();
        this.sheetBase = tbl; this.dialogBase = null;
        this._sendPlan(tbl, 'init');
      });
    }
    ack(rev, applied) { if (applied && this.pending.has(rev)) this.dialogBase = this.pending.get(rev); this.pending.delete(rev); }
    save(plan) {
      return this._enqueue(async () => {
        const next = planToTables(plan), base = this.dialogBase || this.sheetBase;
        const delta = diffTables(base, next);
        if (!isEmptyDelta(delta)) await this.ad.apply(delta);
        this.dialogBase = next;
        const cur = await this._read();
        const differs = !tablesEqual(cur, next);
        this.sheetBase = cur;
        if (differs) this._sendPlan(cur, 'merge'); else this.send({ t: 'saved' });
      });
    }
    changed() {
      return this._enqueue(async () => {
        const cur = await this._read();
        if (this.sheetBase && tablesEqual(cur, this.sheetBase)) return;
        this.sheetBase = cur;
        this._sendPlan(cur, 'remote');
      });
    }
    async _read() {
      const raw = await this.ad.read();
      const fyRow = raw.master.find((r) => norm(r['種別']) === '期' && norm(r['ID']) === 'fy');
      const fy = Number(norm(fyRow && fyRow['値1'])) || this.fyHint;
      const fixT = [], fixM = [];
      raw.tasks.forEach((r) => { if (!norm(r['ID']) && (norm(r['予定名']) || norm(r['開始']))) { r['ID'] = uid(); fixT.push({ row: r._row, col: 'ID', value: r['ID'] }); } });
      raw.master.forEach((r) => { if (ID_KINDS.indexOf(norm(r['種別'])) >= 0 && !norm(r['ID'])) { r['ID'] = uid(); fixM.push({ row: r._row, col: 'ID', value: r['ID'] }); } });
      if (fixT.length) await this.ad.writeCells('tasks', fixT);
      if (fixM.length) await this.ad.writeCells('master', fixM);
      const tasks = raw.tasks.filter((r) => norm(r['ID'])).map((r) => {
        const o = Object.assign({}, r); DATE_COLS.forEach((c) => { o[c] = cellToIso(r[c], fy); }); return o;
      });
      const master = raw.master.filter((r) => norm(r['種別']));
      return { tasks, master };
    }
    _sendPlan(tbl, source) {
      const rev = ++this.rev; this.pending.set(rev, tbl);
      this.send({ t: 'plan', rev, source, plan: tablesToPlan(tbl, this.fyHint) });
    }
  }

  /* ================= Excel アダプタ（Office.js） ================= */
  function makeExcelAdapter(cfg) {
    cfg = cfg || {};
    const SPEC = {
      tasks: { name: cfg.sheetName || '計画シート', cols: TASK_COLS, key: keyTask },
      master: { name: cfg.masterName || '計画マスタ', cols: MASTER_COLS, key: keyMaster }
    };
    async function loadTable(ctx, spec) {
      const ws = ctx.workbook.worksheets.getItem(spec.name);
      const ur = ws.getUsedRangeOrNullObject(); ur.load('values,rowIndex,columnIndex');
      await ctx.sync();
      if (ur.isNullObject) return { ws, idx: {}, rows: [], r0: 0, c0: 0 };
      const header = ur.values[0].map((h) => String(h).trim()), idx = {};
      header.forEach((h, i) => { if (h && !(h in idx)) idx[h] = i; });
      return { ws, idx, rows: ur.values.slice(1), r0: ur.rowIndex, c0: ur.columnIndex };
    }
    function setCell(cell, col, value) {
      if (DATE_COLS.indexOf(col) >= 0 && /^\d{4}-\d{2}-\d{2}$/.test(String(value))) { cell.values = [[isoToSerial(value)]]; cell.numberFormat = [['yyyy/m/d']]; }
      else cell.values = [[value == null ? '' : value]];
    }
    const recOf = (spec, t, row) => { const rec = {}; spec.cols.forEach((c) => { const j = t.idx[c]; rec[c] = j == null ? '' : row[j]; }); return rec; };
    return {
      async ensure() {
        await Excel.run(async (ctx) => {
          for (const spec of Object.values(SPEC)) {
            let ws = ctx.workbook.worksheets.getItemOrNullObject(spec.name);
            await ctx.sync();
            if (ws.isNullObject) ws = ctx.workbook.worksheets.add(spec.name);
            const ur = ws.getUsedRangeOrNullObject(); ur.load('values,rowIndex,columnIndex');
            await ctx.sync();
            const header = ur.isNullObject ? [] : ur.values[0].map((h) => String(h).trim());
            const r0 = ur.isNullObject ? 0 : ur.rowIndex, c0 = ur.isNullObject ? 0 : ur.columnIndex;
            const missing = spec.cols.filter((c) => header.indexOf(c) < 0);
            if (missing.length) { const rng = ws.getRangeByIndexes(r0, c0 + header.length, 1, missing.length); rng.values = [missing]; rng.format.font.bold = true; }
            await ctx.sync();
          }
        });
      },
      async read() {
        return Excel.run(async (ctx) => {
          const out = {};
          for (const k of Object.keys(SPEC)) {
            const spec = SPEC[k], t = await loadTable(ctx, spec);
            out[k] = t.rows.map((row, i) => Object.assign({ _row: t.r0 + i + 2 }, recOf(spec, t, row))).filter((rec) => spec.cols.some((c) => norm(rec[c]) !== ''));
          }
          return out;
        });
      },
      async writeCells(which, cells) {
        if (!cells.length) return;
        await Excel.run(async (ctx) => {
          const t = await loadTable(ctx, SPEC[which]);
          cells.forEach((c) => { const j = t.idx[c.col]; if (j != null) setCell(t.ws.getCell(c.row - 1, t.c0 + j), c.col, c.value); });
          await ctx.sync();
        });
      },
      async apply(delta) {
        await Excel.run(async (ctx) => {
          for (const k of Object.keys(SPEC)) {
            const d = delta[k]; if (!d || (!d.added.length && !d.removed.length && !d.changed.length)) continue;
            const spec = SPEC[k], t = await loadTable(ctx, spec), rowOf = new Map();
            t.rows.forEach((row, i) => { const key = spec.key(recOf(spec, t, row)); if (key && key !== '|' && !rowOf.has(key)) rowOf.set(key, i); });
            d.changed.forEach((ch) => { const i = rowOf.get(ch.key), j = t.idx[ch.col]; if (i != null && j != null) setCell(t.ws.getCell(t.r0 + 1 + i, t.c0 + j), ch.col, ch.value); });
            let next = t.r0 + 1 + t.rows.length;
            d.added.forEach((rec) => { spec.cols.forEach((c) => { const j = t.idx[c]; if (j != null) setCell(t.ws.getCell(next, t.c0 + j), c, rec[c]); }); next++; });
            d.removed.map((key) => rowOf.get(key)).filter((i) => i != null).sort((a, b) => b - a)
              .forEach((i) => { t.ws.getRangeByIndexes(t.r0 + 1 + i, 0, 1, 1).getEntireRow().delete('Up'); });
            await ctx.sync();
          }
        });
      },
      async onChanged(cb) {
        const regs = []; let timer = null;
        const trigger = () => { clearTimeout(timer); timer = setTimeout(cb, 400); };
        await Excel.run(async (ctx) => {
          for (const spec of Object.values(SPEC)) { regs.push(ctx.workbook.worksheets.getItem(spec.name).onChanged.add(async () => { trigger(); })); }
          await ctx.sync();
        });
        return async () => {
          clearTimeout(timer);
          for (const r of regs) { try { await Excel.run(r.context, async (ctx) => { r.remove(); await ctx.sync(); }); } catch (_) {} }
        };
      }
    };
  }

  /* ================= ダイアログを開いて中継する ================= */
  function open(opts) {
    opts = Object.assign({ url: '', width: 95, height: 92, sheetName: '計画シート', masterName: '計画マスタ', fy: null, displayInIframe: true, onClose: null, onError: null }, opts || {});
    if (!opts.url) throw new Error('url（スケジュールHTMLのURL）を指定してください');
    const adapter = makeExcelAdapter(opts);
    let dialog = null, engine = null, unsub = null, closed = false;
    const fail = (e) => { if (opts.onError) opts.onError(e); else if (typeof console !== 'undefined') console.error(e); };
    const send = (o) => { try { dialog.messageChild(JSON.stringify(o)); } catch (e) { fail(e); } };
    const cleanup = async () => { if (closed) return; closed = true; try { if (unsub) await unsub(); } catch (_) {} if (opts.onClose) opts.onClose(); };
    const url = opts.url + (opts.url.indexOf('?') < 0 ? '?' : '&') + 'excel=1';
    Office.context.ui.displayDialogAsync(url, { width: opts.width, height: opts.height, displayInIframe: !!opts.displayInIframe }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { fail(res.error); cleanup(); return; }
      dialog = res.value;
      engine = new Engine(adapter, send, { fy: opts.fy, onError: fail });
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let m; try { m = JSON.parse(arg.message); } catch (_) { return; }
        if (m.t === 'ready') engine.init(m.fy).then(async () => { if (!unsub) unsub = await adapter.onChanged(() => engine.changed()); }).catch(fail);
        else if (m.t === 'save') engine.save(m.plan);
        else if (m.t === 'applied') engine.ack(m.rev, true);
        else if (m.t === 'skipped') engine.ack(m.rev, false);
        else if (m.t === 'close') { try { dialog.close(); } catch (_) {} cleanup(); }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { cleanup(); });
    });
    return { close() { try { if (dialog) dialog.close(); } catch (_) {} cleanup(); } };
  }

  return { open, Engine, makeExcelAdapter, Core: { planToTables, tablesToPlan, diffTables, tablesEqual, isEmptyDelta, cmStr, cmParse, cellToIso, isoToSerial, serialToIso, TASK_COLS, MASTER_COLS, DATE_COLS, keyTask, keyMaster, norm } };
});
