/* 情報漏洩対策集 連携（Excelアドイン）
 * Webページ（情報漏洩対策集）と同じJSON形式を読み書きします。
 * データの原本は非表示シート「_data」にJSONで保持し、
 * 人が見る・入力するシートはそこから毎回作り直します。
 */
(function () {
  'use strict';

  var SH = { summary: 'サマリ', measures: '対策一覧', targets: '取引先', judge: '対象判定', results: '点検結果', ng: '要対応', data: '_data' };
  var ORDER = [SH.summary, SH.measures, SH.targets, SH.judge, SH.results, SH.ng, SH.data];
  var MTYPES = ['予防', '検知', '事後対応'];
  var RES_LABEL = { ok: 'OK', ng: '要対応', na: '対象外' };
  var RES_CODE = { 'OK': 'ok', 'ok': 'ok', '要対応': 'ng', '対象外': 'na' };
  var OUT_LABEL = '対象外（起因）';
  var COL = { head: '#1D4E89', white: '#FFFFFF', ok: '#1D4E89', okWeak: '#D6E3F3', ng: '#C2501A', ngWeak: '#FBEADF', ngInk: '#9A3A0E', na: '#8D96A6', none: '#E6E9EE', sub: '#4A5263' };
  var CHUNK = 30000; // 1セル32,767文字の上限より小さく分割

  function $(id) { return document.getElementById(id); }
  function setStatus(msg, kind) { var el = $('status'); el.textContent = msg; el.className = 'status ' + (kind || ''); }
  function today() { var d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function toDate(v) {
    if (typeof v === 'number' && v > 20000) { var d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
    return str(v).replace(/\//g, '-');
  }

  /* ---------- データ（Webページと同じロジック） ---------- */
  function normalize(d) {
    if (!d || !Array.isArray(d.causes) || !Array.isArray(d.targets)) throw new Error('情報漏洩対策集のJSONではありません（causes / targets が見つかりません）。');
    d.results = d.results || {}; d.off = d.off || {}; d.scopeA = d.scopeA || {};
    d.causes.forEach(function (c) {
      c.scopeQs = c.scopeQs || []; c.checks = c.checks || []; c.reasons = c.reasons || [];
      c.measures = c.measures || {}; c.refs = c.refs || []; c.history = c.history || [];
    });
    return d;
  }
  function manOff(S, t, cid) { var o = S.off[t]; return o && Object.prototype.hasOwnProperty.call(o, cid) ? o[cid] : null; }
  function judge(S, t, c) {
    var qs = c.scopeQs || [];
    if (!qs.length) { var m = manOff(S, t, c.id); return m !== null ? { s: 'out', r: m } : { s: 'in' }; }
    var a = S.scopeA[t] || {}, y = 0, n = 0;
    qs.forEach(function (q) { if (a[q.id] === 'y') y++; else if (a[q.id] === 'n') n++; });
    if (y) return { s: 'in' };
    if (n === qs.length) { var m2 = manOff(S, t, c.id); return { s: 'out', r: m2 ? m2 : '設問の回答がすべて「いいえ」' }; }
    return { s: 'pending' };
  }
  var JUDGE_LABEL = { 'in': '点検対象', out: '対象外', pending: '判定待ち' };
  function stat(S, t, c) {
    var o = { ok: 0, ng: 0, na: 0, none: 0, total: c.checks.length };
    if (judge(S, t, c).s === 'out') { o.na = o.total; return o; }
    c.checks.forEach(function (k) { var r = (S.results[t] || {})[k.id]; if (!r || !r.s) o.none++; else o[r.s]++; });
    return o;
  }
  function inTargets(S) { return S.targets.filter(function (t) { return t.inScope; }); }

  /* ---------- 書き込みヘルパー ---------- */
  function grid(ws, rows, r0, c0, asText) {
    var rg = ws.getRangeByIndexes(r0 || 0, c0 || 0, rows.length, rows[0].length);
    if (asText !== false) rg.numberFormat = rows.map(function (r) { return r.map(function () { return '@'; }); });
    rg.values = rows;
    return rg;
  }
  function headStyle(rg) { rg.format.fill.color = COL.head; rg.format.font.color = COL.white; rg.format.font.bold = true; }
  function widths(ws, list) { list.forEach(function (w, i) { ws.getRangeByIndexes(0, i, 1, 1).getEntireColumn().format.columnWidth = w; }); }
  function table(ws, rows, name) {
    var rg = grid(ws, rows, 0, 0);
    if (rows.length > 1) { var t = ws.tables.add(rg, true); t.name = name; t.style = 'TableStyleLight9'; }
    else headStyle(rg);
    ws.freezePanes.freezeRows(1);
    return rg;
  }
  function listRule(ws, col, n, source) {
    if (n < 1) return;
    var rg = ws.getRangeByIndexes(1, col, n, 1);
    rg.dataValidation.clear();
    rg.dataValidation.rule = { list: { inCellDropDown: true, source: source } };
  }
  function textCF(rg, text, fill, font) {
    var cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    cf.textComparison.format.fill.color = fill;
    if (font) cf.textComparison.format.font.color = font;
    cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: text };
  }

  /* ---------- 各シート ---------- */
  function sheetMeasures(ws, S) {
    var rows = [['起因ID', '起因', '重要度', '発生しやすさ', '責任分界', '概要', '対象判定の設問', '原因', '対策：予防', '対策：検知', '対策：事後対応', '点検方法（頻度）', '関連事例・参考資料', '最終更新']];
    S.causes.forEach(function (c) {
      var h = c.history[c.history.length - 1];
      rows.push([c.id, c.name, c.severity || '', c.likelihood || '', c.owner || '', c.summary || '',
        c.scopeQs.map(function (q) { return '・' + q.text; }).join('\n'),
        c.reasons.map(function (r) { return '・' + r; }).join('\n')]
        .concat(MTYPES.map(function (t) { return (c.measures[t] || []).map(function (m) { return '・' + m; }).join('\n'); }))
        .concat([c.checks.map(function (k, i) { return (i + 1) + '. ' + k.text + '（' + k.freq + '）'; }).join('\n'),
          c.refs.map(function (r) { return r.url ? r.label + ' ' + r.url : r.label; }).join('\n'),
          h ? h.date + ' ' + h.note : '']));
    });
    var rg = table(ws, rows, 'tblMeasures');
    rg.format.wrapText = true; rg.format.verticalAlignment = 'Top';
    widths(ws, [70, 150, 50, 70, 90, 240, 240, 240, 220, 200, 200, 280, 200, 140]);
  }

  function sheetTargets(ws, S) {
    var rows = [['取引先', '点検対象（はい／いいえ）']];
    S.targets.forEach(function (t) { rows.push([t.name, t.inScope ? 'はい' : 'いいえ']); });
    table(ws, rows, 'tblTargets');
    listRule(ws, 1, rows.length - 1, 'はい,いいえ');
    widths(ws, [200, 150]);
  }

  function sheetJudge(ws, S) {
    var rows = [['取引先', '起因ID', '起因', '設問ID', '設問', '回答（はい／いいえ）', '判定（集計更新時）']];
    inTargets(S).forEach(function (t) {
      var a = S.scopeA[t.name] || {};
      S.causes.forEach(function (c) {
        var j = JUDGE_LABEL[judge(S, t.name, c).s];
        if (c.scopeQs.length) c.scopeQs.forEach(function (q) {
          rows.push([t.name, c.id, c.name, q.id, q.text, a[q.id] === 'y' ? 'はい' : a[q.id] === 'n' ? 'いいえ' : '', j]);
        });
        else rows.push([t.name, c.id, c.name, 'manual', '（設問なし）この起因を点検するか', manOff(S, t.name, c.id) !== null ? 'いいえ' : 'はい', j]);
      });
    });
    var rg = table(ws, rows, 'tblJudge');
    rg.format.wrapText = true; rg.format.verticalAlignment = 'Top';
    listRule(ws, 5, rows.length - 1, 'はい,いいえ');
    if (rows.length > 1) { var jr = ws.getRangeByIndexes(1, 6, rows.length - 1, 1); textCF(jr, '判定待ち', '#FBF1D6', '#6E4D00'); textCF(jr, '対象外', COL.none, COL.sub); }
    widths(ws, [160, 60, 150, 90, 340, 120, 110]);
    ws.getRangeByIndexes(0, 3, 1, 1).getEntireColumn().columnHidden = true;
  }

  function sheetResults(ws, S) {
    var rows = [['取引先', '起因ID', '起因', '判定', '点検項目ID', '点検項目', '頻度', '結果', 'メモ・証跡・期限', '点検日']];
    inTargets(S).forEach(function (t) {
      var m = S.results[t.name] || {};
      S.causes.forEach(function (c) {
        var j = judge(S, t.name, c);
        c.checks.forEach(function (k) {
          var r = m[k.id] || {};
          if (j.s === 'out') rows.push([t.name, c.id, c.name, '対象外', k.id, k.text, k.freq, OUT_LABEL, j.r || '', '']);
          else rows.push([t.name, c.id, c.name, JUDGE_LABEL[j.s], k.id, k.text, k.freq, RES_LABEL[r.s] || '未点検', r.m || '', r.d || '']);
        });
      });
    });
    var rg = table(ws, rows, 'tblResults');
    rg.format.wrapText = true; rg.format.verticalAlignment = 'Top';
    listRule(ws, 7, rows.length - 1, 'OK,要対応,対象外,未点検');
    if (rows.length > 1) {
      var rr = ws.getRangeByIndexes(1, 7, rows.length - 1, 1);
      textCF(rr, '要対応', COL.ngWeak, COL.ngInk); textCF(rr, 'OK', COL.okWeak, '#123459'); textCF(rr, '対象外', COL.none, COL.sub);
    }
    widths(ws, [160, 60, 150, 70, 80, 340, 60, 80, 240, 90]);
    ws.getRangeByIndexes(0, 4, 1, 1).getEntireColumn().columnHidden = true;
  }

  function sheetNg(ws, S) {
    var rows = [['取引先', '起因', '点検項目', 'メモ・証跡・期限', '点検日']];
    inTargets(S).forEach(function (t) {
      S.causes.forEach(function (c) {
        if (judge(S, t.name, c).s === 'out') return;
        c.checks.forEach(function (k) { var r = (S.results[t.name] || {})[k.id]; if (r && r.s === 'ng') rows.push([t.name, c.name, k.text, r.m || '', r.d || '']); });
      });
    });
    if (rows.length === 1) rows.push(['要対応の項目はありません', '', '', '', '']);
    var rg = table(ws, rows, 'tblNg');
    rg.format.wrapText = true; rg.format.verticalAlignment = 'Top';
    widths(ws, [160, 150, 340, 260, 90]);
  }

  function sheetSummary(ws, S) {
    var T = inTargets(S), o = { ok: 0, ng: 0, na: 0, none: 0, total: 0 }, done = 0;
    var per = S.causes.map(function (c) {
      var a = { ok: 0, ng: 0, na: 0, none: 0 };
      T.forEach(function (t) { var s = stat(S, t.name, c); a.ok += s.ok; a.ng += s.ng; a.na += s.na; a.none += s.none; });
      o.ok += a.ok; o.ng += a.ng; o.na += a.na; o.none += a.none; return a;
    });
    o.total = o.ok + o.ng + o.na + o.none;
    T.forEach(function (t) { if (S.causes.every(function (c) { return stat(S, t.name, c).none === 0; })) done++; });
    var pct = o.total ? Math.round((o.total - o.none) / o.total * 100) : 0;

    var title = ws.getRange('A1'); title.values = [['情報漏洩対策集 サマリ']]; title.format.font.size = 16; title.format.font.bold = true;
    grid(ws, [['集計日時：' + new Date().toLocaleString('ja-JP') + '（「集計を更新」で最新化）']], 1, 0).format.font.color = COL.sub;

    var k = grid(ws, [['点検対象の取引先', T.length + ' / ' + S.targets.length + ' 社'], ['点検実施率', pct + '%（完了 ' + done + ' 社）'], ['要対応の項目', String(o.ng)], ['未点検の項目', String(o.none)]], 3, 0);
    k.format.font.size = 12;
    ws.getRange('A4:A7').format.font.bold = true;
    ws.getRange('B6').format.font.color = COL.ngInk; ws.getRange('B6').format.font.bold = true;

    var r0 = 8;
    var rows = [['起因', 'OK', '要対応', '対象外', '未点検']].concat(S.causes.map(function (c, i) { return [c.name, per[i].ok, per[i].ng, per[i].na, per[i].none]; }));
    var cr = ws.getRangeByIndexes(r0, 0, rows.length, 5);
    ws.getRangeByIndexes(r0, 0, rows.length, 1).numberFormat = rows.map(function () { return ['@']; });
    cr.values = rows;
    headStyle(ws.getRangeByIndexes(r0, 0, 1, 5));
    if (rows.length > 1) {
      var ch = ws.charts.add(Excel.ChartType.barStacked, cr, Excel.ChartSeriesBy.columns);
      ch.title.text = '起因別の点検状況（点検対象の全取引先）';
      ch.setPosition('G2', 'O' + (r0 + rows.length + 2));
      [COL.ok, COL.ng, COL.na, COL.none].forEach(function (color, i) { ch.series.getItemAt(i).format.fill.setSolidColor(color); });
    }

    var m0 = r0 + rows.length + 3;
    grid(ws, [['取引先 × 起因（セルは 点検済み／項目数、要対応があれば「要 n」）']], m0, 0).format.font.bold = true;
    var head = ['取引先'].concat(S.causes.map(function (c) { return c.name; }), ['要対応計']);
    var mrows = [head], fills = [];
    T.forEach(function (t) {
      var row = [t.name], fr = [null], tot = 0;
      S.causes.forEach(function (c) {
        var j = judge(S, t.name, c).s, s = stat(S, t.name, c), d = s.total - s.none; tot += s.ng;
        if (j === 'pending') { row.push('判定待ち'); fr.push([COL.none, COL.sub]); }
        else if (j === 'out') { row.push('対象外'); fr.push(['#FFFFFF', COL.sub]); }
        else if (s.ng) { row.push('要 ' + s.ng); fr.push([COL.ng, COL.white]); }
        else if (d === 0) { row.push('0/' + s.total); fr.push([COL.none, COL.sub]); }
        else if (d === s.total) { row.push(d + '/' + s.total); fr.push([COL.ok, COL.white]); }
        else { row.push(d + '/' + s.total); fr.push([COL.okWeak, '#161B26']); }
      });
      row.push(tot ? String(tot) : ''); fr.push(null);
      mrows.push(row); fills.push(fr);
    });
    if (mrows.length === 1) mrows.push([ '点検対象の取引先がありません' ].concat(head.slice(1).map(function () { return ''; })));
    var mr = grid(ws, mrows, m0 + 1, 0);
    headStyle(ws.getRangeByIndexes(m0 + 1, 0, 1, head.length));
    ws.getRangeByIndexes(m0 + 1, 1, 1, head.length - 1).format.wrapText = true;
    ws.getRangeByIndexes(m0 + 2, 1, mrows.length - 1, head.length - 1).format.horizontalAlignment = 'Center';
    fills.forEach(function (fr, i) {
      fr.forEach(function (f, j) {
        if (!f) return;
        var cell = ws.getCell(m0 + 2 + i, j);
        cell.format.fill.color = f[0]; cell.format.font.color = f[1];
      });
    });
    mr.format.borders.getItem('InsideHorizontal').color = '#FFFFFF';
    mr.format.borders.getItem('InsideVertical').color = '#FFFFFF';
    widths(ws, [180].concat(S.causes.map(function () { return 86; }), [70]));
    ws.getRange('A1:A7').format.wrapText = false;
  }

  function sheetData(ws, S) {
    var json = JSON.stringify(S), rows = [];
    for (var i = 0; i < json.length; i += CHUNK) rows.push([json.slice(i, i + CHUNK)]);
    grid(ws, rows, 0, 0);
    ws.visibility = Excel.SheetVisibility.veryHidden;
  }

  /* ---------- ブック全体を作り直す ---------- */
  function writeAll(S) {
    return Excel.run(function (ctx) {
      var wss = ctx.workbook.worksheets;
      var tmp = wss.add('_tmp' + (Date.now() % 1000000));
      var olds = ORDER.map(function (n) { return wss.getItemOrNullObject(n); });
      return ctx.sync().then(function () {
        olds.forEach(function (w) { if (!w.isNullObject) w.delete(); });
        return ctx.sync();
      }).then(function () {
        var ws = {};
        ORDER.forEach(function (n) { ws[n] = wss.add(n); });
        tmp.delete();
        sheetMeasures(ws[SH.measures], S);
        sheetTargets(ws[SH.targets], S);
        sheetJudge(ws[SH.judge], S);
        sheetResults(ws[SH.results], S);
        sheetNg(ws[SH.ng], S);
        sheetSummary(ws[SH.summary], S);
        sheetData(ws[SH.data], S);
        ws[SH.summary].activate();
        return ctx.sync();
      });
    });
  }

  /* ---------- Excelの入力を読み取る ---------- */
  function indexer(header) {
    var h = header.map(str);
    return function (name) { var i = h.indexOf(name); if (i < 0) throw new Error('「' + name + '」列が見つかりません。列名を変えないでください。'); return i; };
  }
  function readBack() {
    return Excel.run(function (ctx) {
      var wss = ctx.workbook.worksheets;
      var sheets = {};
      [SH.data, SH.targets, SH.judge, SH.results].forEach(function (n) { sheets[n] = wss.getItemOrNullObject(n); });
      var used = {};
      return ctx.sync().then(function () {
        if (sheets[SH.data].isNullObject) throw new Error('このブックにはまだデータがありません。先に「1. Excelに保存する」を行ってください。');
        Object.keys(sheets).forEach(function (n) { if (!sheets[n].isNullObject) { used[n] = sheets[n].getUsedRange(); used[n].load('values'); } });
        return ctx.sync();
      }).then(function () {
        var S = normalize(JSON.parse(used[SH.data].values.map(function (r) { return r[0] == null ? '' : String(r[0]); }).join('')));

        if (used[SH.targets]) {
          var tv = used[SH.targets].values, ti = indexer(tv[0]), cN = ti('取引先'), cS = ti('点検対象（はい／いいえ）');
          tv.slice(1).forEach(function (r) {
            var name = str(r[cN]); if (!name) return;
            var inScope = str(r[cS]) !== 'いいえ';
            var t = S.targets.find(function (x) { return x.name === name; });
            if (t) t.inScope = inScope; else S.targets.push({ name: name, inScope: inScope });
          });
        }

        if (used[SH.judge]) {
          var jv = used[SH.judge].values, ji = indexer(jv[0]);
          var jT = ji('取引先'), jC = ji('起因ID'), jQ = ji('設問ID'), jA = ji('回答（はい／いいえ）');
          jv.slice(1).forEach(function (r) {
            var t = str(r[jT]), cid = str(r[jC]), qid = str(r[jQ]), ans = str(r[jA]);
            if (!t || !cid || !qid) return;
            if (qid === 'manual') {
              S.off[t] = S.off[t] || {};
              if (ans === 'いいえ') { if (manOff(S, t, cid) === null) S.off[t][cid] = ''; }
              else delete S.off[t][cid];
            } else {
              var a = S.scopeA[t] = S.scopeA[t] || {};
              if (ans === 'はい') a[qid] = 'y'; else if (ans === 'いいえ') a[qid] = 'n'; else delete a[qid];
            }
          });
        }

        if (used[SH.results]) {
          var rv = used[SH.results].values, ri = indexer(rv[0]);
          var rT = ri('取引先'), rK = ri('点検項目ID'), rR = ri('結果'), rM = ri('メモ・証跡・期限'), rD = ri('点検日');
          rv.slice(1).forEach(function (r) {
            var t = str(r[rT]), kid = str(r[rK]), lab = str(r[rR]);
            if (!t || !kid || lab === OUT_LABEL) return;
            var map = S.results[t] = S.results[t] || {}, prev = map[kid] || {};
            var code = RES_CODE[lab], memo = str(r[rM]), date = toDate(r[rD]), nr = {};
            if (code) { nr.s = code; nr.d = date || (prev.s === code && prev.d ? prev.d : today()); }
            if (memo) nr.m = memo;
            if (!nr.s && !nr.m) delete map[kid]; else map[kid] = nr;
          });
        }
        S.updated = new Date().toISOString();
        return S;
      });
    });
  }

  /* ---------- 画面操作 ---------- */
  function fail(err) {
    console.error(err, err && err.debugInfo);
    setStatus('エラー：' + (err && err.message ? err.message : String(err)), 'error');
  }
  function loadJson(text) {
    var S;
    try { S = normalize(JSON.parse(text)); } catch (e) { return Promise.reject(e instanceof SyntaxError ? new Error('JSONとして読み取れません。コピーした内容を確認してください。') : e); }
    setStatus('Excelに書き込んでいます…');
    return writeAll(S).then(function () {
      setStatus('保存しました：起因 ' + S.causes.length + ' 件、取引先 ' + S.targets.length + ' 社。', 'ok');
    });
  }
  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text).catch(function () { return legacyCopy(); });
    return legacyCopy();
    function legacyCopy() { var ta = $('out'); ta.focus(); ta.select(); var ok = document.execCommand('copy'); return ok ? Promise.resolve() : Promise.reject(new Error('コピーできませんでした。テキストを選択してコピーしてください。')); }
  }
  function download(text) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
    a.download = '情報漏洩対策集_' + today() + '.json';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  /* Excel起動時の自動読み込み（共有ランタイム）。
     一度手動で開くと「ブックを開いた時に読み込む」設定が登録され、次回以降は自動で起動する。
     このアドインのシート（対策一覧）があるブックでのみ作業ウィンドウを表示する。 */
  function isSharedRuntime() {
    try { return Office.context.requirements.isSetSupported('SharedRuntime', '1.1'); } catch (e) { return false; }
  }
  function registerStartupLoad() {
    if (!isSharedRuntime()) return Promise.resolve();
    return Promise.resolve().then(function () { return Office.addin.setStartupBehavior(Office.StartupBehavior.load); })
      .catch(function (e) { console.warn('setStartupBehavior failed:', e); });
  }
  function showPaneIfHasSheet() {
    if (!isSharedRuntime()) return Promise.resolve();
    var exists = false;
    return Excel.run(function (ctx) {
      var sh = ctx.workbook.worksheets.getItemOrNullObject(SH.measures);
      return ctx.sync().then(function () { exists = !sh.isNullObject; });
    }).then(function () { if (exists) return Office.addin.showAsTaskpane(); })
      .catch(function (e) { console.warn('showAsTaskpane failed:', e); });
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { setStatus('このアドインはExcelで使います。', 'error'); return; }
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.8')) { setStatus('このExcelでは一部の機能が使えません（ExcelApi 1.8 以上が必要です）。Microsoft 365 版のExcelを使ってください。', 'error'); return; }
    registerStartupLoad();
    showPaneIfHasSheet();
    ['btnPaste', 'btnRefresh', 'btnExport', 'fileIn'].forEach(function (id) { $(id).disabled = false; });
    setStatus('準備ができました。');

    $('btnPaste').onclick = function () {
      var t = $('pasteIn').value.trim();
      if (!t) { setStatus('JSONを貼り付けてください。', 'error'); return; }
      loadJson(t).then(function () { $('pasteIn').value = ''; }).catch(fail);
    };
    $('fileIn').onchange = function (e) {
      var f = e.target.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function () { loadJson(fr.result).catch(fail); e.target.value = ''; };
      fr.readAsText(f);
    };
    $('btnRefresh').onclick = function () {
      setStatus('入力内容を読み取っています…');
      readBack().then(function (S) { setStatus('集計を作り直しています…'); return writeAll(S); })
        .then(function () { setStatus('集計を更新しました。', 'ok'); }).catch(fail);
    };
    $('btnExport').onclick = function () {
      setStatus('JSONを作っています…');
      readBack().then(function (S) {
        var text = JSON.stringify(S);
        $('out').value = text; $('outWrap').hidden = false;
        return writeAll(S).then(function () { return copyText(text); })
          .then(function () { setStatus('JSONをコピーしました。Webページの「JSONを貼り付け」に貼り、「置き換えて読み込む」を押してください。', 'ok'); })
          .catch(function () { setStatus('JSONを作りました。下の欄の内容をコピーしてWebページに貼り付けてください。', 'ok'); });
      }).catch(fail);
    };
    $('btnCopy').onclick = function () { copyText($('out').value).then(function () { setStatus('コピーしました。', 'ok'); }).catch(fail); };
    $('btnDownload').onclick = function () { download($('out').value); };
  });
})();
