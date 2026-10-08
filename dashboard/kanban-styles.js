/* ============================================================
 * kanban-styles.js — WBSカンバンカードの表示スタイル（10種）
 * ------------------------------------------------------------
 * dashboard/index.html と dashboard/clipgen-preview.html の両方から読む。
 * 「バー（標準）」は各ページ側の既存の描画を使い、ここではそれ以外の10種を描く。
 *
 * 描くのは「件数の見せ方」だけ。タスク一覧（アコーディオン）の中身は
 * 呼び出し側の listHtml() に任せるので、一覧の見た目はどのスタイルでも同じ。
 *
 *   KanbanStyles.STYLES                 選べるスタイルの一覧 [{key,label}]
 *   KanbanStyles.normalize(key)         知らない値は 'bars' にする
 *   KanbanStyles.render(style, ctx)     HTML文字列（'bars' のときは null）
 *   KanbanStyles.pickerHtml(cur, attr)  切り替え用の <select>
 *
 * ctx = {
 *   rows,               classify() の戻り値（late/today/week/next/todo）
 *   defs,               行の定義（ROWS）
 *   today,              基準日（Date 0:00）
 *   expanded,           開いているキーの Set
 *   rowAttr(key),       data-kb-row に入れる値（index は "id:key"、見本は "key"）
 *   listHtml(key, label, r)  開いたときのタスク一覧のHTML
 *   esc(s)              HTMLエスケープ
 * }
 * 日別のキー（週ストリップ）は "d20260930" の形。r は { rest, done }。
 * ============================================================ */
(function () {
  'use strict';
  if (window.KanbanStyles) return;

  var VERSION = 'rev_20261008_kbs01';

  var STYLES = [
    { key: 'bars',     label: 'バー（標準）' },
    { key: 'ring',     label: 'リング' },
    { key: 'hero',     label: '数字を大きく' },
    { key: 'tiles',    label: 'タイル' },
    { key: 'dots',     label: 'ドットで1件ずつ' },
    { key: 'focus',    label: '今週を主役に' },
    { key: 'strip',    label: '期限日の週ストリップ' },
    { key: 'labeled',  label: '数字入りの積み上げ' },
    { key: 'minimal',  label: '罫線だけのミニマル' },
    { key: 'heat',     label: '緊急度の色タイル' },
    { key: 'timeline', label: '時間軸' }
  ];

  var STATUS_ORDER = ['doing', 'todo', 'held'];
  var STATUS_SHORT = { doing: '対応中', todo: '未着手', held: '保留', done: '完了' };
  var WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

  /* ---------- CSS（初回だけ差し込む。色はダッシュボードの変数を使う） ---------- */
  var CSS = [
    '.kbs{display:flex;flex-direction:column;gap:6px;--kbs-held:#E0B252;--kbs-held-ink:color-mix(in srgb,#E0B252 70%,var(--text));}',
    '.kbs-hit{appearance:none;-webkit-appearance:none;border:none;background:none;margin:0;padding:6px;font:inherit;color:inherit;',
    '  text-align:inherit;cursor:pointer;box-sizing:border-box;width:100%;border-radius:8px;display:block;min-width:0;}',
    '.kbs-hit:hover{background:var(--surface-2);}',
    '.kbs-hit.open{box-shadow:inset 0 0 0 1px var(--accent);}',
    '.kbs-hit:focus-visible{outline:2px solid var(--accent);outline-offset:-2px;}',
    '.kbs-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px;}',
    '.kbs-lbl{font-size:13px;font-weight:600;color:var(--text);}',
    '.kbs-sub{font-size:11px;color:var(--text-muted);line-height:1.3;}',
    '.kbs-mid{font-size:12px;color:var(--text-secondary);}',
    '.kbs-late,.kbs-late .kbs-lbl,.kbs-late .kbs-big{color:var(--danger);}',
    '.kbs-zero .kbs-big{color:var(--text-muted);}',
    '.kbs-big{font-size:24px;font-weight:600;line-height:1.05;color:var(--text);}',
    '.kbs-bar{display:flex;height:6px;border-radius:3px;overflow:hidden;background:var(--border);}',
    '.kbs-bar>i,.kbs-lbar>i{display:block;height:100%;font-style:normal;}',
    '.kbs-c-doing{background:var(--accent);}.kbs-c-todo{background:var(--text-muted);}',
    '.kbs-c-held{background:var(--kbs-held);}.kbs-c-late{background:var(--danger);}',
    '.kbs-c-done{background:repeating-linear-gradient(135deg,var(--border-strong) 0 3px,transparent 3px 6px);box-shadow:inset 0 0 0 1px var(--border-strong);}',
    '.kbs-list-head{font-size:11px;color:var(--text-muted);margin:6px 4px 2px;}',
    /* リング */
    '.kbs-ring-tile{display:flex;align-items:center;gap:8px;}',
    '.kbs-ring{position:relative;width:52px;height:52px;flex:none;}.kbs-ring svg{display:block;}',
    '.kbs-ring b{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:600;color:var(--text);}',
    /* 数字を大きく */
    '.kbs-banner{display:flex;justify-content:space-between;align-items:center;padding:7px 10px;color:var(--danger);',
    '  background:color-mix(in srgb,var(--danger) 16%,transparent);}',
    '.kbs-banner.ok{color:var(--text-secondary);background:var(--surface-2);}',
    '.kbs-hero-top{display:flex;justify-content:space-between;align-items:baseline;gap:8px;margin-bottom:4px;}',
    /* タイル・色タイル */
    '.kbs-tile{padding:8px 10px;background:var(--surface-2);}',
    '.kbs-tile.kbs-late{background:color-mix(in srgb,var(--danger) 16%,transparent);}',
    '.kbs-heat{padding:9px 10px;color:var(--kbs-c);background:color-mix(in srgb,var(--kbs-c) 15%,transparent);}',
    '.kbs-heat .kbs-lbl,.kbs-heat .kbs-big,.kbs-heat .kbs-sub{color:inherit;}',
    '.kbs-heat .kbs-big{font-size:30px;}',
    '.kbs-heat-bar{height:4px;border-radius:2px;margin-top:6px;background:color-mix(in srgb,currentColor 25%,transparent);overflow:hidden;}',
    '.kbs-heat-bar>i{display:block;height:100%;background:currentColor;}',
    /* ドット */
    '.kbs-dotrow{display:grid;grid-template-columns:48px minmax(0,1fr) auto;gap:6px;align-items:start;}',
    '.kbs-dots{display:flex;flex-wrap:wrap;gap:3px;padding-top:2px;}',
    '.kbs-dot{width:10px;height:10px;border-radius:2px;display:block;}',
    '.kbs-legend{display:flex;flex-wrap:wrap;gap:4px 10px;font-size:11px;color:var(--text-muted);padding:0 6px;}',
    '.kbs-legend i{display:inline-block;width:8px;height:8px;border-radius:2px;margin-right:3px;}',
    /* 今週を主役に */
    '.kbs-gauge{position:relative;max-width:200px;margin:0 auto;}.kbs-gauge svg{display:block;width:100%;height:auto;}',
    '.kbs-gnum{position:absolute;left:0;right:0;bottom:0;text-align:center;}',
    '.kbs-gnum .kbs-big{font-size:32px;}',
    '.kbs-chips{display:flex;flex-wrap:wrap;gap:6px;justify-content:center;}',
    '.kbs-chip{width:auto;display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;box-shadow:inset 0 0 0 1px var(--border-strong);}',
    '.kbs-chip.kbs-late{background:color-mix(in srgb,var(--danger) 16%,transparent);box-shadow:inset 0 0 0 1px color-mix(in srgb,var(--danger) 50%,transparent);}',
    '.kbs-chip.open{box-shadow:inset 0 0 0 1.5px var(--accent);}',
    /* 週ストリップ */
    '.kbs-strip{display:flex;gap:2px;}',
    '.kbs-col{flex:1;min-width:0;padding:4px 1px;text-align:center;}',
    '.kbs-col.today{background:var(--accent-soft);}',
    '.kbs-col.past .kbs-sub{opacity:.7;}',
    '.kbs-stack{display:flex;flex-direction:column-reverse;justify-content:flex-start;gap:2px;height:64px;padding:0 2px;margin-bottom:3px;}',
    '.kbs-blk{display:block;height:10px;border-radius:2px;flex:none;}',
    '.kbs-more{font-size:10px;color:var(--text-muted);line-height:10px;}',
    /* 数字入りの積み上げ */
    '.kbs-lhead{display:flex;justify-content:space-between;gap:8px;font-size:12px;margin-bottom:3px;}',
    '.kbs-lbar{display:flex;height:20px;border-radius:5px;overflow:hidden;background:var(--border);}',
    '.kbs-lbar>i{display:flex;align-items:center;justify-content:center;font-size:11px;white-space:nowrap;overflow:hidden;}',
    '.kbs-t-doing{background:color-mix(in srgb,var(--accent) 22%,transparent);color:var(--accent);}',
    '.kbs-t-todo{background:var(--surface-2);color:var(--text-secondary);box-shadow:inset 0 0 0 1px var(--border-strong);}',
    '.kbs-t-held{background:color-mix(in srgb,var(--kbs-held) 24%,transparent);color:var(--kbs-held-ink);}',
    '.kbs-t-late{background:color-mix(in srgb,var(--danger) 20%,transparent);color:var(--danger);}',
    '.kbs-t-done{color:var(--text-muted);background:repeating-linear-gradient(135deg,var(--border) 0 3px,transparent 3px 6px);box-shadow:inset 0 0 0 1px var(--border-strong);}',
    /* ミニマル */
    '.kbs-mrow{display:grid;grid-template-columns:38px minmax(0,1fr) auto;gap:10px;align-items:center;border-radius:0;}',
    '.kbs-mrow + .kbs-mrow{border-top:1px solid var(--border);}',
    '.kbs-mrow .kbs-big{text-align:right;font-size:26px;}',
    /* 時間軸 */
    '.kbs-tl{position:relative;display:flex;gap:2px;}',
    '.kbs-tl::before{content:"";position:absolute;left:12.5%;right:12.5%;top:26px;height:2px;background:var(--border-strong);}',
    '.kbs-node{flex:1;text-align:center;position:relative;}',
    '.kbs-circ{width:40px;height:40px;border-radius:50%;margin:0 auto 4px;display:flex;align-items:center;justify-content:center;',
    '  font-size:16px;font-weight:600;color:var(--text);background:var(--surface);box-shadow:inset 0 0 0 1px var(--border-strong);}',
    '.kbs-node.kbs-late .kbs-circ{color:var(--danger);background:color-mix(in srgb,var(--danger) 16%,var(--surface));box-shadow:inset 0 0 0 1.5px var(--danger);}',
    '.kbs-node.now .kbs-circ{box-shadow:inset 0 0 0 2px var(--accent);}',
    '.kbs-node.kbs-zero .kbs-circ{color:var(--text-muted);}',
    /* 切り替え */
    '.kb-foot{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:6px 8px;padding-top:8px;}',
    '.kb-foot .kb-forget{white-space:nowrap;margin-left:auto;}',
    '.kb-style-pick{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--text-muted);min-width:0;white-space:nowrap;}',
    '.kb-style-pick select{font:inherit;font-size:11px;color:var(--text-secondary);background:var(--surface-2);border:1px solid var(--border);',
    '  border-radius:6px;padding:2px 4px;max-width:150px;min-width:0;}',
    '@media (prefers-reduced-motion:reduce){.kbs *{transition:none!important;}}'
  ].join('\n');

  function injectCss() {
    if (typeof document === 'undefined' || document.getElementById('kbs-style')) return;
    var s = document.createElement('style');
    s.id = 'kbs-style';
    s.textContent = CSS;
    document.head.appendChild(s);
  }

  /* ---------- 小物 ---------- */
  function normalize(key) {
    for (var i = 0; i < STYLES.length; i++) if (STYLES[i].key === key) return key;
    return 'bars';
  }
  function md(d) { return d ? (d.getMonth() + 1) + '/' + d.getDate() : ''; }
  function count(r, s) { return r.rest.filter(function (t) { return t.status === s; }).length; }
  function pct(r) { return r.total ? Math.round(r.done.length / r.total * 100) : 0; }
  function addDays(d, n) { var t = new Date(d); t.setDate(t.getDate() + n); return t; }
  function dayKey(d) { return 'd' + d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2); }
  function sameDay(a, b) { return !!a && !!b && a.getTime() === b.getTime(); }

  /* 表示する行（定義にあって、rows にも入っているものだけ） */
  function defsOf(ctx) {
    return (ctx.defs || []).filter(function (d) { return ctx.rows && ctx.rows[d.key]; });
  }
  function defOf(ctx, key) {
    var list = defsOf(ctx);
    for (var i = 0; i < list.length; i++) if (list[i].key === key) return list[i];
    return null;
  }
  function isTotalRow(def) { return !!(def && def.total); }

  /* 行の下に添える説明（期間・期限切れなど） */
  function subOf(def, r) {
    if (def.key === 'late') return '期限切れ';
    if (def.key === 'todo') return '日付未設定';
    if (def.key === 'today') return md(r.from) + ' 期限';
    return r.from ? md(r.from) + '～' + md(r.to) : '';
  }
  function progressOf(def, r) {
    if (!isTotalRow(def)) return '';
    if (r.total && !r.rest.length) return 'すべて完了';
    return r.total ? '消化 ' + pct(r) + '%' : '';
  }
  function numText(def, r) {
    return isTotalRow(def) ? '残 ' + r.rest.length + ' / ' + r.total : r.rest.length + ' 件';
  }
  function tipOf(def, r) {
    if (!isTotalRow(def)) return def.label + ' ' + r.rest.length + '件';
    return '残り ' + r.rest.length + '（対応中' + count(r, 'doing') + ' / 未着手' + count(r, 'todo') +
      ' / 保留' + count(r, 'held') + '） 完了 ' + r.done.length;
  }
  function toneClass(def, r) {
    if (def.key === 'late' && r.rest.length) return ' kbs-late';
    if (!r.rest.length) return ' kbs-zero';
    return '';
  }

  /* 押すと一覧が開く部品。ここで作るものはすべてこれを通す */
  function hit(ctx, key, cls, inner, label, style) {
    var open = ctx.expanded && ctx.expanded.has(key);
    return '<button type="button" class="kbs-hit ' + (cls || '') + (open ? ' open' : '') + '"' +
      (style ? ' style="' + style + '"' : '') +
      ' aria-expanded="' + (open ? 'true' : 'false') + '"' +
      (label ? ' aria-label="' + ctx.esc(label) + '" title="' + ctx.esc(label) + '"' : '') +
      ' data-kb-row="' + ctx.esc(ctx.rowAttr(key)) + '">' + inner + '</button>';
  }

  /* 開いている行のタスク一覧を、表示順に並べる */
  function listsHtml(ctx, extra) {
    var order = defsOf(ctx).map(function (d) { return { key: d.key, label: d.label, r: ctx.rows[d.key] }; });
    if (extra) order = order.concat(extra);
    var html = '';
    order.forEach(function (o) {
      if (!ctx.expanded || !ctx.expanded.has(o.key)) return;
      html += '<div class="kbs-list-head">' + ctx.esc(o.label) + 'のタスク</div>' + ctx.listHtml(o.key, o.label, o.r);
    });
    return html;
  }

  /* 状態別の積み上げ（バー系で共通）。scale はグループ内の最大件数 */
  function segs(def, r, scale, prefix) {
    var w = function (x) { return (scale ? x / scale * 100 : 0) + '%'; };
    if (!isTotalRow(def)) {
      var c = def.key === 'late' ? 'late' : 'todo';
      return '<i class="' + prefix + c + '" style="width:' + w(r.rest.length) + '"></i>';
    }
    return STATUS_ORDER.map(function (s) {
      return '<i class="' + prefix + s + '" style="width:' + w(count(r, s)) + '"></i>';
    }).join('') + '<i class="' + prefix + 'done" style="width:' + w(r.done.length) + '"></i>';
  }
  function groupScale(ctx, def) {
    var list = defsOf(ctx).filter(function (d) { return d.group === def.group; });
    var max = 1;
    list.forEach(function (d) {
      var r = ctx.rows[d.key];
      max = Math.max(max, isTotalRow(d) ? r.total : r.rest.length);
    });
    return max;
  }

  /* ---------- 1. リング ---------- */
  function ring(ctx) {
    var C = 2 * Math.PI * 20;
    var tiles = defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var color = def.key === 'late' ? 'var(--danger)' : def.key === 'todo' ? 'var(--text-muted)' : 'var(--accent)';
      var frac = isTotalRow(def) ? (r.total ? r.done.length / r.total : 0) : (r.rest.length ? 1 : 0);
      var arc = frac > 0
        ? '<circle cx="26" cy="26" r="20" fill="none" stroke="' + color + '" stroke-width="5" stroke-dasharray="' +
          (C * frac).toFixed(1) + ' ' + C.toFixed(1) + '" transform="rotate(-90 26 26)"/>' : '';
      var inner =
        '<div class="kbs-ring-tile">' +
          '<div class="kbs-ring"><svg width="52" height="52" viewBox="0 0 52 52" aria-hidden="true">' +
            '<circle cx="26" cy="26" r="20" fill="none" stroke="var(--border)" stroke-width="5"/>' + arc +
          '</svg><b' + (def.key === 'late' && r.rest.length ? ' style="color:var(--danger)"' : '') + '>' + r.rest.length + '</b></div>' +
          '<div style="min-width:0"><div class="kbs-lbl">' + ctx.esc(def.label) + '</div>' +
          '<div class="kbs-mid">' + (isTotalRow(def) ? '残 ' + r.rest.length + ' / ' + r.total : ctx.esc(subOf(def, r))) + '</div>' +
          '<div class="kbs-sub">' + ctx.esc(isTotalRow(def) ? progressOf(def, r) : '') + '</div></div>' +
        '</div>';
      return hit(ctx, def.key, toneClass(def, r), inner, tipOf(def, r));
    });
    return '<div class="kbs-grid">' + tiles.join('') + '</div>';
  }

  /* ---------- 2. 数字を大きく ---------- */
  function hero(ctx) {
    var html = '';
    var late = defOf(ctx, 'late');
    if (late) {
      var lr = ctx.rows.late;
      html += hit(ctx, 'late', 'kbs-banner' + (lr.rest.length ? '' : ' ok'),
        '<span>遅延　期限切れ</span><span style="font-size:18px;font-weight:600">' + lr.rest.length + '件</span>',
        tipOf(late, lr));
    }
    defsOf(ctx).forEach(function (def) {
      if (def.key === 'late') return;
      var r = ctx.rows[def.key];
      if (!isTotalRow(def)) {
        html += hit(ctx, def.key, toneClass(def, r),
          '<div class="kbs-hero-top"><span class="kbs-mid">' + ctx.esc(def.label) + '　' + ctx.esc(subOf(def, r)) +
          '</span><span class="kbs-mid">' + r.rest.length + ' 件</span></div>', tipOf(def, r));
        return;
      }
      html += hit(ctx, def.key, toneClass(def, r),
        '<div class="kbs-hero-top"><span class="kbs-mid">' + ctx.esc(def.label) + '</span>' +
        '<span><span class="kbs-big">' + r.rest.length + '</span><span class="kbs-mid"> / ' + r.total + '　' +
        ctx.esc(progressOf(def, r)) + '</span></span></div>' +
        '<div class="kbs-bar">' + segs(def, r, groupScale(ctx, def), 'kbs-c-') + '</div>', tipOf(def, r));
    });
    return html;
  }

  /* ---------- 3. タイル ---------- */
  function tiles(ctx) {
    var html = defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var bar = isTotalRow(def)
        ? '<div class="kbs-bar" style="height:4px;margin-top:5px"><i class="kbs-c-doing" style="width:' + pct(r) + '%"></i></div>'
        : '<div class="kbs-sub" style="margin-top:3px">' + ctx.esc(subOf(def, r)) + '</div>';
      return hit(ctx, def.key, 'kbs-tile' + toneClass(def, r),
        '<div class="kbs-mid"' + (def.key === 'late' && r.rest.length ? ' style="color:inherit"' : '') + '>' + ctx.esc(def.label) + '</div>' +
        '<div><span class="kbs-big">' + r.rest.length + '</span>' +
        (isTotalRow(def) ? '<span class="kbs-mid"> / ' + r.total + '</span>' : '') + '</div>' + bar,
        tipOf(def, r));
    });
    return '<div class="kbs-grid">' + html.join('') + '</div>';
  }

  /* ---------- 4. ドットで1件ずつ ---------- */
  function dots(ctx) {
    var MAX = 30;
    var rowsHtml = defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var cls = [];
      if (!isTotalRow(def)) {
        r.rest.forEach(function () { cls.push(def.key === 'late' ? 'late' : 'todo'); });
      } else {
        STATUS_ORDER.forEach(function (s) { for (var i = 0; i < count(r, s); i++) cls.push(s); });
        r.done.forEach(function () { cls.push('done'); });
      }
      var shown = cls.slice(0, MAX).map(function (c) { return '<i class="kbs-dot kbs-c-' + c + '"></i>'; }).join('');
      if (cls.length > MAX) shown += '<span class="kbs-sub">+' + (cls.length - MAX) + '</span>';
      if (!cls.length) shown = '<span class="kbs-sub">なし</span>';
      return hit(ctx, def.key, 'kbs-dotrow' + toneClass(def, r),
        '<span class="kbs-lbl" style="font-size:12px">' + ctx.esc(def.label) + '</span>' +
        '<span class="kbs-dots">' + shown + '</span>' +
        '<span class="kbs-mid" style="white-space:nowrap">' + ctx.esc(isTotalRow(def) ? '残' + r.rest.length + '/' + r.total : r.rest.length + '件') + '</span>',
        tipOf(def, r));
    }).join('');
    var legend = '<div class="kbs-legend">' +
      ['doing', 'todo', 'held', 'done'].map(function (s) {
        return '<span><i class="kbs-c-' + s + '"></i>' + STATUS_SHORT[s] + '</span>';
      }).join('') + '</div>';
    return rowsHtml + legend;
  }

  /* ---------- 5. 今週を主役に ---------- */
  function focus(ctx) {
    var html = '';
    var wdef = defOf(ctx, 'week');
    if (wdef) {
      var r = ctx.rows.week;
      var L = Math.PI * 80;
      var frac = r.total ? r.done.length / r.total : 0;
      var arc = frac > 0
        ? '<path d="M20 98 A80 80 0 0 1 180 98" fill="none" stroke="var(--accent)" stroke-width="13" stroke-linecap="round" stroke-dasharray="' +
          (L * frac).toFixed(1) + ' ' + L.toFixed(1) + '"/>' : '';
      html += hit(ctx, 'week', toneClass(wdef, r),
        '<div class="kbs-gauge"><svg viewBox="0 0 200 108" aria-hidden="true">' +
          '<path d="M20 98 A80 80 0 0 1 180 98" fill="none" stroke="var(--border)" stroke-width="13" stroke-linecap="round"/>' + arc +
        '</svg><div class="kbs-gnum"><div class="kbs-big">' + r.rest.length + '</div>' +
        '<div class="kbs-mid">今週 残り / ' + r.total + '件　' + ctx.esc(progressOf(wdef, r)) + '</div></div></div>',
        tipOf(wdef, r));
    }
    var chips = defsOf(ctx).filter(function (d) { return d.key !== 'week'; }).map(function (def) {
      var r = ctx.rows[def.key];
      return hit(ctx, def.key, 'kbs-chip' + (def.key === 'late' && r.rest.length ? ' kbs-late' : ''),
        ctx.esc(def.label) + ' ' + (isTotalRow(def) ? '残' + r.rest.length + '/' + r.total : r.rest.length + '件'),
        tipOf(def, r));
    }).join('');
    return html + '<div class="kbs-chips">' + chips + '</div>';
  }

  /* ---------- 6. 期限日の週ストリップ ---------- */
  function strip(ctx) {
    var MAXB = 5;
    var week = ctx.rows.week;
    var mon = week && week.from ? week.from : null;
    var cols = [];
    var extra = [];

    function blocks(items) {
      var cls = [];
      items.forEach(function (t) { cls.push(t.status === 'done' ? 'done' : t.status); });
      var html = cls.slice(0, MAXB).map(function (c) { return '<i class="kbs-blk kbs-c-' + c + '"></i>'; }).join('');
      if (cls.length > MAXB) html += '<span class="kbs-more">+' + (cls.length - MAXB) + '</span>';
      return html;
    }
    function col(key, label, sub, items, cls, tip) {
      return hit(ctx, key, 'kbs-col ' + (cls || ''),
        '<div class="kbs-stack">' + blocks(items) + '</div>' +
        '<div class="kbs-sub"' + (cls && cls.indexOf('kbs-late') >= 0 ? ' style="color:var(--danger)"' : '') + '>' +
        ctx.esc(label) + '<br>' + ctx.esc(sub) + '</div>', tip);
    }

    var late = ctx.rows.late;
    if (late) cols.push(col('late', '遅延', String(late.rest.length), late.rest.map(function () { return { status: 'late' }; }),
      late.rest.length ? 'kbs-late' : '', '遅延 ' + late.rest.length + '件'));

    if (mon) {
      var weekItems = week.rest.concat(week.done);
      for (var i = 0; i < 7; i++) {
        var d = addDays(mon, i);
        var due = weekItems.filter(function (t) { return sameDay(t.end, d); });
        var rest = due.filter(function (t) { return t.status !== 'done'; });
        var done = due.filter(function (t) { return t.status === 'done'; });
        var items = rest.slice().sort(function (a, b) { return STATUS_ORDER.indexOf(a.status) - STATUS_ORDER.indexOf(b.status); }).concat(done);
        var key = dayKey(d);
        var isToday = sameDay(d, ctx.today);
        var past = d < ctx.today;
        var label = WEEKDAYS[d.getDay()];
        cols.push(col(key, label, String(d.getDate()), items, (isToday ? 'today' : '') + (past ? ' past' : ''),
          md(d) + (isToday ? '（今日）' : '') + ' 期限 ' + due.length + '件（残り ' + rest.length + '）'));
        extra.push({ key: key, label: md(d) + '（' + WEEKDAYS[d.getDay()] + '）期限', r: { rest: rest, done: done } });
      }
    }

    var next = ctx.rows.next;
    if (next) cols.push(col('next', '来週', String(next.rest.length),
      next.rest.slice().sort(function (a, b) { return STATUS_ORDER.indexOf(a.status) - STATUS_ORDER.indexOf(b.status); }).concat(next.done),
      '', '来週 残り ' + next.rest.length + ' / ' + next.total));

    var todo = ctx.rows.todo;
    var tdef = defOf(ctx, 'todo');
    var chip = tdef ? '<div class="kbs-chips" style="justify-content:flex-start">' +
      hit(ctx, 'todo', 'kbs-chip', ctx.esc(tdef.label) + ' ' + todo.rest.length + '件（日付未設定）', tipOf(tdef, todo)) + '</div>' : '';

    return { html: '<div class="kbs-sub" style="padding:0 4px">期限日ごとの件数（今週）</div>' +
      '<div class="kbs-strip">' + cols.join('') + '</div>' + chip, extra: extra };
  }

  /* ---------- 7. 数字入りの積み上げ ---------- */
  function labeled(ctx) {
    return defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var scale = groupScale(ctx, def);
      var w = function (x) { return (scale ? x / scale * 100 : 0) + '%'; };
      var seg = function (cls, n, word) {
        if (!n) return '';
        var units = scale ? n / scale : 0;
        var text = units >= 0.3 ? word + n : String(n);
        return '<i class="kbs-t-' + cls + '" style="width:' + w(n) + '">' + text + '</i>';
      };
      var bar = !isTotalRow(def)
        ? seg(def.key === 'late' ? 'late' : 'todo', r.rest.length, def.key === 'late' ? '遅延' : '')
        : seg('doing', count(r, 'doing'), '対応中') + seg('todo', count(r, 'todo'), '未着手') +
          seg('held', count(r, 'held'), '保留') + seg('done', r.done.length, '完了');
      return hit(ctx, def.key, toneClass(def, r),
        '<div class="kbs-lhead"><span class="kbs-lbl">' + ctx.esc(def.label) + '</span>' +
        '<span class="kbs-mid">' + ctx.esc(numText(def, r)) + '</span></div>' +
        '<div class="kbs-lbar">' + bar + '</div>', tipOf(def, r));
    }).join('');
  }

  /* ---------- 8. 罫線だけのミニマル ---------- */
  function minimal(ctx) {
    return defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var right = isTotalRow(def)
        ? '<div class="kbs-mid">/ ' + r.total + '</div><div class="kbs-sub">' + ctx.esc(progressOf(def, r)) + '</div>'
        : '<div class="kbs-mid">' + (def.key === 'late' && r.rest.length ? '要対応' : '件') + '</div>';
      return hit(ctx, def.key, 'kbs-mrow' + toneClass(def, r),
        '<span class="kbs-big">' + r.rest.length + '</span>' +
        '<span style="min-width:0"><span class="kbs-lbl" style="display:block">' + ctx.esc(def.label) + '</span>' +
        '<span class="kbs-sub">' + ctx.esc(subOf(def, r)) + '</span></span>' +
        '<span style="text-align:right">' + right + '</span>', tipOf(def, r));
    }).join('');
  }

  /* ---------- 9. 緊急度の色タイル ---------- */
  var HEAT = { late: 'var(--danger)', today: 'var(--kbs-held-ink)', week: 'var(--accent)', next: 'var(--text-secondary)', todo: 'var(--text-muted)' };
  function heat(ctx) {
    var html = defsOf(ctx).map(function (def) {
      var r = ctx.rows[def.key];
      var color = HEAT[def.key] || 'var(--text-secondary)';
      var foot = isTotalRow(def)
        ? '<div class="kbs-sub">/ ' + r.total + '　' + (r.total && !r.rest.length ? '完了' : pct(r) + '%') + '</div>' +
          '<div class="kbs-heat-bar"><i style="width:' + pct(r) + '%"></i></div>'
        : '<div class="kbs-sub">' + ctx.esc(subOf(def, r)) + '</div>';
      return hit(ctx, def.key, 'kbs-heat',
        '<div class="kbs-lbl" style="font-weight:500">' + ctx.esc(def.label) + '</div>' +
        '<div class="kbs-big">' + r.rest.length + '</div>' + foot, tipOf(def, r), '--kbs-c:' + color);
    });
    return '<div class="kbs-grid">' + html.join('') + '</div>';
  }

  /* ---------- 10. 時間軸 ---------- */
  function timeline(ctx) {
    var onLine = ['late', 'today', 'week', 'next'];
    var nodes = onLine.map(function (key) {
      var def = defOf(ctx, key);
      if (!def) return '';
      var r = ctx.rows[key];
      var cls = 'kbs-node' + toneClass(def, r) + (key === 'today' ? ' now' : '');
      return hit(ctx, key, cls,
        '<div class="kbs-circ">' + r.rest.length + '</div>' +
        '<div class="kbs-lbl" style="font-size:12px">' + ctx.esc(def.label) + '</div>' +
        '<div class="kbs-sub">' + ctx.esc(isTotalRow(def) ? '残' + r.rest.length + '/' + r.total : subOf(def, r)) + '</div>',
        tipOf(def, r));
    }).join('');
    var tdef = defOf(ctx, 'todo');
    var chip = tdef ? '<div class="kbs-chips" style="justify-content:flex-start">' +
      hit(ctx, 'todo', 'kbs-chip', ctx.esc(tdef.label) + ' ' + ctx.rows.todo.rest.length + '件（日付未設定）', tipOf(tdef, ctx.rows.todo)) +
      '</div>' : '';
    return '<div class="kbs-tl">' + nodes + '</div>' + chip;
  }

  var RENDERERS = { ring: ring, hero: hero, tiles: tiles, dots: dots, focus: focus, strip: strip,
    labeled: labeled, minimal: minimal, heat: heat, timeline: timeline };

  /** スタイルの描画。'bars' と知らない値は null（呼び出し側の標準の描画を使う） */
  function render(style, ctx) {
    var fn = RENDERERS[normalize(style)];
    if (!fn) return null;
    injectCss();
    var out = fn(ctx);
    var html = typeof out === 'string' ? out : out.html;
    var extra = typeof out === 'string' ? null : out.extra;
    return '<div class="kbs kbs-s-' + style + '">' + html + listsHtml(ctx, extra) + '</div>';
  }

  /** 切り替え用の <select>。attr はそのまま要素に付ける（data-kb-style="id" など） */
  function pickerHtml(current, attr, esc) {
    injectCss();
    var cur = normalize(current);
    var e = esc || function (s) { return String(s); };
    return '<label class="kb-style-pick">表示<select ' + (attr || '') + ' aria-label="表示スタイル">' +
      STYLES.map(function (s) {
        return '<option value="' + s.key + '"' + (s.key === cur ? ' selected' : '') + '>' + e(s.label) + '</option>';
      }).join('') + '</select></label>';
  }

  window.KanbanStyles = {
    VERSION: VERSION,
    STYLES: STYLES,
    normalize: normalize,
    render: render,
    pickerHtml: pickerHtml,
    dayKey: dayKey,
    _css: injectCss
  };
})();
