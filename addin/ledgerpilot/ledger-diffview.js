/* LedgerPilot 差分ビュー（取込ツール／アドイン共用）
 * LedgerDiffView.render(container, diff, { onChange(selectedYms) }) → { getSelected() }
 */
(function (root) {
  "use strict";
  var TAG = { "追加": "add", "変更": "chg", "削除": "del", "同一": "same" };

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function yen(n) { return Number(n || 0).toLocaleString("ja-JP"); }
  function delta(a, b) {
    var d = b - a;
    if (!d) return '<span class="muted">±0</span>';
    return '<span class="' + (d > 0 ? "pos" : "neg") + '">' + (d > 0 ? "+" : "") + yen(d) + "</span>";
  }

  function voucherRow(v) {
    var s = v.summary, extra = "";
    if (v.status === "変更") {
      var move = v.oldYm && v.oldYm !== v.ym ? ' <span class="lp-tag move">' + esc(v.oldYm) + " から移動</span>" : "";
      var items = (v.diffs || []).slice(0, 6).map(function (d) {
        if (d.kind !== "変更") return "行" + esc(d.gyo) + " " + d.kind;
        return "行" + esc(d.gyo) + " <b>" + esc(d.col) + "</b> <del>" + esc(d.before || "（空）") + "</del> → <ins>" + esc(d.after || "（空）") + "</ins>";
      });
      if ((v.diffs || []).length > 6) items.push("ほか " + (v.diffs.length - 6) + " 件");
      extra = '<div class="fd">' + move + items.join("<br>") + "</div>";
    }
    return "<tr><td>" + '<span class="lp-tag ' + TAG[v.status] + '">' + v.status + "</span></td>" +
      "<td class=\"num\">" + esc(s.date) + "</td>" +
      "<td>" + esc(s.memo || "（摘要なし）") + extra + "</td>" +
      '<td class="num">' + yen(s.debit) + (v.oldSummary && v.oldSummary.debit !== s.debit ? '<div class="fd">旧 ' + yen(v.oldSummary.debit) + "</div>" : "") + "</td></tr>";
  }

  function render(container, diff, opts) {
    opts = opts || {};
    var months = Object.keys(diff.months).sort().map(function (k) { return diff.months[k]; });
    var selected = {};
    months.forEach(function (m) { if (m.changed) selected[m.ym] = true; });
    var open = {}, showSame = false;

    function draw() {
      var anyChange = months.some(function (m) { return m.changed; });
      var html = "";
      if (diff.outOfPeriodCsv) {
        html += '<div class="lp-note warn">CSVのうち ' + diff.outOfPeriodCsv + " 行は指定期間の外にあるため取り込みません。</div>";
      }
      if (!anyChange) {
        html += '<div class="lp-note ok">指定期間の台帳はCSVと一致しています。差し替えは不要です。</div>';
      }
      html += '<div class="lp-scroll"><table class="lp-table lp-diff-months"><thead><tr>' +
        '<th class="chk"><input type="checkbox" data-all' + (anyChange ? "" : " disabled") + "></th>" +
        "<th>年月</th><th class=\"num\">追加</th><th class=\"num\">変更</th><th class=\"num\">削除</th><th class=\"num\">同一</th>" +
        "<th class=\"num\">台帳の行</th><th class=\"num\">CSVの行</th><th class=\"num\">借方合計（台帳 → CSV）</th></tr></thead><tbody>";
      months.forEach(function (m) {
        var link = m.linked.length ? ' <span class="lp-tag move" title="月をまたいで日付が変わった伝票があるため一緒に差し替えます">' + m.linked.join("・") + " と連動</span>" : "";
        html += '<tr class="clickable ' + (m.changed ? "" : "nochange") + '" data-ym="' + m.ym + '">' +
          '<td class="chk"><input type="checkbox" data-sel="' + m.ym + '"' + (selected[m.ym] ? " checked" : "") + (m.changed ? "" : " disabled") + "></td>" +
          "<td><b>" + m.ym + "</b>" + link + (m.moveOut ? ' <span class="lp-tag move">移動 ' + m.moveOut + "</span>" : "") + "</td>" +
          '<td class="num">' + (m.add ? '<span class="lp-tag add">' + m.add + "</span>" : "0") + "</td>" +
          '<td class="num">' + (m.chg ? '<span class="lp-tag chg">' + m.chg + "</span>" : "0") + "</td>" +
          '<td class="num">' + (m.del ? '<span class="lp-tag del">' + m.del + "</span>" : "0") + "</td>" +
          '<td class="num">' + m.same + "</td>" +
          '<td class="num">' + m.existingRows + "</td>" +
          '<td class="num">' + m.incomingRows + "</td>" +
          '<td class="num">' + yen(m.oldDebit) + " → " + yen(m.newDebit) + " " + delta(m.oldDebit, m.newDebit) + "</td></tr>";
        if (open[m.ym]) {
          var vs = diff.vouchers.filter(function (v) {
            return (v.ym === m.ym || (v.oldYm === m.ym && v.ym !== m.ym)) && (showSame || v.status !== "同一");
          });
          html += '<tr class="lp-diff-detail"><td></td><td colspan="8">' +
            '<label class="lp-toggle"><input type="checkbox" data-same' + (showSame ? " checked" : "") + "> 同一の伝票も表示</label>" +
            (vs.length ? '<table class="lp-vlist">' + vs.map(voucherRow).join("") + "</table>" : '<div class="muted">表示する伝票はありません。</div>') +
            "</td></tr>";
        }
      });
      html += "</tbody></table></div>";
      container.innerHTML = html;

      container.querySelectorAll("[data-sel]").forEach(function (cb) {
        cb.addEventListener("click", function (e) { e.stopPropagation(); });
        cb.addEventListener("change", function () {
          var ym = cb.getAttribute("data-sel");
          setSel(ym, cb.checked);
          draw(); fire();
        });
      });
      var all = container.querySelector("[data-all]");
      if (all) {
        all.checked = months.filter(function (m) { return m.changed; }).every(function (m) { return selected[m.ym]; });
        all.addEventListener("change", function () {
          months.forEach(function (m) { if (m.changed) selected[m.ym] = all.checked; });
          draw(); fire();
        });
      }
      container.querySelectorAll("tr[data-ym]").forEach(function (tr) {
        tr.addEventListener("click", function () {
          var ym = tr.getAttribute("data-ym");
          open[ym] = !open[ym]; draw();
        });
      });
      var same = container.querySelector("[data-same]");
      if (same) same.addEventListener("change", function () { showSame = same.checked; draw(); });
    }
    // 連動月は一緒にON/OFF（月移動した伝票の2重計上・消失を防ぐ）
    function setSel(ym, on, seen) {
      seen = seen || {};
      if (seen[ym]) return; seen[ym] = 1;
      selected[ym] = on;
      var m = diff.months[ym];
      if (m) m.linked.forEach(function (l) { setSel(l, on, seen); });
    }
    function getSelected() {
      return Object.keys(selected).filter(function (k) { return selected[k]; }).sort();
    }
    function fire() { if (opts.onChange) opts.onChange(getSelected()); }
    draw(); fire();
    return { getSelected: getSelected };
  }

  root.LedgerDiffView = { render: render, esc: esc, yen: yen };
})(window);
