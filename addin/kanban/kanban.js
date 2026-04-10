<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Kanban</title>

  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script src="kanban.js"></script>
  <link rel="stylesheet" href="style.css">

  <script>
    const SHOW_VERSION = true;
    const VIEW_VERSION = "rev_20260410_xxxxxx";

    function renderVersion() {
      if (!SHOW_VERSION) return;

      const el = document.getElementById("version-label");
      el.textContent =
        `view(${VIEW_VERSION}) js(${window.APP_VERSION || "unknown"})`;
    }
  </script>
</head>

<body onload="renderVersion()">

<h2>カンバン</h2>

<div style="position:absolute; top:10px; right:10px; display:flex; gap:10px;">
  <span id="version-label" style="font-size:10px; color:#666;"></span>
  <button onclick="init()">Reload</button>
</div>

<h4>担当者</h4>
<div id="user-filters"></div>

<h4>分類</h4>
<div id="category-filters"></div>

<h4>期間</h4>
<div>
  <button onclick="setPeriod('all')">全期間</button>
  <button onclick="setPeriod('past')">以前</button>
  <button onclick="setPeriod('week')">今週</button>
  <button onclick="setPeriod('nextweek')">来週</button>
  <button onclick="setPeriod('future')">以降</button>
</div>

<hr>

<div id="board">

  <div class="lane" id="todo"
       ondragover="allowDrop(event)"
       ondrop="drop(event,'todo')">
    <h3>未着手</h3>
    <div class="card-list"></div>
  </div>

  <div class="lane" id="doing"
       ondragover="allowDrop(event)"
       ondrop="drop(event,'doing')">
    <h3>対応中</h3>
    <div class="card-list"></div>
  </div>

  <div class="lane" id="done"
       ondragover="allowDrop(event)"
       ondrop="drop(event,'done')">
    <h3>完了</h3>
    <div class="card-list"></div>
  </div>

</div>

<!-- モーダル -->
<div id="modal" class="modal hidden">
  <div class="modal-content">
    <h3 id="modal-title"></h3>
    <textarea id="modal-note" rows="6"></textarea>
    <br>
    <button onclick="saveNote()">保存</button>
    <button onclick="closeModal()">閉じる</button>
  </div>
</div>

</body>
</html>