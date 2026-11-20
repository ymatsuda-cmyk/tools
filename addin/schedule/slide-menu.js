/* ============================================================
 * slide-menu.js — 共通スライドメニューコンポーネント（JSON駆動）
 * ------------------------------------------------------------
 * メニューの「名前・URL」は menu.json で一元管理する。
 * menu.json を編集すれば、組み込んでいる全アプリのメニューが
 * 一斉に更新される。CSS・DOMはこのJSが自動注入するため、
 * ホスト側にHTML/CSSの追加は不要。
 *
 * ■ 配置
 *   tools/addin/common/slide-menu.js   （このファイル）
 *   tools/addin/common/menu.json       （メニュー定義）
 *
 * ■ ホスト側の組み込み（遅延ロードの例）
 *
 *   const COMMON = 'https://ymatsuda-cmyk.github.io/tools/addin/common';
 *   let menuReady = null;
 *   function openMenu() {
 *     if (!menuReady) {
 *       menuReady = new Promise((resolve, reject) => {
 *         const s = document.createElement('script');
 *         s.src = COMMON + '/slide-menu.js';
 *         s.onload = () => {
 *           SlideMenu.init({
 *             appName:  'Excel Kanban',
 *             version:  APP_VERSION,
 *             currentId:'kanban',                 // menu.json 内のidと一致で強調
 *             menuUrl:  COMMON + '/menu.json',
 *             localItems: [                        // アプリ固有の操作
 *               { section: '操作' },
 *               { label: '再読み込み', icon: '🔄', onClick: () => init() }
 *             ]
 *           });
 *           resolve();
 *         };
 *         s.onerror = reject;
 *         document.head.appendChild(s);
 *       });
 *     }
 *     menuReady.then(() => SlideMenu.open());
 *   }
 *
 * ■ menu.json の形式
 *   {
 *     "sections": [
 *       { "title": "ツール",
 *         "items": [
 *           { "id": "kanban", "label": "カンバン", "icon": "📋",
 *             "url": "https://.../kanban/", "external": false }
 *         ] }
 *     ]
 *   }
 *
 * ■ ポップアップ項目（Excel連携ツール用）
 *   menu.json の項目に "popup": "https://.../app.html" を書くと、画面遷移せずに、
 *   Officeのポップアップ（ダイアログ）でそのページを開く。同じ common フォルダの
 *   schedule-bridge.js を自動で読み込み、Excelとの読み書きを仲介する。
 *   例: { "id": "schedule", "label": "スケジュール",
 *         "popup": "https://.../schedule/app.html", "width": 95, "height": 92 }
 *   ・ポップアップは、メニューを開いているページ（作業ウィンドウ側）が仲介する。
 *     そのページを閉じる・別ツールへ遷移すると、ポップアップも閉じる。
 *   ・ポップアップのURLは、メニューを開いているページと同じドメインにすること。
 *
 * ■ 公開API
 *   SlideMenu.init(config)   初期化（再実行で設定を更新）
 *   SlideMenu.open()         開く（初回はmenu.jsonを取得）
 *   SlideMenu.close()
 *   SlideMenu.toggle()
 *   SlideMenu.isOpen()
 *   SlideMenu.reload()       menu.jsonを再取得して描画し直す
 *   SlideMenu.openPopup(item) ポップアップ項目を開く（menu.json の1項目と同じ形）
 * ============================================================ */
(function (global) {
  "use strict";

  // 二重読み込みガード
  if (global.SlideMenu) return;

  var config = null;
  var built = false;
  var opened = false;
  var lastFocus = null;
  var menuData = null;     // menu.json のキャッシュ
  var fetching = null;     // 取得中のPromise
  var SELF_SRC = (document.currentScript && document.currentScript.src) || "";   // このファイルのURL（共通フォルダの特定に使う）
  var popupHandles = {};   // ポップアップ項目ごとの、開いているダイアログ
  var bridgeLoading = null;

  /* ---------------- CSS ---------------- */
  var CSS = "" +
":root{--sm-accent:#0E7A5F;--sm-bg:#FFFFFF;--sm-ink:#1C2733;--sm-ink-soft:#5B6B7B;" +
"--sm-ink-faint:#93A1AF;--sm-line:#E2E7EC;--sm-hover:#F3F5F7;--sm-width:250px;" +
"--sm-z:9000;--sm-font:'Segoe UI Variable','Segoe UI','Yu Gothic UI','Hiragino Sans','Meiryo',sans-serif}" +
".sm-overlay{position:fixed;inset:0;background:rgba(28,39,51,.35);opacity:0;" +
"visibility:hidden;transition:opacity .2s ease,visibility .2s;z-index:var(--sm-z)}" +
".sm-overlay.sm-show{opacity:1;visibility:visible}" +
".sm-panel{position:fixed;top:0;bottom:0;width:var(--sm-width);max-width:85vw;" +
"background:var(--sm-bg);box-shadow:0 8px 32px rgba(28,39,51,.24);" +
"z-index:calc(var(--sm-z) + 1);display:flex;flex-direction:column;" +
"font-family:var(--sm-font);font-size:13px;color:var(--sm-ink);" +
"transition:transform .22s cubic-bezier(.2,.8,.3,1)}" +
".sm-panel.sm-right{right:0;transform:translateX(100%)}" +
".sm-panel.sm-left{left:0;transform:translateX(-100%)}" +
".sm-panel.sm-show{transform:translateX(0)}" +
".sm-head{display:flex;align-items:center;gap:10px;padding:14px 14px 12px;" +
"border-bottom:1px solid var(--sm-line)}" +
".sm-app-mark{width:30px;height:30px;border-radius:8px;background:var(--sm-accent);" +
"color:#fff;display:flex;align-items:center;justify-content:center;" +
"font-weight:700;font-size:14px;flex-shrink:0}" +
".sm-app-name{font-weight:700;font-size:13px;line-height:1.2}" +
".sm-app-ver{font-size:10px;color:var(--sm-ink-faint);font-variant-numeric:tabular-nums}" +
".sm-close{margin-left:auto;border:none;background:transparent;color:var(--sm-ink-faint);" +
"width:26px;height:26px;border-radius:6px;cursor:pointer;font-size:14px;" +
"display:flex;align-items:center;justify-content:center;font-family:inherit}" +
".sm-close:hover{background:var(--sm-hover);color:var(--sm-ink)}" +
".sm-body{flex:1;overflow-y:auto;padding:8px}" +
".sm-section{font-size:10px;font-weight:700;letter-spacing:.06em;" +
"color:var(--sm-ink-faint);padding:10px 10px 4px}" +
".sm-divider{height:1px;background:var(--sm-line);margin:6px 8px}" +
".sm-item{display:flex;align-items:center;gap:10px;width:100%;" +
"border:none;background:transparent;padding:8px 10px;border-radius:8px;" +
"font-size:13px;color:var(--sm-ink);cursor:pointer;text-align:left;" +
"text-decoration:none;font-family:inherit;box-sizing:border-box;" +
"transition:background .12s}" +
".sm-item:hover{background:var(--sm-hover)}" +
".sm-item.sm-current{background:#E4F2ED;color:var(--sm-accent);font-weight:600}" +
".sm-item .sm-icon{width:18px;text-align:center;flex-shrink:0;font-size:14px}" +
".sm-item .sm-ext{margin-left:auto;color:var(--sm-ink-faint);font-size:10px}" +
".sm-note{padding:10px;font-size:11px;color:var(--sm-ink-faint);text-align:center}" +
".sm-foot{border-top:1px solid var(--sm-line);padding:10px 14px;" +
"font-size:10px;color:var(--sm-ink-faint)}" +
".sm-toast{position:fixed;left:12px;right:12px;bottom:14px;z-index:calc(var(--sm-z) + 2);background:#1C2733;color:#fff;" +
"padding:10px 12px;border-radius:8px;font-size:12px;line-height:1.5;font-family:var(--sm-font);" +
"box-shadow:0 6px 20px rgba(28,39,51,.3);opacity:0;visibility:hidden;transition:opacity .2s,visibility .2s}" +
".sm-toast.sm-show{opacity:1;visibility:visible}.sm-toast.sm-err{background:#8A1F14}" +
"@media (prefers-reduced-motion:reduce){.sm-panel,.sm-overlay,.sm-toast{transition:none}}";

  function injectCSS() {
    if (document.getElementById("sm-style")) return;
    var st = document.createElement("style");
    st.id = "sm-style";
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  /* ---------------- menu.json 取得 ---------------- */
  function fetchMenu(force) {
    if (menuData && !force) return Promise.resolve(menuData);
    if (fetching && !force) return fetching;

    if (!config || !config.menuUrl) {
      return Promise.resolve(null);
    }

    // キャッシュ更新を確実にするためタイムスタンプを付与
    var url = config.menuUrl + (config.menuUrl.indexOf("?") < 0 ? "?" : "&") + "t=" + Date.now();

    fetching = fetch(url)
      .then(function (res) {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
      })
      .then(function (json) {
        menuData = json;
        fetching = null;
        return menuData;
      })
      .catch(function (err) {
        console.warn("SlideMenu: menu.json の取得に失敗:", err);
        fetching = null;
        return null; // 失敗してもローカル項目だけで表示する
      });

    return fetching;
  }

  /* ---------------- DOM ---------------- */
  var overlay, panel, bodyEl;

  function build() {
    if (built) return;
    injectCSS();

    overlay = document.createElement("div");
    overlay.className = "sm-overlay";
    overlay.addEventListener("click", close);

    panel = document.createElement("nav");
    panel.className = "sm-panel " + (config.position === "right" ? "sm-right" : "sm-left");
    panel.setAttribute("aria-label", config.appName || "メニュー");
    panel.style.setProperty("--sm-width", (config.width || 250) + "px");
    if (config.theme && config.theme.accent) {
      panel.style.setProperty("--sm-accent", config.theme.accent);
    }

    // ヘッダー
    var head = document.createElement("div");
    head.className = "sm-head";
    var mark = document.createElement("div");
    mark.className = "sm-app-mark";
    mark.textContent = (config.appName || "M").charAt(0).toUpperCase();
    var nameWrap = document.createElement("div");
    var nm = document.createElement("div");
    nm.className = "sm-app-name";
    nm.textContent = config.appName || "";
    nameWrap.appendChild(nm);
    if (config.version) {
      var vr = document.createElement("div");
      vr.className = "sm-app-ver";
      vr.textContent = config.version;
      nameWrap.appendChild(vr);
    }
    var cls = document.createElement("button");
    cls.className = "sm-close";
    cls.setAttribute("aria-label", "メニューを閉じる");
    cls.textContent = "✕";
    cls.addEventListener("click", close);
    head.appendChild(mark);
    head.appendChild(nameWrap);
    head.appendChild(cls);

    bodyEl = document.createElement("div");
    bodyEl.className = "sm-body";

    panel.appendChild(head);
    panel.appendChild(bodyEl);

    if (config.footer) {
      var ft = document.createElement("div");
      ft.className = "sm-foot";
      ft.textContent = config.footer;
      panel.appendChild(ft);
    }

    document.body.appendChild(overlay);
    document.body.appendChild(panel);

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && opened) close();
    });

    built = true;
  }

  /* ローカル項目（アプリ固有）＋ menu.json 項目を統合して描画 */
  function render(data) {
    bodyEl.innerHTML = "";

    // 1) アプリ固有の項目（config.localItems）
    (config.localItems || []).forEach(function (it) {
      appendItem(it);
    });

    // 2) menu.json の項目
    if (data && Array.isArray(data.sections)) {
      if ((config.localItems || []).length > 0) {
        appendItem({ divider: true });
      }
      data.sections.forEach(function (sec) {
        if (sec.title) appendItem({ section: sec.title });
        (sec.items || []).forEach(function (raw) {
          appendItem({
            label: raw.label,
            icon: raw.icon,
            href: raw.popup ? null : raw.url,          // popup 指定の項目は画面遷移させない
            target: raw.external ? "_blank" : null,
            popup: raw.popup || null,
            onClick: raw.popup ? function () { openPopup(raw); } : undefined,
            current: config.currentId && raw.id === config.currentId
          });
        });
      });
    } else if (config.menuUrl) {
      // JSON取得失敗時の表示
      var note = document.createElement("div");
      note.className = "sm-note";
      note.textContent = "メニュー一覧を取得できませんでした";
      bodyEl.appendChild(note);
    }
  }

  function appendItem(it) {
    if (it.divider) {
      var dv = document.createElement("div");
      dv.className = "sm-divider";
      bodyEl.appendChild(dv);
      return;
    }
    if (it.section) {
      var sc = document.createElement("div");
      sc.className = "sm-section";
      sc.textContent = it.section;
      bodyEl.appendChild(sc);
      return;
    }

    var el;
    if (it.href) {
      el = document.createElement("a");
      el.href = it.href;
      if (it.target) el.target = it.target;
    } else {
      el = document.createElement("button");
      el.type = "button";
    }
    el.className = "sm-item" + (it.current ? " sm-current" : "");

    var ic = document.createElement("span");
    ic.className = "sm-icon";
    ic.textContent = it.icon || "";
    el.appendChild(ic);

    var lb = document.createElement("span");
    lb.textContent = it.label || "";
    el.appendChild(lb);

    if (it.target === "_blank") {
      var ex = document.createElement("span");
      ex.className = "sm-ext";
      ex.textContent = "↗";
      el.appendChild(ex);
    } else if (it.popup) {
      var px = document.createElement("span");
      px.className = "sm-ext";
      px.title = "ポップアップで開きます";
      px.textContent = "⧉";
      el.appendChild(px);
    }

    el.addEventListener("click", function () {
      if (typeof it.onClick === "function") it.onClick();
      if (!it.href) close();
    });

    bodyEl.appendChild(el);
  }

  /* ---------------- ポップアップ項目（画面遷移せずにダイアログで開く） ---------------- */
  function commonBase() {
    var u = SELF_SRC || (config && config.menuUrl) || "";
    return u.replace(/[?#].*$/, "").replace(/\/[^\/]*$/, "");
  }

  function toast(msg, isErr) {
    injectCSS();
    var t = document.getElementById("sm-toast");
    if (!t) {
      t = document.createElement("div");
      t.id = "sm-toast";
      t.className = "sm-toast";
      t.setAttribute("role", "status");
      document.body.appendChild(t);
    }
    t.textContent = msg;
    t.className = "sm-toast sm-show" + (isErr ? " sm-err" : "");
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.className = "sm-toast"; }, isErr ? 9000 : 3500);
  }

  function loadBridge() {
    if (global.ScheduleBridge) return Promise.resolve();
    if (bridgeLoading) return bridgeLoading;
    bridgeLoading = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = commonBase() + "/schedule-bridge.js?v=" + Date.now();
      s.onload = function () { resolve(); };
      s.onerror = function () { bridgeLoading = null; reject(new Error("schedule-bridge.js を読み込めませんでした")); };
      document.head.appendChild(s);
    });
    return bridgeLoading;
  }

  function popupError(e) {
    var c = e && e.code;
    if (c === 12007) return "すでにポップアップが開いています。";
    if (c === 12009) return "ポップアップの表示が許可されませんでした。もう一度押して、許可してください。";
    return "ポップアップを開けませんでした：" + ((e && (e.message || e.name)) || "不明なエラー") + (c ? "（" + c + "）" : "");
  }

  function openPopup(item) {
    var id = item.id || item.popup;
    if (typeof Office === "undefined" || !Office.context || !Office.context.ui) {
      toast("Excelのアドインの中でのみ開けます。", true);
      return;
    }
    if (Office.context.requirements && !Office.context.requirements.isSetSupported("DialogApi", "1.2")) {
      toast("このExcelはポップアップとの双方向通信（DialogApi 1.2）に対応していません。", true);
      return;
    }
    if (popupHandles[id]) {
      toast("すでに開いています。見つからないときは、ポップアップを閉じてからもう一度押してください。");
      return;
    }
    toast("ポップアップを開いています…");
    loadBridge().then(function () {
      var done = false, hadErr = false;
      var h = global.ScheduleBridge.open({
        url: item.popup,
        width: item.width || 95,
        height: item.height || 92,
        sheetName: item.sheetName,
        masterName: item.masterName,
        onClose: function () {
          done = true;
          delete popupHandles[id];
          if (!hadErr) toast("ポップアップを閉じました。");
        },
        onError: function (e) { hadErr = true; toast(popupError(e), true); }
      });
      if (!done) popupHandles[id] = h;
    }).catch(function (e) {
      toast(String((e && e.message) || e), true);
    });
  }

  // このページを離れる（別ツールへ遷移・作業ウィンドウを閉じる）と、仲介が止まるのでポップアップも閉じる
  global.addEventListener("pagehide", function () {
    Object.keys(popupHandles).forEach(function (k) { try { popupHandles[k].close(); } catch (e) { /* 無視 */ } });
  });

  /* ---------------- 開閉 ---------------- */
  function open() {
    if (!config) { console.warn("SlideMenu: init() を先に呼んでください"); return; }
    build();

    // まずローカル項目だけで即表示 → JSON到着後に差し替え
    render(menuData);
    lastFocus = document.activeElement;
    overlay.classList.add("sm-show");
    panel.classList.add("sm-show");
    opened = true;

    fetchMenu().then(function (data) {
      if (opened && data) render(data);
    });

    var first = panel.querySelector(".sm-item, .sm-close");
    if (first) first.focus();
  }

  function close() {
    if (!built) return;
    overlay.classList.remove("sm-show");
    panel.classList.remove("sm-show");
    opened = false;
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  /* ---------------- 公開API ---------------- */
  global.SlideMenu = {
    init: function (cfg) {
      config = cfg || {};
      if (built) {
        overlay.remove();
        panel.remove();
        built = false;
      }
    },
    open: open,
    close: close,
    toggle: function () { opened ? close() : open(); },
    isOpen: function () { return opened; },
    openPopup: openPopup,
    reload: function () {
      return fetchMenu(true).then(function (data) {
        if (built) render(data);
        return data;
      });
    }
  };
})(window);
