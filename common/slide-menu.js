/* ============================================================
 * slide-menu.js — 共通スライドメニューコンポーネント（JSON駆動）
 * ------------------------------------------------------------
 * メニューの「名前・URL」は menu.json で一元管理する。
 * menu.json を編集すれば、組み込んでいる全アプリのメニューが
 * 一斉に更新される。CSS・DOMはこのJSが自動注入するため、
 * ホスト側にHTML/CSSの追加は不要。
 *
 * ■ 配置
 *   tools/common/slide-menu.js   （このファイル）
 *   tools/common/menu.json       （メニュー定義）
 *
 * ■ ホスト側の組み込み（遅延ロードの例）
 *
 *   const COMMON = 'https://ymatsuda-cmyk.github.io/tools/common';
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
 * ■ 公開API
 *   SlideMenu.init(config)   初期化（再実行で設定を更新）
 *   SlideMenu.open()         開く（初回はmenu.jsonを取得）
 *   SlideMenu.close()
 *   SlideMenu.toggle()
 *   SlideMenu.isOpen()
 *   SlideMenu.reload()       menu.jsonを再取得して描画し直す
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
"@media (prefers-reduced-motion:reduce){.sm-panel,.sm-overlay{transition:none}}";

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
    panel.className = "sm-panel " + (config.position === "left" ? "sm-left" : "sm-right");
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
            href: raw.url,
            target: raw.external ? "_blank" : null,
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
    }

    el.addEventListener("click", function () {
      if (typeof it.onClick === "function") it.onClick();
      if (!it.href) close();
    });

    bodyEl.appendChild(el);
  }

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
    reload: function () {
      return fetchMenu(true).then(function (data) {
        if (built) render(data);
        return data;
      });
    }
  };
})(window);
