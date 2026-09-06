/* ============================================================
 * roi-core.js — 議事録・ROI試算の共有ライブラリ
 * ------------------------------------------------------------
 * 「営業報告」アドインと「提案ナレッジ」アドインの両方から読み込む。
 * 同じGitHub Pagesドメイン配下に置くこと（両アドインとも
 * ymatsuda-cmyk.github.io なので相互に script タグで読み込める）。
 *
 * このファイルは DOM を一切触らない「データ層」。
 * ・営業報告アドイン　　→ 簡単な操作のみ（紐づけ・参照・提案作成・提案参照）
 * ・提案ナレッジアドイン → ROIマスタ編集、抽出結果の手動レビュー、
 *                        プロンプト組み立てなど詳細機能
 * という役割分担で、両方がこのファイルの関数を呼ぶ。
 *
 * 読み込み例:
 *   <script src="https://ymatsuda-cmyk.github.io/tools/addin/roi-knowledge/roi-core.js"></script>
 *   window.RoiCore.xxx(...) で呼び出す。
 * ============================================================ */

(function (global) {

  const EIGYO_SHEET = "営業報告";
  const CUST_SHEET = "顧客マスタ";

  const MASTER_SHEET = "ROIマスタ";
  const MASTER_COLUMNS = ["課題カテゴリ", "項目ID", "項目名", "区分", "単位", "デフォルト値", "数式", "信頼度初期値"];

  const HEARING_SHEET = "議事録";
  // 参照URL: PLAUD/Notion等、既存の議事録ビューアに保管されているテキストへの
  // リンクだけを持たせるケースを想定した列。発言テキストは空でもよい。
  // 議事録ID: ROI試算側から「どの議事録から抽出したか」を辿るための一意キー。
  // タイトル: 「初回訪問ヒアリング」など、1案件に複数の議事録を並べたときの識別用。
  const HEARING_COLUMNS = ["案件ID", "タイトル", "参照URL", "発言テキスト", "登録日時", "議事録ID"];

  const CALC_SHEET = "ROI試算";
  // 根拠議事録ID: そのカテゴリを作成するときに参照した議事録IDのカンマ区切り。
  // 「提案書データの元データリンク一覧」表示に使う。
  const CALC_COLUMNS = ["案件ID", "課題カテゴリ", "項目ID", "項目名", "区分", "値", "単位", "信頼度", "選択", "更新日時", "根拠議事録ID"];

  const SOLUTION_SHEET = "ソリューションDB";
  // 1つの課題カテゴリに対して複数の解決案を持たせる（低コスト〜高コストまで並べる）。
  // コスト感: 低 / 中 / 高 の3段階。実現方法: 定性的な説明文。
  // 概算導入費・改善率はあくまで初期値で、案件ごとにROI試算シート側で上書きできる。
  // 根拠区分: 実績（自社導入事例あり） / 推定（類似案件から） / 一般値（業界一般値）
  //   → 削減額の確からしさを示す。入力値の信頼度とあわせて全体の信頼度に反映する。
  const SOLUTION_COLUMNS = ["課題カテゴリ", "解決策名", "コスト感", "実現方法", "概算導入費", "改善率", "削減根拠", "根拠区分"];

  const BASIS_LEVELS = ["実績", "推定", "一般値"];
  // 根拠区分の重み。信頼度計算で入力値の 確定/推定/未確認 と同じ尺度に載せる。
  const BASIS_WEIGHT = { "実績": 1, "推定": 0.5, "一般値": 0.25 };

  // AIが議事録から抽出した課題（カテゴリ・タイトル・内容）。
  // 1つの議事録から複数カテゴリが抽出されるため、案件ID×課題カテゴリで1行。
  // 既存ワークブックの「課題」シート（課題管理表）とは別物なので名前を分けている。
  const ISSUE_SHEET = "抽出課題";
  // AI原文タイトル / AI原文内容: AIが最後に出力したそのままの文章を保持する。
  // 現在の課題タイトル・課題内容がこれと異なれば「営業が編集した」と判定し、
  // 再抽出時に自動で上書きせず確認を出す。
  const ISSUE_COLUMNS = ["案件ID", "課題カテゴリ", "課題タイトル", "課題内容",
    "AI原文タイトル", "AI原文内容", "根拠議事録ID", "抽出日時"];

  // 案件ごとの「どの解決策に決めたか」を保持するシート。
  // 導入費・改善率・削減根拠はソリューションDBの値を初期値としてコピーし、
  // 営業側が案件ごとに書き換えられるようにする（マスタ側は変更しない）。
  const DECISION_SHEET = "提案決定";
  const DECISION_COLUMNS = ["案件ID", "課題カテゴリ", "解決策名", "導入費", "改善率", "削減根拠", "根拠区分", "提案に含める", "更新日時"];

  const CONF_LEVELS = ["確定", "推定", "未確認"];
  const CONF_WEIGHT = { "確定": 1, "推定": 0.5, "未確認": 0 };

  const MASTER_SEED = [
    ["在庫管理", "stk_people", "棚卸人数", "入力", "人", 3, "", "未確認"],
    ["在庫管理", "stk_hours", "棚卸時間", "入力", "時間", 4, "", "未確認"],
    ["在庫管理", "stk_freq", "棚卸回数/年", "入力", "回", 12, "", "未確認"],
    ["在庫管理", "stk_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["在庫管理", "stk_improve", "改善率", "入力", "%", 60, "", "推定"],
    ["在庫管理", "stk_hours_yr", "年間棚卸工数", "出力", "時間", "", "stk_people*stk_hours*stk_freq", ""],
    ["在庫管理", "stk_cost_yr", "年間棚卸コスト", "出力", "円", "", "stk_hours_yr*stk_wage", ""],
    ["在庫管理", "stk_saving", "削減額", "出力", "円", "", "stk_cost_yr*stk_improve/100", ""],
    ["ロット管理", "lot_hours", "追跡時間", "入力", "時間", 2, "", "未確認"],
    ["ロット管理", "lot_freq", "追跡回数/年", "入力", "回", 100, "", "未確認"],
    ["ロット管理", "lot_people", "担当人数", "入力", "人", 2, "", "未確認"],
    ["ロット管理", "lot_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["ロット管理", "lot_improve", "改善率", "入力", "%", 60, "", "推定"],
    ["ロット管理", "lot_hours_yr", "年間追跡工数", "出力", "時間", "", "lot_hours*lot_freq*lot_people", ""],
    ["ロット管理", "lot_saving", "削減額", "出力", "円", "", "lot_hours_yr*lot_wage*lot_improve/100", ""],
    ["AI議事録", "min_meetings", "会議回数/月", "入力", "回", 8, "", "未確認"],
    ["AI議事録", "min_people", "参加人数", "入力", "人", 3, "", "未確認"],
    ["AI議事録", "min_hours", "議事録作成時間", "入力", "時間", 1, "", "未確認"],
    ["AI議事録", "min_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["AI議事録", "min_hours_yr", "年間工数", "出力", "時間", "", "min_meetings*12*min_people*min_hours", ""],
    ["AI議事録", "min_saving", "削減額", "出力", "円", "", "min_hours_yr*min_wage", ""],
    // 製造管理
    ["製造管理", "mfg_revenue", "生産額", "入力", "円", 50000000, "", "未確認"],
    ["製造管理", "mfg_defect_rate", "不良率", "入力", "%", 3, "", "推定"],
    ["製造管理", "mfg_defect_rate_after", "改善後不良率", "入力", "%", 1, "", "推定"],
    ["製造管理", "mfg_loss", "不良損失", "出力", "円", "", "mfg_revenue*mfg_defect_rate/100", ""],
    ["製造管理", "mfg_loss_after", "改善後不良損失", "出力", "円", "", "mfg_revenue*mfg_defect_rate_after/100", ""],
    ["製造管理", "mfg_saving", "削減額", "出力", "円", "", "mfg_loss-mfg_loss_after", ""],
    // OCR受注入力
    ["OCR受注入力", "ocr_count", "注文書件数/月", "入力", "件", 200, "", "未確認"],
    ["OCR受注入力", "ocr_input_min", "入力時間", "入力", "分", 5, "", "未確認"],
    ["OCR受注入力", "ocr_fix_min", "修正時間", "入力", "分", 2, "", "未確認"],
    ["OCR受注入力", "ocr_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["OCR受注入力", "ocr_improve", "改善率", "入力", "%", 70, "", "推定"],
    ["OCR受注入力", "ocr_hours_yr", "年間入力時間", "出力", "時間", "", "ocr_count*12*(ocr_input_min+ocr_fix_min)/60", ""],
    ["OCR受注入力", "ocr_saving", "削減額", "出力", "円", "", "ocr_hours_yr*ocr_wage*ocr_improve/100", ""],
    // Delphi移行
    ["Delphi移行", "delphi_staff", "保守担当人数", "入力", "人", 2, "", "未確認"],
    ["Delphi移行", "delphi_inquiries", "問い合わせ件数/年", "入力", "件", 300, "", "未確認"],
    ["Delphi移行", "delphi_avg_hours", "平均対応時間", "入力", "時間", 1, "", "推定"],
    ["Delphi移行", "delphi_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["Delphi移行", "delphi_improve", "改善率", "入力", "%", 50, "", "推定"],
    ["Delphi移行", "delphi_hours_yr", "年間保守工数", "出力", "時間", "", "delphi_inquiries*delphi_avg_hours", ""],
    ["Delphi移行", "delphi_cost_yr", "年間保守コスト", "出力", "円", "", "delphi_hours_yr*delphi_wage", ""],
    ["Delphi移行", "delphi_saving", "削減額", "出力", "円", "", "delphi_cost_yr*delphi_improve/100", ""],
    // 属人化・情報共有・問合せ対応: ROI試算は行わず、定性的な解決案候補のみ持つ
    // カテゴリ一覧に出てくるよう最低限の入力項目だけ用意している
    ["属人化", "attrib_people", "該当ベテラン人数", "入力", "人", 1, "", "未確認"],
    ["情報共有", "info_tools", "使用ツール数", "入力", "個", 3, "", "未確認"],
    ["問合せ対応", "inquiry_count", "月間問合せ件数", "入力", "件", 20, "", "未確認"],
  ];

  const SOLUTION_SEED = [
    ["在庫管理", "Excelテンプレ改善", "低", "既存Excelにマクロと入力規則を追加し、集計だけ自動化する", 800000, 20, "集計作業のみ自動化され、現物カウント作業は残るため。類似規模2社の実績から推定", "推定"],
    ["在庫管理", "在庫管理システム導入", "中", "既製クラウドSaaSを導入し、棚卸をハンディ端末で読み取り自動集計する", 3000000, 50, "カウント作業がハンディ読み取りに置き換わり、集計・差異照合が自動化される", "実績"],
    ["在庫管理", "専用システム新規開発", "高", "在庫・ロット・原価まで一気通貫の基幹システムとして開発する", 6800000, 70, "受注〜出荷まで転記がなくなる前提。自社実績がなく業界一般値のため幅を持たせて説明すること", "一般値"],
    ["ロット管理", "台帳テンプレ改善", "低", "Excel台帳にロットIDのバーコード読み取りと自動採番を追加する", 600000, 25, "採番と検索が速くなるが、現物との突合は人手のまま残るため", "推定"],
    ["ロット管理", "ロット管理システム化", "中", "追跡専用のクラウドサービスを導入し、トレーサビリティを確保する", 2800000, 55, "追跡依頼への回答が台帳検索で完結する。食品製造2社で半日→30分の実績", "実績"],
    ["ロット管理", "トレーサビリティ基幹システム新規開発", "高", "原材料〜出荷までを一気通貫で追跡できる専用システムを新規構築する", 7200000, 75, "全工程が自動記録される前提。自社実績がなく業界一般値", "一般値"],
    ["製造管理", "検査記録のデジタル化", "低", "紙の検査記録をExcel・タブレット入力に置き換え、不良傾向を集計しやすくする", 900000, 20, "記録・集計の手間は減るが、不良の発生自体は減らないため効果は限定的", "推定"],
    ["製造管理", "製造管理システム導入", "中", "工程・検査データを一元管理し、不良率の見える化と是正を迅速化する", 4000000, 40, "不良の早期検知により手戻りが減る。部品製造1社で不良率3%→1.8%の実績", "実績"],
    ["製造管理", "MES・生産管理システム新規開発", "高", "ライン設備と連携した生産管理システムを新規構築する", 12000000, 60, "設備連携による自動検知を前提とした業界一般値。設備条件に大きく依存する", "一般値"],
    ["AI議事録", "AI議事録", "低", "既存の会議ツールにAI議事録サービスを連携し、文字起こし・要約・タスク抽出を自動化する", 600000, 70, "議事録作成の清書作業がほぼ不要になる。自社導入で作成時間1時間→15分の実績", "実績"],
    ["AI議事録", "社内ナレッジ基盤との統合", "中", "議事録を検索可能なナレッジベースに蓄積し、過去案件を横断検索できるようにする", 2000000, 80, "作成に加え、過去議事録の検索時間も削減される想定。検索工数は未計測のため推定", "推定"],
    ["OCR受注入力", "OCR・入力チェックツール導入", "低", "既存のFAX・PDF注文書をOCRで読み取り、人はチェックのみ行う", 1200000, 50, "入力はなくなるが目視チェックは残る。読み取り精度により変動する", "推定"],
    ["OCR受注入力", "受発注システム連携", "中", "OCR結果を受発注システムに自動連携し、二重入力をなくす", 3500000, 70, "OCR後の再入力が不要になる。卸1社で入力工数7割減の実績", "実績"],
    ["OCR受注入力", "EDI・Web受注への切り替え", "高", "取引先とのやり取り自体をEDIやWeb発注フォームに置き換える", 6000000, 90, "紙受注そのものがなくなる前提。取引先の協力度合いに全面的に依存する", "一般値"],
    ["Delphi移行", "保守体制の見直し", "低", "現行Delphiシステムの保守体制・対応フローを整理し、対応時間を短縮する", 500000, 15, "問い合わせ対応の手順整理による短縮のみ。システム自体は変わらない", "推定"],
    ["Delphi移行", "画面のみWeb化", "中", "利用頻度の高い画面から順にWebシステムへ段階移行する", 4500000, 45, "移行済み画面の保守が軽くなる。移行範囲に比例するため案件ごとに要調整", "推定"],
    ["Delphi移行", "フルWeb移行", "高", "Delphiシステム全体をWebシステムとして再構築する", 15000000, 80, "Delphi固有の保守作業がなくなる。移行実績1社で保守工数8割減", "実績"],
    ["属人化", "業務手順書の作成", "低", "ベテランの作業手順を文書化し、引き継ぎリスクを下げる", 400000, 0, "金額換算していない。引き継ぎ時のリスク低減が主目的", "一般値"],
    ["属人化", "動画マニュアル化", "中", "作業風景を動画で残し、新人教育に活用する", 900000, 0, "金額換算していない。教育期間の短縮が主目的", "一般値"],
    ["情報共有", "既存ツール運用の整理", "低", "既存の共有フォルダ・チャットツールの使い方を整理し、情報の置き場所を統一する", 300000, 0, "金額換算していない。運用ルール整備が主目的", "一般値"],
    ["情報共有", "社内ポータル導入", "中", "情報共有専用のポータルサイトを構築する", 2500000, 0, "金額換算していない。探す時間の削減は未計測", "一般値"],
    ["問合せ対応", "FAQ整備", "低", "よくある問合せをFAQ化し、一次対応を減らす", 400000, 30, "一次対応の3割程度がFAQで自己解決する想定。問合せ内容により変動", "推定"],
    ["問合せ対応", "チャットボット導入", "中", "FAQを元にしたチャットボットで一次対応を自動化する", 1800000, 50, "FAQ整備済みが前提。自社実績がなく業界一般値", "一般値"],
  ];

  let masterItems = null; // 初回 ensureAllSheets() 後にキャッシュ

  /* ---------- 設定（AIエンドポイント） ----------
   * localStorage にはエンドポインURLと簡易トークンのみ保存する。
   * APIキー本体はサーバ（GAS等）側にのみ保持し、ここには保存しない。
   * 営業報告・提案ナレッジの両アドインは同一オリジンなので localStorage を共有する。 */
  function getConfig() {
    try { return JSON.parse(localStorage.getItem("roiAddinConfig") || "{}"); }
    catch (e) { return {}; }
  }
  function setConfig(cfg) { localStorage.setItem("roiAddinConfig", JSON.stringify(cfg)); }

  /* ---------- シートの自動作成 ---------- */
  function colLetterOf(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  async function getOrCreateSheet(ctx, name, columns, seedRows) {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let ws = sheets.items.find(s => s.name === name);
    if (!ws) {
      ws = sheets.add(name);
      const lastCol = colLetterOf(columns.length);
      const hdr = ws.getRange(`A1:${lastCol}1`);
      hdr.values = [columns];
      hdr.format.fill.color = "#44546A";
      hdr.format.font.color = "#FFFFFF";
      hdr.format.font.bold = true;
      if (seedRows && seedRows.length) ws.getRange(`A2:${lastCol}${seedRows.length + 1}`).values = seedRows;
      await ctx.sync();
    }
    return ws;
  }

  function rowToMasterItem(r) {
    return { category: r[0], itemId: r[1], name: r[2], kind: r[3], unit: r[4], defaultVal: r[5], formula: r[6], confDefault: r[7] };
  }

  /* すべての必要シートを用意し、ROIマスタをキャッシュして返す。
   * 営業報告・提案ナレッジどちらの起動時にも最初に呼ぶ。 */
  async function ensureAllSheets() {
    if (!global.Office || !global.Excel) {
      masterItems = MASTER_SEED.map(rowToMasterItem);
      return { demo: true, masterItems };
    }
    await Excel.run(async ctx => {
      await getOrCreateSheet(ctx, MASTER_SHEET, MASTER_COLUMNS, MASTER_SEED);
      await getOrCreateSheet(ctx, HEARING_SHEET, HEARING_COLUMNS, null);
      await getOrCreateSheet(ctx, CALC_SHEET, CALC_COLUMNS, null);
      await getOrCreateSheet(ctx, SOLUTION_SHEET, SOLUTION_COLUMNS, SOLUTION_SEED);
      await getOrCreateSheet(ctx, ISSUE_SHEET, ISSUE_COLUMNS, null);
      await getOrCreateSheet(ctx, DECISION_SHEET, DECISION_COLUMNS, null);
      const rng = ctx.workbook.worksheets.getItem(MASTER_SHEET).getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      masterItems = rng.values.slice(1).filter(r => r[1]).map(rowToMasterItem);
    });
    return { demo: false, masterItems };
  }

  function getMasterItems() { return masterItems || []; }
  function getCategories() { return Array.from(new Set(getMasterItems().map(m => m.category))); }

  /* 案件IDの候補。営業報告シートのID列に加え、顧客マスタから顧客名、
   * 営業報告シートの案件名列（あれば）を引いてラベルを組み立てる。 */
  async function listCaseIds() {
    if (!global.Office || !global.Excel) return [];
    const out = [];
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      if (!sheets.items.find(s => s.name === EIGYO_SHEET)) return;
      const sheet = ctx.workbook.worksheets.getItem(EIGYO_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values");
      const custSheet = sheets.items.find(s => s.name === CUST_SHEET)
        ? ctx.workbook.worksheets.getItem(CUST_SHEET).getUsedRange(true) : null;
      if (custSheet) custSheet.load("values");
      await ctx.sync();

      const custMap = {};
      if (custSheet) custSheet.values.slice(1).forEach(r => { if (r[0]) custMap[String(r[0])] = r[1] || ""; });

      const header = used.values[0] || [];
      // 案件名らしき列を探す（見つからなければラベルは顧客名のみ）
      const titleCol = header.findIndex(h => /案件名|件名|タイトル|概要/.test(String(h)));
      const seen = new Set();
      used.values.slice(1).forEach(r => {
        const id = r[0] ? String(r[0]) : "";
        if (!id || seen.has(id)) return;
        seen.add(id);
        const custName = custMap[id.split("-")[0]] || "";
        const title = titleCol >= 0 ? String(r[titleCol] || "") : "";
        const label = [id, [custName, title].filter(Boolean).join(" ・")].filter(Boolean).join(" ／ ");
        out.push({ caseId: id, customer: custName, title, label });
      });
    });
    return out;
  }

  /* ---------- 議事録：紐づけ・参照 ---------- */
  async function appendHearingLog(caseId, title, { text = "", url = "" } = {}) {
    const hearingId = genId("H");
    const row = [caseId, title || "議事録", url, text, nowStr(), hearingId];
    if (!global.Office || !global.Excel) return hearingId;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const nextRow = Math.max(used.rowCount, 1) + 1;
      sheet.getRange(`A${nextRow}:F${nextRow}`).values = [row];
      await ctx.sync();
    });
    return hearingId;
  }

  async function listHearingLogs(caseId) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
    return rows.map(r => ({
      title: r[1] || "議事録", url: r[2], text: r[3],
      registeredAt: r[4], hearingId: r[5],
    }));
  }

  /* 議事録が1件以上ある案件IDの一覧（未抽出案件をグレー表示するために使う） */
  async function listCasesWithHearings() {
    if (!global.Office || !global.Excel) return [];
    let hearing = [], issues = [];
    await Excel.run(async ctx => {
      const h = ctx.workbook.worksheets.getItem(HEARING_SHEET).getUsedRange(true);
      h.load("values");
      const i = ctx.workbook.worksheets.getItem(ISSUE_SHEET).getUsedRange(true);
      i.load("values");
      await ctx.sync();
      hearing = h.values.slice(1).filter(r => r[0]);
      issues = i.values.slice(1).filter(r => r[0] && r[1]);
    });
    const map = {};
    hearing.forEach(r => {
      const id = String(r[0]);
      map[id] = map[id] || { caseId: id, hearingCount: 0, issueCount: 0 };
      map[id].hearingCount++;
    });
    issues.forEach(r => {
      const id = String(r[0]);
      map[id] = map[id] || { caseId: id, hearingCount: 0, issueCount: 0 };
      map[id].issueCount++;
    });
    return Object.values(map);
  }

  function genId(prefix) {
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  }

  /* ---------- AI抽出（GAS等のWebhook） ----------
   * url が指定されている場合はサーバ側（GAS）でページ内容を取得させる。
   * レスポンス: { items: [{ itemId, value, confidence }] } */
  async function callExtractionWebhook(caseId, category, { text = "", url = "" } = {}) {
    const cfg = getConfig();
    if (!cfg.webhookUrl) throw new Error("AI連携エンドポイントが未設定です");
    const items = getMasterItems().filter(m => m.category === category && m.kind === "入力");
    const res = await fetch(cfg.webhookUrl, {
      method: "POST",
      body: JSON.stringify({
        token: cfg.token || "",
        caseId, category, text, url,
        items: items.map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
      }),
    });
    const raw = await res.text();
    console.log("[RoiCore] extraction raw response:", raw);
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      throw new Error("サーバーの応答がJSONではありません（GASのデプロイ設定を確認してください）。実際の応答はconsoleに出力しています。");
    }
    if (data.error) throw new Error("GASエラー: " + data.error);
    return data.items || [];
  }

  /* ---------- ROIマスタの数式(項目IDのトークン式)をExcel数式に変換 ---------- */
  function translateFormula(masterFormula, rowNum) {
    if (!masterFormula) return "";
    const translated = masterFormula.replace(/[a-zA-Z_][a-zA-Z0-9_]*/g, tok =>
      `INDEX($F$2:$F$9999,MATCH(1,($A$2:$A$9999=$A${rowNum})*($C$2:$C$9999="${tok}"),0))`);
    return "=" + translated;
  }

  /* ---------- ROI試算シートへの反映（カテゴリ丸ごと） ----------
   * 1. 対象カテゴリの全項目がこの案件の行としてまだ無ければROIマスタからコピー
   *   （出力行は数式化）。
   * 2. 渡された抽出値・信頼度を、該当する入力行の値・信頼度に上書きする。 */
  async function applyCategoryToCalcSheet(caseId, category, extractedItems, hearingIds = []) {
    if (!global.Office || !global.Excel) return;
    const provenance = hearingIds.filter(Boolean).join(",");
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values, rowCount");
      await ctx.sync();

      const existingRows = used.values.slice(1);
      const keyOf = r => `${r[0]}__${r[2]}`;
      const existingIndex = {};
      existingRows.forEach((r, i) => { existingIndex[keyOf(r)] = i + 2; });

      const catItems = getMasterItems().filter(m => m.category === category);
      const toAppend = [];
      const nowTs = nowStr();

      catItems.forEach(m => {
        const key = `${caseId}__${m.itemId}`;
        if (existingIndex[key]) return;
        const rowNum = used.rowCount + 1 + toAppend.length;
        let valueCell;
        if (m.kind === "出力") {
          valueCell = translateFormula(m.formula, rowNum);
        } else {
          const ex = extractedItems.find(e => e.itemId === m.itemId);
          valueCell = ex && ex.value != null ? ex.value : m.defaultVal;
        }
        const confCell = m.kind === "出力" ? "" : ((extractedItems.find(e => e.itemId === m.itemId) || {}).confidence || m.confDefault);
        toAppend.push([caseId, m.category, m.itemId, m.name, m.kind, valueCell, m.unit, confCell, "FALSE", nowTs, provenance]);
      });

      if (toAppend.length) {
        const startRow = used.rowCount + 1;
        const endRow = startRow + toAppend.length - 1;
        const rangeAll = sheet.getRange(`A${startRow}:K${endRow}`);
        rangeAll.values = toAppend.map(r => r.map((v, i) => (i === 5 && typeof v === "string" && v.startsWith("=")) ? "" : v));
        const formulaCol = sheet.getRange(`F${startRow}:F${endRow}`);
        formulaCol.formulas = toAppend.map(r => [typeof r[5] === "string" && r[5].startsWith("=") ? r[5] : (r[5] === "" ? "" : r[5])]);
        await ctx.sync();
      }

      const updates = [];
      extractedItems.forEach(e => {
        const rowNum = existingIndex[`${caseId}__${e.itemId}`];
        if (rowNum) updates.push({ rowNum, value: e.value, confidence: e.confidence });
      });
      updates.forEach(u => {
        sheet.getRange(`F${u.rowNum}`).values = [[u.value]];
        sheet.getRange(`H${u.rowNum}`).values = [[u.confidence]];
        sheet.getRange(`J${u.rowNum}`).values = [[nowTs]];
      });
      if (updates.length) await ctx.sync();
    });
  }

  /* ---------- 「簡単な操作」向け：抽出→即反映を1呼び出しでまとめる ----------
   * 営業報告アドインの「ROI提案」ボタンはこれだけを呼ぶ。
   * レビュー画面は挟まず、AIの抽出結果をそのまま保存する（詳細な確認・修正が
   * 必要な場合は提案ナレッジアドイン側の runExtractionForReview を使う）。 */
  async function quickCreateProposal(caseId, category, { text = "", url = "", title = "" } = {}) {
    let hearingId = null;
    if (text || url) hearingId = await appendHearingLog(caseId, title || "議事録", { text, url });
    const items = await callExtractionWebhook(caseId, category, { text, url });
    await applyCategoryToCalcSheet(caseId, category, items, hearingId ? [hearingId] : []);
    return getProposalSummaryForCase(caseId, category);
  }

  /* ---------- 複数課題の一括抽出（2段階） ----------
   * previewExtraction()  … AIを呼び、既存の抽出課題と突き合わせた差分を返す。保存はしない。
   * commitExtraction()   … 差分に対する選択（残す/上書き）を受け取って保存する。
   * 数値（ROI試算）は選択に関わらず更新する。文章だけが選択の対象。
   * 1つの議事録から複数の課題が出るのが前提。 */
  async function previewExtraction(caseId, memoText = "") {
    const cfg = getConfig();
    if (!cfg.webhookUrl) throw new Error("AI連携エンドポイントが未設定です");
    const logs = await listHearingLogs(caseId);
    const combinedText = [memoText, ...logs.map(l => l.text || l.url || "")].filter(Boolean).join("\n\n");
    if (!combinedText) return { hearingIds: [], results: [] };
    const hearingIds = logs.map(l => l.hearingId).filter(Boolean);

    const categoryDefs = getCategories().map(cat => ({
      category: cat,
      items: getMasterItems().filter(m => m.category === cat && m.kind === "入力")
        .map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
    }));

    const res = await fetch(cfg.webhookUrl, {
      method: "POST",
      body: JSON.stringify({ mode: "auto", token: cfg.token || "", caseId, text: combinedText, categories: categoryDefs }),
    });
    const raw = await res.text();
    console.log("[RoiCore] auto-extraction raw response:", raw);
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      throw new Error("サーバーの応答がJSONではありません（GASのデプロイ設定を確認してください）。実際の応答はconsoleに出力しています。");
    }
    if (data.error) throw new Error("GASエラー: " + data.error);

    const existing = await getIssues(caseId);
    const results = (data.results || []).filter(r => r && r.category).map(r => {
      const cur = existing.find(x => x.category === r.category);
      return {
        category: r.category,
        newTitle: r.title || r.category,
        newSummary: r.summary || "",
        items: r.items || [],
        current: cur || null,
        // 新規 / 編集済み（要確認） / 未編集（自動更新）
        status: !cur ? "新規" : (cur.edited ? "編集済み" : "未編集"),
        // 既定値: 編集済みは残す、それ以外は上書き
        keepText: cur ? cur.edited : false,
      };
    });
    return { hearingIds, results };
  }

  async function commitExtraction(caseId, hearingIds, results) {
    for (const r of results) {
      await saveIssue(caseId, r.category, {
        title: r.newTitle, summary: r.newSummary, hearingIds, keepText: !!r.keepText,
      });
      if (r.items && r.items.length) {
        await applyCategoryToCalcSheet(caseId, r.category, r.items, hearingIds);
      }
    }
    return results.length;
  }

  /* 確認なしで一括反映する簡易版（営業報告アドインの「作成」アイコン用）。
   * 編集済みの課題文は自動的に残す。 */
  async function autoExtractProposals(caseId, memoText = "") {
    const { hearingIds, results } = await previewExtraction(caseId, memoText);
    await commitExtraction(caseId, hearingIds, results);
    return results.map(r => ({
      category: r.category, title: r.newTitle, summary: r.newSummary,
      itemCount: (r.items || []).length, status: r.status,
    }));
  }

  /* 選択済みカテゴリの元データ（議事録）リンク一覧を返す。
   * ROI試算の「根拠議事録ID」列に記録されたIDから議事録を引く。 */
  async function getSourceEntriesForCategory(caseId, category) {
    if (!global.Office || !global.Excel) return [];
    let provenance = "";
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const row = rng.values.slice(1).find(r => String(r[0]) === caseId && r[1] === category && r[10]);
      provenance = row ? row[10] : "";
    });
    const ids = new Set(provenance.split(",").filter(Boolean));
    if (!ids.size) return [];
    const logs = await listHearingLogs(caseId);
    return logs.filter(l => ids.has(l.hearingId));
  }

  /* 詳細レビュー用：抽出だけ行い、保存はしない（提案ナレッジのレビューUIが使う） */
  async function runExtractionForReview(caseId, category, { text = "", url = "" } = {}) {
    return callExtractionWebhook(caseId, category, { text, url });
  }
  /* レビュー後、編集済みの値で保存する（提案ナレッジのレビューUIが使う） */
  async function applyReviewedItems(caseId, category, items) {
    return applyCategoryToCalcSheet(caseId, category, items);
  }

  /* ---------- 提案サマリの参照（両アドイン共通） ----------
   * category を指定すればそのカテゴリだけ、省略すれば案件の全カテゴリを返す。 */
  async function getProposalSummaryForCase(caseId, onlyCategory) {
    if (!global.Office || !global.Excel) return [];
    let calcRows = [], solutions = [];
    await Excel.run(async ctx => {
      const calcSheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const r1 = calcSheet.getUsedRange(true);
      r1.load("values");
      const solSheet = ctx.workbook.worksheets.getItem(SOLUTION_SHEET);
      const r2 = solSheet.getUsedRange(true);
      r2.load("values");
      await ctx.sync();
      calcRows = r1.values.slice(1).filter(r => String(r[0]) === caseId);
      solutions = r2.values.slice(1);
    });

    const byCategory = {};
    calcRows.forEach(r => {
      const cat = r[1];
      byCategory[cat] = byCategory[cat] || { inputs: [], outputs: [] };
      (r[4] === "出力" ? byCategory[cat].outputs : byCategory[cat].inputs).push(r);
    });

    const cats = onlyCategory ? [onlyCategory] : Object.keys(byCategory);
    return cats.filter(c => byCategory[c]).map(cat => {
      const g = byCategory[cat];
      const matchingSolutions = solutions.filter(s => s[0] === cat);
      const savingRow = g.outputs.find(r => String(r[2]).endsWith("_saving")) || g.outputs[0];
      const conf = g.inputs.some(r => r[7] === "未確認") ? "未確認"
        : g.inputs.some(r => r[7] === "推定") ? "推定" : "確定";
      const selected = g.outputs.concat(g.inputs).some(r => String(r[8]).toUpperCase() === "TRUE");
      return {
        category: cat,
        solutionCount: matchingSolutions.length,
        saving: savingRow ? savingRow[5] : null,
        unit: savingRow ? savingRow[6] : "",
        confidence: conf,
        selected,
      };
    });
  }

  /* プロンプト出力など、生の明細行が必要な場面向け（提案ナレッジ側で使用）。
   * onlySelected=true なら「選択」列がTRUEのカテゴリの行のみを返す。 */
  async function getCalcRowsForCase(caseId, { onlySelected = false } = {}) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
    if (onlySelected) rows = rows.filter(r => String(r[8]).toUpperCase() === "TRUE");
    return rows.map(r => ({
      caseId: r[0], category: r[1], itemId: r[2], name: r[3], kind: r[4],
      value: r[5], unit: r[6], confidence: r[7], selected: String(r[8]).toUpperCase() === "TRUE",
    }));
  }

  async function toggleSelection(caseId, category, checked) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rng.values.forEach((r, i) => {
        if (i === 0) return;
        if (String(r[0]) === caseId && r[1] === category) sheet.getRange(`I${i + 1}`).values = [[checked ? "TRUE" : "FALSE"]];
      });
      await ctx.sync();
    });
  }

  function solRow(r) {
    return {
      category: r[0], name: r[1], cost: r[2], method: r[3],
      initialCost: Number(r[4]) || 0,
      rate: Number(r[5]) || 0,
      basis: r[6] || "",
      basisLevel: r[7] || "一般値",
    };
  }

  async function getSolutions() {
    if (!global.Office || !global.Excel) return SOLUTION_SEED.map(solRow);
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SOLUTION_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1);
    });
    return rows.filter(r => r[1]).map(solRow);
  }

  /* 指定カテゴリの解決案だけを返す（複数件）。低→中→高の順に並べる。 */
  async function getSolutionsForCategory(category) {
    const all = await getSolutions();
    const order = { "低": 0, "中": 1, "高": 2 };
    return all.filter(s => s.category === category).sort((a, b) => (order[a.cost] ?? 9) - (order[b.cost] ?? 9));
  }

  /* ---------- 抽出課題（AIが議事録から抽出した課題カテゴリと内容） ---------- */
  async function getIssues(caseId) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId && r[1]);
    });
    return rows.map(r => {
      const title = r[2] || "", summary = r[3] || "";
      const aiTitle = r[4] || "", aiSummary = r[5] || "";
      return {
        caseId: r[0], category: r[1], title, summary, aiTitle, aiSummary,
        // AIが書いた原文と現在の内容が違えば、営業が手を入れたとみなす
        edited: (title !== aiTitle) || (summary !== aiSummary),
        hearingIds: String(r[6] || "").split(",").filter(Boolean),
        extractedAt: r[7],
      };
    });
  }

  /* 課題を保存する。keepText=true なら課題タイトル・内容は既存のまま残し、
   * AI原文の列だけを更新する（営業の編集を守りつつ、次回の編集判定は
   * 最新のAI出力を基準にするため）。 */
  async function saveIssue(caseId, category, { title = "", summary = "", hearingIds = [], keepText = false } = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values, rowCount");
      await ctx.sync();
      const idx = used.values.slice(1).findIndex(r => String(r[0]) === caseId && r[1] === category);
      const rowNum = idx >= 0 ? idx + 2 : Math.max(used.rowCount, 1) + 1;
      const cur = idx >= 0 ? used.values[idx + 1] : null;
      const keepTitle = keepText && cur ? (cur[2] || "") : title;
      const keepSummary = keepText && cur ? (cur[3] || "") : summary;
      sheet.getRange(`A${rowNum}:H${rowNum}`).values = [[
        caseId, category, keepTitle, keepSummary, title, summary,
        hearingIds.filter(Boolean).join(","), nowStr(),
      ]];
      await ctx.sync();
    });
  }

  /* 営業が課題タイトル・内容を手で編集したときに呼ぶ（AI原文列は触らない）。 */
  async function updateIssueText(caseId, category, { title, summary } = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const idx = rng.values.slice(1).findIndex(r => String(r[0]) === caseId && r[1] === category);
      if (idx < 0) return;
      const rowNum = idx + 2;
      if (title !== undefined) sheet.getRange(`C${rowNum}`).values = [[title]];
      if (summary !== undefined) sheet.getRange(`D${rowNum}`).values = [[summary]];
      await ctx.sync();
    });
  }

  /* ---------- 提案決定（案件ごとの解決策の選択と、書き換えた導入費・改善率） ---------- */
  async function getDecisions(caseId) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(DECISION_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
    return rows.map(r => ({
      caseId: r[0], category: r[1], solutionName: r[2],
      cost: Number(r[3]) || 0, rate: Number(r[4]) || 0,
      basis: r[5] || "", basisLevel: r[6] || "一般値",
      included: String(r[7]).toUpperCase() === "TRUE",
    }));
  }

  /* 決定内容を保存する。同じ案件・同じ課題カテゴリの行があれば上書き、無ければ追加。 */
  async function saveDecision(caseId, category, d) {
    if (!global.Office || !global.Excel) return;
    const row = [caseId, category, d.solutionName, d.cost, d.rate, d.basis, d.basisLevel,
      d.included ? "TRUE" : "FALSE", nowStr()];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(DECISION_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values, rowCount");
      await ctx.sync();
      const idx = used.values.slice(1).findIndex(r => String(r[0]) === caseId && r[1] === category);
      const rowNum = idx >= 0 ? idx + 2 : Math.max(used.rowCount, 1) + 1;
      sheet.getRange(`A${rowNum}:I${rowNum}`).values = [row];
      await ctx.sync();
    });
  }

  /* ---------- 信頼度 ----------
   * 入力値（確定/推定/未確認）と、採用した解決策の根拠区分（実績/推定/一般値）を
   * 合わせて算出する。改善率は削減額への影響が大きいため、入力値の平均と
   * 同じ重みで扱う（単純平均ではなく 入力値平均:根拠 = 1:1）。 */
  function confidenceOf(inputRows, basisLevel) {
    const w = inputRows.length
      ? inputRows.reduce((a, r) => a + (CONF_WEIGHT[r.confidence] ?? 0), 0) / inputRows.length
      : null;
    const b = basisLevel ? (BASIS_WEIGHT[basisLevel] ?? 0.25) : null;
    if (w === null && b === null) return 0;
    if (w === null) return b;
    if (b === null) return w;
    return (w + b) / 2;
  }

  function confidenceLabel(r) {
    return r >= 0.9 ? "確定ベース" : r >= 0.5 ? "推定を含む" : "未確認が多い";
  }

  /* 個別の入力項目の値・信頼度を更新する（ウィザードのスライダー・信頼度チップ用）。 */
  async function updateCalcInput(caseId, itemId, { value, confidence } = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const idx = rng.values.slice(1).findIndex(r => String(r[0]) === caseId && r[2] === itemId);
      if (idx < 0) return;
      const rowNum = idx + 2;
      if (value !== undefined) sheet.getRange(`F${rowNum}`).values = [[value]];
      if (confidence !== undefined) sheet.getRange(`H${rowNum}`).values = [[confidence]];
      sheet.getRange(`J${rowNum}`).values = [[nowStr()]];
      await ctx.sync();
    });
  }

  async function getCustomerInfo(caseId) {
    if (!global.Office || !global.Excel) return null;
    let row = null;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CUST_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const code = caseId.split("-")[0];
      row = rng.values.slice(1).find(r => r[0] === code) || null;
    });
    if (!row) return null;
    return { code: row[0], name: row[1], contact: row[2] };
  }

  function nowStr() {
    const d = new Date();
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  global.RoiCore = {
    CONF_LEVELS, CONF_WEIGHT, BASIS_LEVELS, BASIS_WEIGHT,
    getConfig, setConfig,
    ensureAllSheets, getMasterItems, getCategories,
    listCaseIds, listCasesWithHearings,
    appendHearingLog, listHearingLogs,
    callExtractionWebhook, runExtractionForReview, applyReviewedItems,
    previewExtraction, commitExtraction,
    quickCreateProposal, autoExtractProposals, getSourceEntriesForCategory,
    getProposalSummaryForCase, getCalcRowsForCase, toggleSelection, updateCalcInput,
    getDecisions, saveDecision,
    getIssues, saveIssue, updateIssueText,
    confidenceOf, confidenceLabel,
    getCustomerInfo, getSolutions, getSolutionsForCategory,
  };

})(window);
