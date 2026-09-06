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
  // 利用範囲: "共通"（全案件で使う正式カテゴリ） / "一時"（作成した案件でのみ使う）
  // 作成案件ID: 利用範囲が"一時"のときに、どの案件用に作られたかを持つ。
  //   一時カテゴリは他案件のカテゴリ一覧には出さず、マスタを汚さない。
  //   複数案件で使われるようになったら、利用範囲を"共通"に変えるだけで昇格できる。
  const MASTER_COLUMNS = ["課題カテゴリ", "項目ID", "項目名", "区分", "単位", "デフォルト値", "数式", "信頼度初期値", "利用範囲", "作成案件ID"];

  const SCOPE_COMMON = "共通";
  const SCOPE_TEMP = "一時";

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
  // 種別: 経営課題（提案・ROI試算の対象） / 改修要望（既存案件の課題管理表へ） / 対象外
  //   議事録が既存案件の進捗会議だった場合、改修要望が多くなるのは正常。
  // 根拠事象: まとめる前の個別事象をJSON配列で保持（何をまとめたか辿れるようにするため）
  // 課題カテゴリ: 割り当て済みならカテゴリ名、未割り当てなら空。
  // 割当状態: 未設定 / 割当済 / カテゴリなしで提案 / 対象外
  // AI候補: AIが提示したカテゴリ候補をJSON文字列で保持（候補名・既存or新規・理由）
  const ISSUE_COLUMNS = ["案件ID", "課題ID", "種別", "課題カテゴリ", "課題タイトル", "課題内容",
    "AI原文タイトル", "AI原文内容", "割当状態", "AI候補", "根拠事象", "根拠議事録ID", "抽出日時"];

  const KIND_BUSINESS = "経営課題";
  const KIND_REQUEST = "改修要望";
  const KIND_NONE = "対象外";

  const ASSIGN_UNSET = "未設定";
  const ASSIGN_DONE = "割当済";
  const ASSIGN_NOCAT = "カテゴリなしで提案";
  const ASSIGN_SKIP = "対象外";

  // 案件ごとの「どの解決策に決めたか」を保持するシート。
  // 導入費・改善率・削減根拠はソリューションDBの値を初期値としてコピーし、
  // 営業側が案件ごとに書き換えられるようにする（マスタ側は変更しない）。
  const DECISION_SHEET = "提案決定";
  const DECISION_COLUMNS = ["案件ID", "課題カテゴリ", "解決策名", "導入費", "改善率", "削減根拠", "根拠区分", "提案に含める", "更新日時"];

  /* 案件IDの前後に空白が混じっていても同一とみなす（営業報告シート等の
   * コピー&ペースト由来の余分な空白で一致しなくなる事故を防ぐ）。 */
  function sameId(a, b) { return String(a ?? "").trim() === String(b ?? "").trim(); }

  const CONF_LEVELS = ["確定", "推定", "未確認"];
  const CONF_WEIGHT = { "確定": 1, "推定": 0.5, "未確認": 0 };

  const MASTER_SEED = [
    ["在庫管理", "stk_people", "棚卸人数", "入力", "人", 3, "", "未確認", "共通", ""],
    ["在庫管理", "stk_hours", "棚卸時間", "入力", "時間", 4, "", "未確認", "共通", ""],
    ["在庫管理", "stk_freq", "棚卸回数/年", "入力", "回", 12, "", "未確認", "共通", ""],
    ["在庫管理", "stk_wage", "時給", "入力", "円", 3000, "", "推定", "共通", ""],
    ["在庫管理", "stk_improve", "改善率", "入力", "%", 60, "", "推定", "共通", ""],
    ["在庫管理", "stk_hours_yr", "年間棚卸工数", "出力", "時間", "", "stk_people*stk_hours*stk_freq", "", "共通", ""],
    ["在庫管理", "stk_cost_yr", "年間棚卸コスト", "出力", "円", "", "stk_hours_yr*stk_wage", "", "共通", ""],
    ["在庫管理", "stk_saving", "削減額", "出力", "円", "", "stk_cost_yr*stk_improve/100", "", "共通", ""],
    ["ロット管理", "lot_hours", "追跡時間", "入力", "時間", 2, "", "未確認", "共通", ""],
    ["ロット管理", "lot_freq", "追跡回数/年", "入力", "回", 100, "", "未確認", "共通", ""],
    ["ロット管理", "lot_people", "担当人数", "入力", "人", 2, "", "未確認", "共通", ""],
    ["ロット管理", "lot_wage", "時給", "入力", "円", 3000, "", "推定", "共通", ""],
    ["ロット管理", "lot_improve", "改善率", "入力", "%", 60, "", "推定", "共通", ""],
    ["ロット管理", "lot_hours_yr", "年間追跡工数", "出力", "時間", "", "lot_hours*lot_freq*lot_people", "", "共通", ""],
    ["ロット管理", "lot_saving", "削減額", "出力", "円", "", "lot_hours_yr*lot_wage*lot_improve/100", "", "共通", ""],
    ["AI議事録", "min_meetings", "会議回数/月", "入力", "回", 8, "", "未確認", "共通", ""],
    ["AI議事録", "min_people", "参加人数", "入力", "人", 3, "", "未確認", "共通", ""],
    ["AI議事録", "min_hours", "議事録作成時間", "入力", "時間", 1, "", "未確認", "共通", ""],
    ["AI議事録", "min_wage", "時給", "入力", "円", 3000, "", "推定", "共通", ""],
    ["AI議事録", "min_hours_yr", "年間工数", "出力", "時間", "", "min_meetings*12*min_people*min_hours", "", "共通", ""],
    ["AI議事録", "min_saving", "削減額", "出力", "円", "", "min_hours_yr*min_wage", "", "共通", ""],
    // 製造管理
    ["製造管理", "mfg_revenue", "生産額", "入力", "円", 50000000, "", "未確認", "共通", ""],
    ["製造管理", "mfg_defect_rate", "不良率", "入力", "%", 3, "", "推定", "共通", ""],
    ["製造管理", "mfg_defect_rate_after", "改善後不良率", "入力", "%", 1, "", "推定", "共通", ""],
    ["製造管理", "mfg_loss", "不良損失", "出力", "円", "", "mfg_revenue*mfg_defect_rate/100", "", "共通", ""],
    ["製造管理", "mfg_loss_after", "改善後不良損失", "出力", "円", "", "mfg_revenue*mfg_defect_rate_after/100", "", "共通", ""],
    ["製造管理", "mfg_saving", "削減額", "出力", "円", "", "mfg_loss-mfg_loss_after", "", "共通", ""],
    // OCR受注入力
    ["OCR受注入力", "ocr_count", "注文書件数/月", "入力", "件", 200, "", "未確認", "共通", ""],
    ["OCR受注入力", "ocr_input_min", "入力時間", "入力", "分", 5, "", "未確認", "共通", ""],
    ["OCR受注入力", "ocr_fix_min", "修正時間", "入力", "分", 2, "", "未確認", "共通", ""],
    ["OCR受注入力", "ocr_wage", "時給", "入力", "円", 3000, "", "推定", "共通", ""],
    ["OCR受注入力", "ocr_improve", "改善率", "入力", "%", 70, "", "推定", "共通", ""],
    ["OCR受注入力", "ocr_hours_yr", "年間入力時間", "出力", "時間", "", "ocr_count*12*(ocr_input_min+ocr_fix_min)/60", "", "共通", ""],
    ["OCR受注入力", "ocr_saving", "削減額", "出力", "円", "", "ocr_hours_yr*ocr_wage*ocr_improve/100", "", "共通", ""],
    // Delphi移行
    ["Delphi移行", "delphi_staff", "保守担当人数", "入力", "人", 2, "", "未確認", "共通", ""],
    ["Delphi移行", "delphi_inquiries", "問い合わせ件数/年", "入力", "件", 300, "", "未確認", "共通", ""],
    ["Delphi移行", "delphi_avg_hours", "平均対応時間", "入力", "時間", 1, "", "推定", "共通", ""],
    ["Delphi移行", "delphi_wage", "時給", "入力", "円", 3000, "", "推定", "共通", ""],
    ["Delphi移行", "delphi_improve", "改善率", "入力", "%", 50, "", "推定", "共通", ""],
    ["Delphi移行", "delphi_hours_yr", "年間保守工数", "出力", "時間", "", "delphi_inquiries*delphi_avg_hours", "", "共通", ""],
    ["Delphi移行", "delphi_cost_yr", "年間保守コスト", "出力", "円", "", "delphi_hours_yr*delphi_wage", "", "共通", ""],
    ["Delphi移行", "delphi_saving", "削減額", "出力", "円", "", "delphi_cost_yr*delphi_improve/100", "", "共通", ""],
    // 属人化・情報共有・問合せ対応: ROI試算は行わず、定性的な解決案候補のみ持つ
    // カテゴリ一覧に出てくるよう最低限の入力項目だけ用意している
    ["属人化", "attrib_people", "該当ベテラン人数", "入力", "人", 1, "", "未確認", "共通", ""],
    ["情報共有", "info_tools", "使用ツール数", "入力", "個", 3, "", "未確認", "共通", ""],
    ["問合せ対応", "inquiry_count", "月間問合せ件数", "入力", "件", 20, "", "未確認", "共通", ""],
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
   * APIキー本体もここに保存する（社内利用限定の前提。GASは廃止）。
   * 営業報告・提案ナレッジの両アドインは同一オリジンなので localStorage を共有する。 */
  function getConfig() {
    try { return JSON.parse(localStorage.getItem("roiAddinConfig") || "{}"); }
    catch (e) { return {}; }
  }
  function setConfig(cfg) { localStorage.setItem("roiAddinConfig", JSON.stringify(cfg)); }

  /* ---------- AI直接呼び出し（GAS廃止） ----------
   * 社内利用限定という前提で、ブラウザから直接AI APIを呼ぶ。
   * 設定（⚙）に登録した「ベースURL・APIキー・モデル」をそのまま使う。
   * 注意: この構成はAPIキーがブラウザ側に露出する（開発者ツールで見える、
   * roi-core.js自体もGitHub Pagesで公開されている）。社外に公開しないこと。 */

  /* Claudeがコードフェンス付き（```json ... ```）で返すことがあるため除去してからパースする。 */
  function safeParseJson(text, fallback) {
    try {
      const cleaned = String(text || "").replace(/```json/gi, "").replace(/```/g, "").trim();
      if (!cleaned) return fallback;
      return JSON.parse(cleaned);
    } catch (e) { return fallback; }
  }

  /* プロバイダ差異を吸収してAIを呼び、テキスト応答だけを返す。
   * provider: "anthropic"（既定） または "openai"（OpenAI互換。Gemini等もこちら）。 */
  async function callAiModel(prompt, maxTokens) {
    const cfg = getConfig();
    if (!cfg.aiApiKey) throw new Error("設定（⚙）でAPIキーを登録してください");
    const provider = cfg.aiProvider || "anthropic";

    if (provider === "openai") {
      if (!cfg.aiBaseUrl) throw new Error("設定（⚙）でベースURLを登録してください（OpenAI互換の場合は必須）");
      const url = cfg.aiBaseUrl.replace(/\/+$/, "") + "/chat/completions";
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": "Bearer " + cfg.aiApiKey },
        body: JSON.stringify({
          model: cfg.aiModel || "gpt-4o-mini",
          max_tokens: maxTokens,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      const raw = await res.text();
      const body = safeParseJson(raw, null);
      if (!body) throw new Error(`AI APIの応答がJSONではありません（HTTP ${res.status}）: ` + raw.slice(0, 300));
      if (body.error) throw new Error("AI APIエラー: " + JSON.stringify(body.error));
      const text = body.choices && body.choices[0] && body.choices[0].message && body.choices[0].message.content;
      return text || "";
    }

    // 既定: Anthropic Messages API
    const url = cfg.aiBaseUrl || "https://api.anthropic.com/v1/messages";
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": cfg.aiApiKey,
        "anthropic-version": "2023-06-01",
        // ブラウザから直接呼ぶ場合に必要（Anthropicはブラウザからの直接アクセスを
        // 既定でブロックしており、この明示ヘッダーで許可する）。
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: cfg.aiModel || "claude-sonnet-4-6",
        max_tokens: maxTokens,
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const raw = await res.text();
    const body = safeParseJson(raw, null);
    if (!body) throw new Error(`AI APIの応答がJSONではありません（HTTP ${res.status}）: ` + raw.slice(0, 300));
    if (body.error) throw new Error("AI APIエラー: " + JSON.stringify(body.error));
    const textBlock = (body.content || []).find(c => c.type === "text");
    return (textBlock && textBlock.text) || "";
  }

  /* 参照URLの中身をブラウザから直接取得する。GAS（サーバ側）が無くなったため、
   * fetch先がCORSを許可していないと失敗する（Notion等の多くのサービスは許可していない）。
   * 失敗した場合はそのURLをスキップし、consoleに警告を出す。 */
  async function fetchTextFromUrlClient(url) {
    try {
      const res = await fetch(url);
      const html = await res.text();
      return html.replace(/<script[\s\S]*?<\/script>/gi, "")
        .replace(/<style[\s\S]*?<\/style>/gi, "")
        .replace(/<[^>]+>/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 8000);
    } catch (e) {
      console.warn("[RoiCore] 参照URLの取得に失敗しました（CORSの可能性）:", url, e);
      return "";
    }
  }

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
    return {
      category: r[0], itemId: r[1], name: r[2], kind: r[3], unit: r[4],
      defaultVal: r[5], formula: r[6], confDefault: r[7],
      // 既存シート（列が無い旧版）は空になるので「共通」扱いにする
      scope: r[8] || SCOPE_COMMON,
      ownerCaseId: r[9] || "",
    };
  }

  /* この案件で使えるマスタ項目。共通カテゴリ＋この案件用の一時カテゴリ。 */
  function getMasterItemsFor(caseId) {
    return getMasterItems().filter(m =>
      m.scope !== SCOPE_TEMP || sameId(m.ownerCaseId, caseId));
  }

  /* この案件で使えるカテゴリ名の一覧。 */
  function getCategoriesFor(caseId) {
    return Array.from(new Set(getMasterItemsFor(caseId).map(m => m.category)));
  }

  /* 一時カテゴリの一覧（マスタ化候補の確認用）。
   * 利用状況（何案件で使われているか）も返す。 */
  async function listTempCategories() {
    const temps = {};
    getMasterItems().filter(m => m.scope === SCOPE_TEMP).forEach(m => {
      temps[m.category] = temps[m.category] || { category: m.category, ownerCaseId: m.ownerCaseId, itemCount: 0 };
      temps[m.category].itemCount++;
    });
    if (!global.Office || !global.Excel) return Object.values(temps);
    // ROI試算シートで実際に使われている案件数を数える
    let calcRows = [];
    await Excel.run(async ctx => {
      const rng = ctx.workbook.worksheets.getItem(CALC_SHEET).getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      calcRows = rng.values.slice(1);
    });
    Object.values(temps).forEach(t => {
      const cases = new Set(calcRows.filter(r => r[1] === t.category).map(r => String(r[0]).trim()));
      t.usedCaseCount = cases.size;
    });
    return Object.values(temps);
  }

  /* 一時カテゴリを共通カテゴリへ昇格させる（利用範囲を"共通"に書き換える）。 */
  async function promoteTempCategory(category) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(MASTER_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rng.values.forEach((r, i) => {
        if (i === 0 || r[0] !== category) return;
        sheet.getRange(`I${i + 1}`).values = [[SCOPE_COMMON]];
        sheet.getRange(`J${i + 1}`).values = [[""]];
      });
      await ctx.sync();
    });
    await ensureAllSheets(); // マスタのキャッシュを更新
  }

  /* AIとの壁打ちで決めたカテゴリ定義をROIマスタに追加する。
   * scope に SCOPE_TEMP を渡すと、その案件専用の一時カテゴリになる。
   * def = { category, inputs:[{itemId,name,unit,defaultVal,confDefault}],
   *         outputs:[{itemId,name,unit,formula}] } */
  async function addCategoryToMaster(def, { scope = SCOPE_TEMP, caseId = "" } = {}) {
    if (!global.Office || !global.Excel) return;
    const rows = [];
    (def.inputs || []).forEach(i => rows.push([
      def.category, i.itemId, i.name, "入力", i.unit || "",
      i.defaultVal ?? "", "", i.confDefault || "未確認",
      scope, scope === SCOPE_TEMP ? caseId : "",
    ]));
    (def.outputs || []).forEach(o => rows.push([
      def.category, o.itemId, o.name, "出力", o.unit || "",
      "", o.formula || "", "",
      scope, scope === SCOPE_TEMP ? caseId : "",
    ]));
    if (!rows.length) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(MASTER_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const start = Math.max(used.rowCount, 1) + 1;
      sheet.getRange(`A${start}:J${start + rows.length - 1}`).values = rows;
      await ctx.sync();
    });
    await ensureAllSheets();
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
        const id = r[0] ? String(r[0]).trim() : "";
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
      rows = rng.values.slice(1).filter(r => sameId(r[0], caseId));
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
      const id = String(r[0]).trim();
      map[id] = map[id] || { caseId: id, hearingCount: 0, issueCount: 0 };
      map[id].hearingCount++;
    });
    issues.forEach(r => {
      const id = String(r[0]).trim();
      map[id] = map[id] || { caseId: id, hearingCount: 0, issueCount: 0 };
      map[id].issueCount++;
    });
    return Object.values(map);
  }

  function genId(prefix) {
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  }

  /* ---------- AI抽出（単一カテゴリ） ----------
   * url が指定されている場合はブラウザから直接ページ内容を取得する。
   * レスポンス: { items: [{ itemId, value, confidence }] } */
  async function callExtractionWebhook(caseId, category, { text = "", url = "" } = {}) {
    const items = getMasterItems().filter(m => m.category === category && m.kind === "入力");
    let sourceText = capText(text || "");
    if (!sourceText && url) sourceText = await fetchTextFromUrlClient(url);
    if (!sourceText) throw new Error("text and url are both empty");

    const itemList = items.map(i => `- ${i.itemId} (${i.name}, 単位:${i.unit})`).join("\n");
    const prompt =
`以下の議事録テキストから、指定した項目IDに対応する数値を抽出してください。
読み取れない項目は value を null にし、confidence は "未確認" としてください。
数値が明言されておらず推測が入る場合は confidence を "推定" にしてください。
明確に数値が述べられている場合のみ confidence を "確定" にしてください。

項目一覧:
${itemList}

議事録テキスト:
"""
${sourceText}
"""

出力は次のJSON形式のみとしてください（説明文は不要）:
{"items":[{"itemId":"...","value":数値またはnull,"confidence":"確定|推定|未確認"}]}`;

    const raw = await callAiModel(prompt, 1500);
    console.log("[RoiCore] extraction raw response:", raw);
    const data = safeParseJson(raw, { items: [] });
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
  // 以前は404対策として6,000文字で切り詰めていたが、実際の原因は
  // 以前GAS経由で運用していた際、デプロイ・バージョン管理の問題だと判明したため、
  // 冒頭の雑談・ノイズで実質的な内容が切り捨てられないよう上限を引き上げる。
  const HEARING_TEXT_LIMIT = 30000;
  function capText(s) {
    if (!s) return s;
    return s.length > HEARING_TEXT_LIMIT ? s.slice(0, HEARING_TEXT_LIMIT) + "\n…（以降省略）" : s;
  }

  async function previewExtraction(caseId, memoText = "") {
    const logs = await listHearingLogs(caseId);
    // テキストがある議事録はそのまま、URLのみの議事録はブラウザから直接取得を試みる
    // （CORSを許可していないサービスだと失敗し、そのURLはスキップされる）。
    const hearings = [];
    for (const l of logs) {
      let text = capText(l.text || "");
      if (!text && l.url) text = await fetchTextFromUrlClient(l.url);
      if (text || l.url) hearings.push({ title: l.title, text, url: l.url || "" });
    }
    console.log("[RoiCore] previewExtraction: logs=" + logs.length
      + " hearings(after filter)=" + hearings.length
      + " memoText.length=" + (memoText || "").length);
    if (!hearings.length && !memoText) {
      console.log("[RoiCore] previewExtraction: 議事録が空と判定したためAI呼び出しをせず終了します");
      return { hearingIds: [], results: [] };
    }
    const hearingIds = logs.map(l => l.hearingId).filter(Boolean);

    // この案件で使えるカテゴリ（共通＋この案件の一時カテゴリ）を候補判定の材料として渡す
    const categoryDefs = getCategoriesFor(caseId).map(cat => ({
      category: cat,
      items: getMasterItemsFor(caseId).filter(m => m.category === cat && m.kind === "入力")
        .map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
    }));
    const categoryBlock = categoryDefs.map(c =>
      `### ${c.category}\n${c.items.map(i => `- ${i.itemId} (${i.name}, 単位:${i.unit})`).join("\n")}`
    ).join("\n\n");

    const combinedText = [memoText, ...hearings.map(h => `【${h.title || "議事録"}】\n${h.text}`)]
      .filter(Boolean).join("\n\n");

    const prompt =
`あなたは中小企業向けの業務システム提案を行う営業担当です。
以下の議事録・メモを読み、顧客が抱えている課題を抽出してください。

## 粒度の基準（最重要）

個別の事象を1件ずつ並べるのではなく、**提案の単位**にまとめてください。
目安は「その課題を解決する提案書を1本書けるか」です。

- 同じ原因・同じ業務から生じている事象は、1つの課題にまとめること
  例)「フリー予約がない」「割り振りが属人化」「予約を受けきれない」
    → まとめて「予約機会の取りこぼし」1件とする
- 金額換算したときに年間数十万円以上のインパクトがある規模を目安にする
  操作性の細かい不満や、打合せ中の一時的な出来事は課題として挙げない
- 出力は多くても5件まで。重要度の高い順に並べること
  細かく分けたくなっても、まとめられないか必ず一度検討すること

## 種別の判定

各課題に kind を付けてください。

- "経営課題": 顧客の事業運営上の困りごと。システム提案・ROI試算の対象になるもの
- "改修要望": すでに導入済み・開発中のシステムに対する不具合報告や改善要望。
  提案の対象ではなく、既存案件の課題管理表に登録すべきもの
- "対象外": 打合せ中の一時的な事象など、課題として扱う必要がないもの

議事録が既存案件の進捗確認会議だった場合、"改修要望" が多くなるのは正常です。
"経営課題" を無理に作り出さないでください。

## 抽出の範囲

顧客が困りごととして語っている内容は、下の「既存カテゴリ一覧」に
当てはまるかどうかに関わらず抽出してください。
既存カテゴリに無いという理由で課題を捨てないでください。

## 各課題の出力項目

- kind: 上記の3種別のいずれか
- title: その顧客固有の課題を一行で。カテゴリ名をそのまま書かず、具体的な状況を反映する
- summary: 2〜3文で。現状の進め方、何に困っているか、その影響を含める。
  議事録に書かれていないことは推測で補わないこと
- sources: この課題の根拠になった議事録中の具体的な事象を、短い文で1〜5件の配列。
  まとめた場合、元が何だったか分かるようにするため
- matchedCategory: kind が "経営課題" で、既存カテゴリのどれかに**明確に**当てはまる場合のみ、
  そのカテゴリ名。少しでも迷う場合、および "改修要望"・"対象外" の場合は空文字 ""
- items: matchedCategory を設定した場合のみ、そのカテゴリの項目一覧に対応する数値。
  読み取れない項目は value を null、confidence を "未確認" とする。
  数値が明言されていれば "確定"、文脈から推測した場合は "推定"。
  matchedCategory が空なら items は空配列にすること
- candidates: kind が "経営課題" の場合のみ、担当者がカテゴリを選ぶための候補を1〜3件。
  各候補は { name, isNew, reason } の形。既存カテゴリなら一覧の名前をそのまま使い isNew は false。
  新規候補の name は、その顧客固有の言葉ではなく他業種でも通じる一般的な名前にすること。
  "改修要望"・"対象外" の場合は空配列にすること

既存カテゴリ一覧（この案件で使えるもの）:
${categoryBlock || "（まだカテゴリが登録されていません）"}

議事録・メモ:
"""
${combinedText}
"""

出力は次のJSON形式のみとしてください（説明文やコードフェンスは不要）:
{"issues":[{"kind":"経営課題","title":"...","summary":"...","sources":["..."],"matchedCategory":"","items":[],"candidates":[{"name":"...","isNew":false,"reason":"..."}]}]}`;

    const raw = await callAiModel(prompt, 8000);
    console.log("[RoiCore] auto-extraction raw response:", raw);
    let data = safeParseJson(raw, null);
    let warning = null;
    if (!data) {
      console.warn("[RoiCore] JSON解析に失敗しました（max_tokens超過の可能性）。出力末尾200文字:", raw.slice(-200));
      data = { issues: [] };
      warning = "AIの応答が途中で切れた可能性があります（max_tokens超過）。もう一度お試しください。";
    }

    const existing = await getIssues(caseId);
    const results = (data.issues || []).filter(r => r && r.title).map(r => {
      // 既存課題との突き合わせは、AIが返した課題IDがあればそれで、
      // 無ければAI原文タイトルの一致で行う（再抽出時に重複を作らないため）
      const cur = existing.find(x => (r.issueId && x.issueId === r.issueId) || x.aiTitle === r.title);
      return {
        issueId: cur ? cur.issueId : genId("I"),
        kind: r.kind || KIND_BUSINESS,
        newTitle: r.title,
        newSummary: r.summary || "",
        sources: r.sources || [],
        candidates: r.candidates || [],
        // AIが「既存カテゴリにそのまま当てはまる」と判断した場合のみ設定される
        matchedCategory: r.matchedCategory || "",
        items: r.items || [],
        current: cur || null,
        status: !cur ? "新規" : (cur.edited ? "編集済み" : "未編集"),
        keepText: cur ? cur.edited : false,
      };
    });
    return { hearingIds, results, warning };
  }

  async function commitExtraction(caseId, hearingIds, results) {
    for (const r of results) {
      const assigned = !!r.matchedCategory;
      // 改修要望・対象外はROI試算しないので、割当状態は「対象外」扱いにする
      const isBusiness = r.kind === KIND_BUSINESS;
      await saveIssue(caseId, r.issueId, {
        kind: r.kind,
        category: assigned ? r.matchedCategory : (r.current ? r.current.category : ""),
        title: r.newTitle, summary: r.newSummary,
        sources: r.sources,
        assignStatus: !isBusiness ? ASSIGN_SKIP
          : (assigned ? ASSIGN_DONE : (r.current ? r.current.assignStatus : ASSIGN_UNSET)),
        candidates: r.candidates,
        hearingIds, keepText: !!r.keepText,
      });
      if (isBusiness && assigned && r.items && r.items.length) {
        await applyCategoryToCalcSheet(caseId, r.matchedCategory, r.items, hearingIds);
      }
    }
    return results.length;
  }

  /* ---------- AIと壁打ちして新カテゴリを設計する ----------
   * 会話履歴を渡すと、AIの返答とカテゴリ定義案（名前・入力項目・計算式）を返す。
   * 返ってきた定義は addCategoryToMaster() で登録する。 */
  async function proposeCategoryDefinition(caseId, issue, messages = []) {
    const history = (messages || [])
      .map(m => (m.role === "user" ? "担当者: " : "AI: ") + m.content).join("\n");
    const existingCategories = getCategoriesFor(caseId);

    const prompt =
`あなたは中小企業向けの業務システム提案で使うROI試算マスタの設計を手伝います。
担当者と相談しながら、この課題を試算するための課題カテゴリを設計してください。

対象の課題:
タイトル: ${issue.title || ""}
内容: ${issue.summary || ""}

既にあるカテゴリ（重複させないこと）:
${(existingCategories || []).join(" / ") || "（なし）"}

これまでのやり取り:
${history || "（まだありません）"}

設計の方針:
- カテゴリ名は、この顧客固有の言葉ではなく、他業種でも通じる一般的な名前にする
  （例:「美容室の予約電話対応」ではなく「予約・スケジュール調整」）
- 既存カテゴリで足りるなら、無理に新規を作らずそう伝える
- 入力項目は、顧客にヒアリングすれば答えられる具体的な数値にする
  （件数・時間・人数・単価など。抽象的な指標は避ける）
- 出力項目には必ず削減額を含め、項目IDは "_saving" で終わらせる
- 数式は入力項目の項目IDだけを使った式にする（例: "xxx_count*12*xxx_min/60"）
- 項目IDは英小文字とアンダースコアのみ。カテゴリごとに共通の接頭辞をつける

reply には担当者への返答を書いてください。設計の意図や、迷っている点への質問を
1〜3文で簡潔に。専門用語を並べず、平易な日本語で書くこと。

出力は次のJSON形式のみとしてください（説明文やコードフェンスは不要）:
{"reply":"...","definition":{"category":"...","inputs":[{"itemId":"...","name":"...","unit":"...","defaultVal":0,"confDefault":"未確認"}],"outputs":[{"itemId":"..._saving","name":"削減額","unit":"円","formula":"..."}]}}`;

    const raw = await callAiModel(prompt, 2500);
    console.log("[RoiCore] category-design raw response:", raw);
    return safeParseJson(raw, { reply: "", definition: null });
  }

  /* 確認なしで一括反映する簡易版（営業報告アドインの「作成」アイコン用）。
   * 編集済みの課題文は自動的に残す。 */
  async function autoExtractProposals(caseId, memoText = "") {
    const { hearingIds, results } = await previewExtraction(caseId, memoText);
    await commitExtraction(caseId, hearingIds, results);
    return results.map(r => ({
      issueId: r.issueId, category: r.matchedCategory, title: r.newTitle,
      summary: r.newSummary, itemCount: (r.items || []).length, status: r.status,
      unassigned: !r.matchedCategory,
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
      const row = rng.values.slice(1).find(r => sameId(r[0], caseId) && r[1] === category && r[10]);
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
      calcRows = r1.values.slice(1).filter(r => sameId(r[0], caseId));
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
      rows = rng.values.slice(1).filter(r => sameId(r[0], caseId));
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
        if (sameId(r[0], caseId) && r[1] === category) sheet.getRange(`I${i + 1}`).values = [[checked ? "TRUE" : "FALSE"]];
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
      rows = rng.values.slice(1).filter(r => sameId(r[0], caseId) && r[1]);
    });
    return rows.map(r => {
      const title = r[4] || "", summary = r[5] || "";
      const aiTitle = r[6] || "", aiSummary = r[7] || "";
      let candidates = [], sources = [];
      try { candidates = JSON.parse(r[9] || "[]"); } catch (e) { candidates = []; }
      try { sources = JSON.parse(r[10] || "[]"); } catch (e) { sources = []; }
      return {
        caseId: r[0], issueId: r[1],
        kind: r[2] || KIND_BUSINESS,
        category: r[3] || "", title, summary,
        aiTitle, aiSummary,
        edited: (title !== aiTitle) || (summary !== aiSummary),
        assignStatus: r[8] || ASSIGN_UNSET,
        candidates, sources,
        hearingIds: String(r[11] || "").split(",").filter(Boolean),
        extractedAt: r[12],
      };
    });
  }

  /* 課題を保存する。課題IDで既存行を探し、あれば上書き、無ければ追加。
   * keepText=true なら課題タイトル・内容は既存のまま残し、AI原文だけ更新する。 */
  async function saveIssue(caseId, issueId, d = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values, rowCount");
      await ctx.sync();
      const idx = used.values.slice(1).findIndex(r => sameId(r[0], caseId) && r[1] === issueId);
      const rowNum = idx >= 0 ? idx + 2 : Math.max(used.rowCount, 1) + 1;
      const cur = idx >= 0 ? used.values[idx + 1] : null;
      const keepTitle = d.keepText && cur ? (cur[4] || "") : (d.title || "");
      const keepSummary = d.keepText && cur ? (cur[5] || "") : (d.summary || "");
      sheet.getRange(`A${rowNum}:M${rowNum}`).values = [[
        caseId, issueId,
        d.kind || (cur ? cur[2] : KIND_BUSINESS) || KIND_BUSINESS,
        d.category !== undefined ? d.category : (cur ? cur[3] : ""),
        keepTitle, keepSummary,
        d.title || "", d.summary || "",
        d.assignStatus || (cur ? cur[8] : ASSIGN_UNSET) || ASSIGN_UNSET,
        JSON.stringify(d.candidates || (cur ? (() => { try { return JSON.parse(cur[9] || "[]"); } catch (e) { return []; } })() : [])),
        JSON.stringify(d.sources || (cur ? (() => { try { return JSON.parse(cur[10] || "[]"); } catch (e) { return []; } })() : [])),
        (d.hearingIds || []).filter(Boolean).join(","),
        nowStr(),
      ]];
      await ctx.sync();
    });
  }

  /* 担当者がカテゴリを割り当てた/扱いを決めたときに呼ぶ。 */
  async function assignIssueCategory(caseId, issueId, { category = "", assignStatus } = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const idx = rng.values.slice(1).findIndex(r => sameId(r[0], caseId) && r[1] === issueId);
      if (idx < 0) return;
      const rowNum = idx + 2;
      sheet.getRange(`D${rowNum}`).values = [[category]];
      if (assignStatus) sheet.getRange(`I${rowNum}`).values = [[assignStatus]];
      await ctx.sync();
    });
  }

  /* 営業が課題タイトル・内容を手で編集したときに呼ぶ（AI原文列は触らない）。 */
  async function updateIssueText(caseId, issueId, { title, summary } = {}) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(ISSUE_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const idx = rng.values.slice(1).findIndex(r => sameId(r[0], caseId) && r[1] === issueId);
      if (idx < 0) return;
      const rowNum = idx + 2;
      if (title !== undefined) sheet.getRange(`E${rowNum}`).values = [[title]];
      if (summary !== undefined) sheet.getRange(`F${rowNum}`).values = [[summary]];
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
      rows = rng.values.slice(1).filter(r => sameId(r[0], caseId));
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
      const idx = used.values.slice(1).findIndex(r => sameId(r[0], caseId) && r[1] === category);
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
      const idx = rng.values.slice(1).findIndex(r => sameId(r[0], caseId) && r[2] === itemId);
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
      const code = String(caseId).trim().split("-")[0];
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
    getIssues, saveIssue, updateIssueText, assignIssueCategory,
    proposeCategoryDefinition, addCategoryToMaster,
    listTempCategories, promoteTempCategory,
    getMasterItemsFor, getCategoriesFor,
    ASSIGN_UNSET, ASSIGN_DONE, ASSIGN_NOCAT, ASSIGN_SKIP,
    KIND_BUSINESS, KIND_REQUEST, KIND_NONE,
    SCOPE_COMMON, SCOPE_TEMP,
    confidenceOf, confidenceLabel,
    getCustomerInfo, getSolutions, getSolutionsForCategory,
  };

})(window);
