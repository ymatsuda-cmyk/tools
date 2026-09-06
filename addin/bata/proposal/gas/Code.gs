/**
 * 提案ナレッジ アドイン用 AI連携Webhook（サンプル）
 * ------------------------------------------------------------
 * デプロイ: 「Webアプリとして公開」→ アクセスできるユーザー：全員
 * 発行されたURLを、アドインの設定（⚙）の「AI連携エンドポイント」に登録する。
 *
 * APIキーは PropertiesService に保存する（クライアント側には一切渡さない）。
 * スクリプトエディタ → プロジェクトの設定 → スクリプト プロパティ で
 * AI_API_KEY を設定しておくこと。
 *
 * リクエスト（roi-core.js の callExtractionWebhook から送られる）:
 * {
 *   token: "簡易トークン（任意、照合用）",
 *   caseId: "KM-01",
 *   category: "在庫管理",
 *   text: "議事録の文字起こし・メモ（空でもよい）",
 *   url: "議事録ビューア等へのリンク（空でもよい。text優先、textが空ならurlを取得する）",
 *   items: [{ itemId: "stk_people", name: "棚卸人数", unit: "人" }, ...]
 * }
 * text と url のどちらか一方があればよい。両方空はエラーとする。
 *
 * レスポンス:
 * { items: [{ itemId: "stk_people", value: 5, confidence: "確定" }, ...] }
 *
 * mode: "auto" を指定すると複数カテゴリ一括判定モードになる（handleAutoMode参照）。
 * 営業報告アドインの「作成」アイコンはこちらを使う。
 *
 * 対応プロバイダ: Anthropic（既定）／ OpenAI互換（Gemini等。AI_PROVIDER=openai）。
 * getAiSettings() と callAiModel() を参照。
 */

const SHARED_TOKEN = "";              // 空なら照合をスキップ（開発時のみ推奨）

/* AI_API_KEY・AI_API_URL・AI_MODEL・AI_PROVIDER はすべてスクリプトプロパティに設定する。
 * スクリプトエディタ → プロジェクトの設定 → スクリプト プロパティ で以下を登録:
 *   AI_API_KEY  … 必須。プロバイダのAPIキー
 *   AI_PROVIDER … "anthropic"（既定） または "openai"（OpenAI互換。Gemini等もこちら）
 *   AI_API_URL  … anthropicなら未設定でデフォルト値を使う。
 *                 openaiの場合は必須（例: Geminiなら
 *                 https://generativelanguage.googleapis.com/v1beta/openai ）
 *                 ※ "/chat/completions" は自動で付加するので含めなくてよい
 *   AI_MODEL    … 例: claude-sonnet-4-6 / gemini-2.5-flash 等
 * コードを直接編集しなくても、モデル・プロバイダ差し替えができるようにするため。 */
function getAiSettings() {
  const props = PropertiesService.getScriptProperties();
  const provider = (props.getProperty("AI_PROVIDER") || "anthropic").toLowerCase();
  return {
    provider,
    apiKey: props.getProperty("AI_API_KEY"),
    apiUrl: props.getProperty("AI_API_URL") || (provider === "anthropic" ? "https://api.anthropic.com/v1/messages" : ""),
    model: props.getProperty("AI_MODEL") || "claude-sonnet-4-6",
  };
}

/* プロバイダ差異を吸収してAIを呼び、テキスト応答だけを返す共通関数。
 * 戻り値: { text } または { error } */
function callAiModel(prompt, maxTokens, settings) {
  if (!settings.apiKey) return { error: "AI_API_KEY not configured" };

  if (settings.provider === "openai") {
    if (!settings.apiUrl) return { error: "AI_API_URL is required when AI_PROVIDER=openai" };
    const url = settings.apiUrl.replace(/\/+$/, "") + "/chat/completions";
    const payload = {
      model: settings.model,
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }],
    };
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + settings.apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const body = safeParseJson(res.getContentText(), null);
    if (!body) return { error: `AI API response was not valid JSON (HTTP ${res.getResponseCode()} @ ${url}): ` + res.getContentText().slice(0, 300) };
    if (body.error) return { error: "AI API error: " + JSON.stringify(body.error) };
    const text = body.choices && body.choices[0] && body.choices[0].message && body.choices[0].message.content;
    return { text: text || "" };
  }

  // デフォルト: Anthropic Messages API
  const res = UrlFetchApp.fetch(settings.apiUrl, {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": settings.apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({ model: settings.model, max_tokens: maxTokens, messages: [{ role: "user", content: prompt }] }),
    muteHttpExceptions: true,
  });
  const body = safeParseJson(res.getContentText(), null);
  if (!body) return { error: `AI API response was not valid JSON (HTTP ${res.getResponseCode()} @ ${settings.apiUrl}): ` + res.getContentText().slice(0, 300) };
  if (body.error) return { error: "AI API error: " + JSON.stringify(body.error) };
  const textBlock = (body.content || []).find(c => c.type === "text");
  return { text: (textBlock && textBlock.text) || "" };
}

/* ------------------------------------------------------------
 * 手動テスト用。スクリプトエディタでこの関数を選んで「実行」すると、
 * ・外部リクエスト（api.anthropic.com）の権限承認ダイアログが出る（初回のみ）
 * ・実行数ログに、AI APIへのリクエスト結果がそのまま出る
 * ので、Webアプリ経由ではなくこの関数単体でAI連携の疎通確認ができる。
 * 実行後は「実行数」メニューでログを確認すること。
 * ------------------------------------------------------------ */
function testRun() {
  const settings = getAiSettings();
  Logger.log("provider: [" + settings.provider + "]");
  Logger.log("apiUrl: [" + settings.apiUrl + "]");
  Logger.log("model: [" + settings.model + "]");
  Logger.log("apiKey set: " + (settings.apiKey ? ("yes, length=" + settings.apiKey.length) : "NO"));

  const result = handleSingleCategoryMode({
    caseId: "TEST-01",
    category: "在庫管理",
    text: "議事1 棚卸は5人で行っていて、月に1回、1回あたり4時間かかっている。時給は3000円くらい。",
    items: [
      { itemId: "stk_people", name: "棚卸人数", unit: "人" },
      { itemId: "stk_hours", name: "棚卸時間", unit: "時間" },
      { itemId: "stk_freq", name: "棚卸回数/年", unit: "回" },
      { itemId: "stk_wage", name: "時給", unit: "円" },
    ],
  });
  Logger.log(result.getContent());
}

/* ------------------------------------------------------------
 * 複数課題の一括抽出テスト。1つの議事録から複数カテゴリが
 * 抽出されるか（title・summary が入るか）を確認する。
 * 実行後は「実行数」メニューでログを確認すること。
 * ------------------------------------------------------------ */
function testAutoRun() {
  const result = handleAutoMode({
    caseId: "TEST-01",
    memoText: "",
    hearings: [{
      title: "初回訪問ヒアリング",
      text: "棚卸は5人で月1回、1回4時間かかっている。集計はExcelで手作業のため、"
          + "在庫差異が判明するのが翌月になってしまう。時給は3000円くらい。"
          + "またロット追跡の依頼が入ると、担当者が台帳を探して回答するのに半日かかる。"
          + "追跡依頼は月に10件ほど、対応できるのは2名だけ。"
          + "あと、原価計算のやり方はベテランの田中さんしか分からず、"
          + "田中さんが休むと月次が止まる。これは前から不安に思っている。",
    }],
    categories: [
      { category: "在庫管理", items: [
        { itemId: "stk_people", name: "棚卸人数", unit: "人" },
        { itemId: "stk_hours", name: "棚卸時間", unit: "時間" },
        { itemId: "stk_freq", name: "棚卸回数/年", unit: "回" },
        { itemId: "stk_wage", name: "時給", unit: "円" },
        { itemId: "stk_improve", name: "改善率", unit: "%" }]},
      { category: "ロット管理", items: [
        { itemId: "lot_hours", name: "追跡時間", unit: "時間" },
        { itemId: "lot_freq", name: "追跡回数/年", unit: "回" },
        { itemId: "lot_people", name: "担当人数", unit: "人" },
        { itemId: "lot_wage", name: "時給", unit: "円" }]},
      { category: "属人化", items: [
        { itemId: "attrib_people", name: "該当ベテラン人数", unit: "人" }]},
      { category: "AI議事録", items: [
        { itemId: "min_meetings", name: "会議回数/月", unit: "回" }]},
    ],
  });
  Logger.log(result.getContent());
}

/* ------------------------------------------------------------
 * 参照URLのみ（本文テキストが空）の議事録が正しく扱えるかのテスト。
 * 実在するURLに差し替えて実行し、実行数ログで
 * 「取得したテキストの先頭」を確認すること。
 * ------------------------------------------------------------ */
function testUrlOnlyFetch() {
  const url = "https://example.com/"; // ← 実際の議事録ビューアのURLに差し替える
  const text = fetchTextFromUrl(url);
  Logger.log("取得文字数: " + text.length);
  Logger.log("先頭200文字: " + text.slice(0, 200));
}

/* 動作確認用。ブラウザでURLを直接開いたときに「スクリプト関数が見つかりません」
 * ではなく分かりやすい応答が返るようにする。実行数ログにも残る。 */
function doGet(e) {
  Logger.log("[doGet] called at " + new Date().toISOString());
  return respond({ status: "ok", time: new Date().toISOString(), note: "この提案ナレッジWebアプリは動作しています（doGet）" });
}

/* ブラウザがpreflight（OPTIONS）を送ってきた場合の応答。
 * GAS Web Appはカスタムヘッダーを自由に設定できないため気休め程度だが、
 * 何も応答しないよりは良い。 */
function doOptions(e) {
  Logger.log("[doOptions] called");
  return ContentService.createTextOutput("");
}

function doPost(e) {
  // ここが実行数ログに出るかどうかで、「GASまで届いているか」を切り分けられる。
  try {
    const bodyLen = e && e.postData && e.postData.contents ? e.postData.contents.length : -1;
    const contentType = e && e.postData ? e.postData.type : "(none)";
    Logger.log("[doPost] called at " + new Date().toISOString()
      + " / bodyLength=" + bodyLen + " / contentType=" + contentType);
  } catch (logErr) {
    Logger.log("[doPost] logging failed: " + String(logErr));
  }

  try {
    const req = JSON.parse(e.postData.contents);
    Logger.log("[doPost] mode=" + (req.mode || "(single)") + " caseId=" + req.caseId);

    if (SHARED_TOKEN && req.token !== SHARED_TOKEN) {
      Logger.log("[doPost] unauthorized: token mismatch");
      return respond({ error: "unauthorized" });
    }

    let result;
    if (req.mode === "auto") result = handleAutoMode(req);
    else if (req.mode === "category") result = handleCategoryMode(req);
    else result = handleSingleCategoryMode(req);
    Logger.log("[doPost] responded OK");
    return result;
  } catch (err) {
    Logger.log("[doPost] ERROR: " + String(err) + (err && err.stack ? ("\n" + err.stack) : ""));
    return respond({ error: String(err) });
  }
}

/* 単一カテゴリ抽出（提案ナレッジ側の詳細レビュー、営業報告側の旧フロー互換） */
function handleSingleCategoryMode(req) {
    const settings = getAiSettings();
    if (!settings.apiKey) return respond({ error: "AI_API_KEY not configured" });

    let sourceText = req.text || "";
    if (!sourceText && req.url) {
      sourceText = fetchTextFromUrl(req.url);
    }
    if (!sourceText) return respond({ error: "text and url are both empty" });

    const itemList = (req.items || [])
      .map(i => `- ${i.itemId} (${i.name}, 単位:${i.unit})`)
      .join("\n");

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

    const ai = callAiModel(prompt, 1500, settings);
    if (ai.error) return respond({ error: ai.error });
    const parsed = safeParseJson(ai.text, { items: [] });
    return respond(parsed);
}

/* 複数カテゴリの一括判定（営業報告アドインの「作成」アイコン用）。
 * リクエスト:
 * { mode:"auto", caseId, text: "議事録＋メモを連結したテキスト",
 *   categories: [{ category, items:[{itemId,name,unit}] }, ...] }
 * レスポンス:
 * { results: [{ category, items:[{itemId,value,confidence}] }, ...] }
 * （該当なしのカテゴリは items:[] または results に含めない） */
/* 複数課題の一括抽出（営業報告アドインの「作成」アイコン、提案ナレッジの①タブ）。
 * リクエスト:
 * { mode:"auto", caseId, memoText: "任意のメモ",
 *   hearings: [{ title, text, url }, ...],
 *   categories: [{ category, items:[{itemId,name,unit}] }, ...] }
 * hearings の各要素は text か url のどちらかがあればよい。text が空で url が
 * ある場合は、このサーバー側で UrlFetchApp によりページ内容を取得してから
 * まとめてプロンプトに渡す（認証が必要なページは取得できない）。
 * レスポンス:
 * { results: [{ category, title, summary, items:[{itemId,value,confidence}] }, ...] }
 * 1つの議事録から複数の課題が出るのが前提。数値が読み取れない課題でも、
 * 課題として言及されていれば items:[] で results に含める。 */
function handleAutoMode(req) {
  const settings = getAiSettings();
  if (!settings.apiKey) return respond({ error: "AI_API_KEY not configured" });

  const hearings = req.hearings || [];
  const sections = [];
  if (req.memoText) sections.push(req.memoText);
  hearings.forEach(h => {
    const body = h.text || (h.url ? fetchTextFromUrl(h.url) : "");
    if (body) sections.push(`【${h.title || "議事録"}】\n${body}`);
  });
  const combinedText = sections.join("\n\n");
  if (!combinedText) return respond({ error: "text is empty" });

  const categoryBlock = (req.categories || []).map(c =>
    `### ${c.category}\n${c.items.map(i => `- ${i.itemId} (${i.name}, 単位:${i.unit})`).join("\n")}`
  ).join("\n\n");

  const prompt =
`あなたは中小企業向けの業務システム提案を行う営業担当です。
以下の議事録・メモを読み、顧客が抱えている課題を抽出してください。

【最重要】顧客が困りごととして語っている内容は、下の「既存カテゴリ一覧」に
当てはまるかどうかに関わらず、すべて抽出してください。
既存カテゴリに無いという理由で課題を捨てないでください。
自社のカテゴリ整備が追いついていないだけで、顧客にとっては同じ課題です。

【重要】1つの議事録に複数の課題が含まれているのが普通です。1つに絞らないでください。

各課題について、次を出力してください。

- title: その顧客固有の課題を一行で（例「棚卸と在庫差異の調査に時間がかかっている」）
  カテゴリ名をそのまま書かず、議事録に出てきた具体的な状況を反映すること
- summary: 課題の内容を2〜3文で。現状の進め方、何に困っているか、その影響を含める。
  議事録に書かれていないことは推測で補わないこと
- matchedCategory: 既存カテゴリのどれかに**明確に**当てはまる場合のみ、そのカテゴリ名。
  少しでも迷う場合は空文字 "" にすること（担当者が後で判断する）
- items: matchedCategory を設定した場合のみ、そのカテゴリの項目一覧に対応する数値。
  読み取れない項目は value を null、confidence を "未確認" とする。
  数値が明言されていれば "確定"、文脈から推測した場合は "推定"。
  matchedCategory が空なら items は空配列にすること
- candidates: 担当者がカテゴリを選ぶための候補を1〜3件。既存・新規を混ぜてよい。
  各候補は { name, isNew, reason } の形。
  - name: カテゴリ名。既存カテゴリなら一覧の名前をそのまま使う
  - isNew: 既存カテゴリなら false、新しく作る提案なら true
  - reason: なぜその候補か、および当てはまらない可能性があればその留保も一行で
  新規候補の name は、その顧客固有の言葉ではなく、他業種でも通じる一般的な名前にすること

既存カテゴリ一覧（この案件で使えるもの）:
${categoryBlock || "（まだカテゴリが登録されていません）"}

議事録・メモ:
"""
${combinedText}
"""

出力は次のJSON形式のみとしてください（説明文やコードフェンスは不要）:
{"issues":[{"title":"...","summary":"...","matchedCategory":"","items":[],"candidates":[{"name":"...","isNew":false,"reason":"..."}]}]}`;

  const ai = callAiModel(prompt, 4000, settings);
  if (ai.error) return respond({ error: ai.error });
  const parsed = safeParseJson(ai.text, { issues: [] });
  return respond(parsed);
}

/* Claudeがコードフェンス付き（```json ... ```）で返すことがあるため、
 * JSON.parse前に取り除く。空文字や解析不能な場合はフォールバック値を返す。 */
function safeParseJson(text, fallback) {
  try {
    const cleaned = String(text || "").replace(/```json/gi, "").replace(/```/g, "").trim();
    if (!cleaned) return fallback;
    return JSON.parse(cleaned);
  } catch (e) {
    return fallback;
  }
}
/* カテゴリ設計モード（AIとの壁打ち）。
 * リクエスト:
 * { mode:"category", caseId,
 *   issue: { title, summary },
 *   existingCategories: ["在庫管理", ...],
 *   messages: [{ role:"user"|"assistant", content:"..." }, ...] }
 * レスポンス:
 * { reply: "AIの返答文", definition: { category, inputs:[...], outputs:[...] } }
 * definition は毎回「現時点での案」を返す。担当者が画面で編集して登録する。 */
function handleCategoryMode(req) {
  const settings = getAiSettings();
  if (!settings.apiKey) return respond({ error: "AI_API_KEY not configured" });

  const issue = req.issue || {};
  const history = (req.messages || [])
    .map(m => (m.role === "user" ? "担当者: " : "AI: ") + m.content).join("\n");

  const prompt =
`あなたは中小企業向けの業務システム提案で使うROI試算マスタの設計を手伝います。
担当者と相談しながら、この課題を試算するための課題カテゴリを設計してください。

対象の課題:
タイトル: ${issue.title || ""}
内容: ${issue.summary || ""}

既にあるカテゴリ（重複させないこと）:
${(req.existingCategories || []).join(" / ") || "（なし）"}

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

  const ai = callAiModel(prompt, 2500, settings);
  if (ai.error) return respond({ error: ai.error });
  const parsed = safeParseJson(ai.text, { reply: "", definition: null });
  return respond(parsed);
}

/* 参照URL先のテキストを取得する。
 * 議事録ビューア（Notion等）が公開URLでプレーンテキスト/HTMLを返す前提。
 * 認証が必要なページは取得できないため、その場合はtext欄への貼り付けを使う。 */
function fetchTextFromUrl(url) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const html = res.getContentText();
    return html.replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 8000); // プロンプトが肥大化しすぎないよう上限を設ける
  } catch (e) {
    return "";
  }
}

/* GAS Web AppはカスタムCORSヘッダーを設定できないため、
 * text/plain で返し、クライアント側で JSON.parse する。
 * Content-Type: application/json で返すと preflight に引っかかることがある。 */
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}
