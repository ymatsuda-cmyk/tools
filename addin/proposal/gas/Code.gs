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
 */

const SHARED_TOKEN = "";              // 空なら照合をスキップ（開発時のみ推奨）

/* AI_API_KEY・AI_API_URL・AI_MODEL はすべてスクリプトプロパティに設定する。
 * スクリプトエディタ → プロジェクトの設定 → スクリプト プロパティ で以下を登録:
 *   AI_API_KEY  … 例: sk-ant-xxxxx（必須）
 *   AI_API_URL  … 例: https://api.anthropic.com/v1/messages（未設定ならデフォルト値を使う）
 *   AI_MODEL    … 例: claude-sonnet-4-6（未設定ならデフォルト値を使う）
 * コードを直接編集しなくても、モデル差し替えやAPI切り替えができるようにするため。 */
function getAiSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    apiKey: props.getProperty("AI_API_KEY"),
    apiUrl: props.getProperty("AI_API_URL") || "https://api.anthropic.com/v1/messages",
    model: props.getProperty("AI_MODEL") || "claude-sonnet-4-6",
  };
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    if (SHARED_TOKEN && req.token !== SHARED_TOKEN) {
      return respond({ error: "unauthorized" });
    }

    if (req.mode === "auto") return handleAutoMode(req);
    return handleSingleCategoryMode(req);
  } catch (err) {
    return respond({ error: String(err) });
  }
}

/* 単一カテゴリ抽出（提案ナレッジ側の詳細レビュー、営業報告側の旧フロー互換） */
function handleSingleCategoryMode(req) {
    const { apiKey, apiUrl, model } = getAiSettings();
    if (!apiKey) return respond({ error: "AI_API_KEY not configured" });

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

    const payload = {
      model: model,
      max_tokens: 1500,
      messages: [{ role: "user", content: prompt }],
    };

    const res = UrlFetchApp.fetch(apiUrl, {
      method: "post",
      contentType: "application/json",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const body = safeParseJson(res.getContentText(), null);
    if (!body) return respond({ error: "AI API response was not valid JSON: " + res.getContentText().slice(0, 300) });
    if (body.error) return respond({ error: "AI API error: " + JSON.stringify(body.error) });
    const textBlock = (body.content || []).find(c => c.type === "text");
    const parsed = safeParseJson(textBlock && textBlock.text, { items: [] });

    return respond(parsed);
}

/* 複数カテゴリの一括判定（営業報告アドインの「作成」アイコン用）。
 * リクエスト:
 * { mode:"auto", caseId, text: "議事録＋メモを連結したテキスト",
 *   categories: [{ category, items:[{itemId,name,unit}] }, ...] }
 * レスポンス:
 * { results: [{ category, items:[{itemId,value,confidence}] }, ...] }
 * （該当なしのカテゴリは items:[] または results に含めない） */
function handleAutoMode(req) {
  const { apiKey, apiUrl, model } = getAiSettings();
  if (!apiKey) return respond({ error: "AI_API_KEY not configured" });
  if (!req.text) return respond({ error: "text is empty" });

  const categoryBlock = (req.categories || []).map(c =>
    `### ${c.category}\n${c.items.map(i => `- ${i.itemId} (${i.name}, 単位:${i.unit})`).join("\n")}`
  ).join("\n\n");

  const prompt =
`以下のテキスト（議事録・メモ）を読み、当てはまる課題カテゴリを判定してください。
各カテゴリの項目一覧に対応する数値が読み取れる場合のみ、そのカテゴリをresultsに含めてください。
数値が全く読み取れないカテゴリは results に含めないでください。
数値の確度は confidence（確定|推定|未確認）で表してください。

カテゴリと項目一覧:
${categoryBlock}

テキスト:
"""
${req.text}
"""

出力は次のJSON形式のみとしてください（説明文は不要）:
{"results":[{"category":"...","items":[{"itemId":"...","value":数値またはnull,"confidence":"確定|推定|未確認"}]}]}`;

  const payload = { model: model, max_tokens: 2000, messages: [{ role: "user", content: prompt }] };
  const res = UrlFetchApp.fetch(apiUrl, {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const body = safeParseJson(res.getContentText(), null);
  if (!body) return respond({ error: "AI API response was not valid JSON: " + res.getContentText().slice(0, 300) });
  if (body.error) return respond({ error: "AI API error: " + JSON.stringify(body.error) });
  const textBlock = (body.content || []).find(c => c.type === "text");
  const parsed = safeParseJson(textBlock && textBlock.text, { results: [] });
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
