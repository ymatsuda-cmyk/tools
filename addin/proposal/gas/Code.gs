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
 * リクエスト（app.js runExtraction から送られる）:
 * {
 *   token: "簡易トークン（任意、照合用）",
 *   caseId: "KM-01",
 *   category: "在庫管理",
 *   text: "議事録の文字起こし・メモ",
 *   items: [{ itemId: "stk_people", name: "棚卸人数", unit: "人" }, ...]
 * }
 *
 * レスポンス:
 * { items: [{ itemId: "stk_people", value: 5, confidence: "確定" }, ...] }
 */

const SHARED_TOKEN = "";              // 空なら照合をスキップ（開発時のみ推奨）
const AI_API_URL = "https://api.anthropic.com/v1/messages";
const AI_MODEL = "claude-sonnet-4-6"; // 登録するモデルをここで指定する

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    if (SHARED_TOKEN && req.token !== SHARED_TOKEN) {
      return respond({ error: "unauthorized" });
    }

    const apiKey = PropertiesService.getScriptProperties().getProperty("AI_API_KEY");
    if (!apiKey) return respond({ error: "AI_API_KEY not configured" });

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
${req.text}
"""

出力は次のJSON形式のみとしてください（説明文は不要）:
{"items":[{"itemId":"...","value":数値またはnull,"confidence":"確定|推定|未確認"}]}`;

    const payload = {
      model: AI_MODEL,
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    };

    const res = UrlFetchApp.fetch(AI_API_URL, {
      method: "post",
      contentType: "application/json",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const body = JSON.parse(res.getContentText());
    const textBlock = (body.content || []).find(c => c.type === "text");
    const parsed = JSON.parse((textBlock && textBlock.text) || '{"items":[]}');

    return respond(parsed);
  } catch (err) {
    return respond({ error: String(err) });
  }
}

/* GAS Web AppはカスタムCORSヘッダーを設定できないため、
 * text/plain で返し、クライアント側で JSON.parse する。
 * Content-Type: application/json で返すと preflight に引っかかることがある。 */
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}
