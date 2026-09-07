// ================================================================
// AIスコアリング機能（Claude APIを直接呼び出し）
// 既存の code_complete.gs（または相当のGASファイル）に追記して使う
// ================================================================

// --- 設定 ---------------------------------------------------------
// 事前に「プロジェクトの設定」→「スクリプト プロパティ」で以下を登録しておくこと:
//   ANTHROPIC_API_KEY : Anthropic APIキー（console.anthropic.com で発行）
//   NOTION_TOKEN      : 既存のNotion連携トークン（既に登録済みのはず）
var WEB_DB_ID = '35a0e7a535dc803eab2ff9f11444e3f7'; // webDB（記事DB）
var CLAUDE_MODEL = 'claude-haiku-4-5-20251001'; // コスト重視。精度を上げたい場合は 'claude-sonnet-5' に変更

// スコアリング基準（web-summarizerスキルの評価軸と同一に固定）
var SCORING_SYSTEM_PROMPT =
  'あなたは記事をビジネス活用の観点で評価するアシスタントです。\n' +
  '評価対象者は、ロボットアーム制御による決済端末の回帰テスト自動化を専門とし、' +
  'cobot／物理UI自動化市場に関心を持つエンジニアです。\n\n' +
  '与えられた記事本文（またはタイトルとスニペット）をもとに、次のJSON形式のみを出力してください。' +
  '前置きや説明文、Markdownのコードフェンスは一切付けないこと。\n\n' +
  '{\n' +
  '  "要約": "3〜5文の日本語要約（2000文字以内）",\n' +
  '  "解説": "IT初心者向けの解説。専門用語をかみ砕く（300文字程度）",\n' +
  '  "活用": "■ 見出し形式で2〜3個のビジネス活用アイデアと効果。最後に【期待される効果】を1〜2文（400文字程度）",\n' +
  '  "確度スコア": 0から100の整数。自社関連度（ロボットアーム・決済端末回帰テスト・cobot市場との関連性）を最重要視し、次に実装難易度と時間軸を加味する,\n' +
  '  "実装難易度": "低" | "中" | "高",\n' +
  '  "時間軸": "即時" | "6ヶ月" | "中期" | "情報のみ"\n' +
  '}\n\n' +
  '確度スコアの目安: 80-100=自社業務に直接関連し実装容易 / 50-79=関連はあるが検証が必要、または関連度中程度で実装容易 ' +
  '/ 20-49=参考程度、業界動向として把握する価値のみ / 0-19=自社と無関係な一般ニュース。';

// --- URLから本文テキストを取得（簡易HTML→テキスト変換） -------------
function fetchArticleText_(url){
  try{
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if(res.getResponseCode() >= 400) return '';
    var html = res.getContentText();
    // script/styleを除去 → タグ除去 → 空白正規化
    var text = html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
    return text.substring(0, 6000); // プロンプトサイズを抑えるため6000文字で切る
  }catch(e){
    return '';
  }
}

// --- Claude APIを呼んでスコアリングJSONを得る -----------------------
function callClaudeForScoring_(title, url, bodyText){
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if(!apiKey) throw new Error('ANTHROPIC_API_KEY が未設定です（スクリプトプロパティを確認）');

  var userContent = 'タイトル: ' + title + '\nURL: ' + url + '\n\n本文:\n' +
    (bodyText || '(本文取得不可。タイトルとURLのみから推測して評価してください)');

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 1500,
    system: SCORING_SYSTEM_PROMPT,
    messages: [{ role: 'user', content: userContent }]
  };

  var options = {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var json = JSON.parse(res.getContentText());
  if(res.getResponseCode() >= 400){
    throw new Error('Claude API error: ' + (json.error ? json.error.message : res.getContentText()));
  }

  var raw = json.content && json.content[0] && json.content[0].text ? json.content[0].text : '';
  // 万一コードフェンスが付いた場合に備えて除去
  raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

  var parsed;
  try{ parsed = JSON.parse(raw); }
  catch(e){ throw new Error('Claudeの応答をJSONとして解釈できませんでした: ' + raw.substring(0,200)); }

  return parsed;
}

// --- Notionへの書き込み（日本語プロパティ名を壊さないようBlobで送信） ---
function updateWebDbPage_(pageId, scoring){
  var token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  var 処理状態 = (scoring['確度スコア'] >= 50) ? '要検討' : 'アーカイブ';

  var props = {
    '要約': { rich_text: [{ text: { content: String(scoring['要約']||'').substring(0,2000) } }] },
    '解説': { rich_text: [{ text: { content: String(scoring['解説']||'').substring(0,2000) } }] },
    '活用': { rich_text: [{ text: { content: String(scoring['活用']||'').substring(0,2000) } }] },
    '確度スコア': { number: scoring['確度スコア'] },
    '実装難易度': { select: { name: scoring['実装難易度'] } },
    '時間軸': { select: { name: scoring['時間軸'] } },
    '処理状態': { select: { name: 処理状態 } }
  };

  var bodyStr = JSON.stringify({ properties: props });
  var options = {
    method: 'patch',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Notion-Version': '2022-06-28'
    },
    payload: Utilities.newBlob(bodyStr, 'application/json; charset=UTF-8').getBytes(),
    muteHttpExceptions: true
  };

  var res = UrlFetchApp.fetch('https://api.notion.com/v1/pages/' + pageId, options);
  if(res.getResponseCode() >= 400){
    throw new Error('Notion更新失敗: ' + res.getContentText());
  }
  return 処理状態;
}

// --- 記事1件をスコアリングして書き込む（handleReqから呼ばれる） -------
function handleScoreArticle(payload){
  var pageId = payload.pageId;
  var title = payload.title || '';
  var url = payload.url || '';
  if(!pageId) return { error: 'pageIdが必要です' };

  var bodyText = url ? fetchArticleText_(url) : '';
  var scoring;
  try{
    scoring = callClaudeForScoring_(title, url, bodyText);
  }catch(e){
    return { error: e.message };
  }

  var 処理状態;
  try{
    処理状態 = updateWebDbPage_(pageId, scoring);
  }catch(e){
    return { error: e.message };
  }

  return {
    pageId: pageId,
    score: scoring['確度スコア'],
    difficulty: scoring['実装難易度'],
    timeframe: scoring['時間軸'],
    processingStatus: 処理状態,
    summary: scoring['要約']
  };
}

// --- 未処理記事をまとめてスコアリング（バッチ実行） -------------------
// GAS実行時間の上限（6分）を考慮し、1回の呼び出しでは既定10件まで処理する
function handleScoreUnprocessed(payload){
  var limit = (payload && payload.limit) ? Number(payload.limit) : 10;
  var token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');

  var queryBody = {
    filter: { property: '確度スコア', number: { is_empty: true } },
    page_size: limit
  };
  var options = {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' },
    payload: JSON.stringify(queryBody),
    muteHttpExceptions: true
  };
  var res = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + WEB_DB_ID + '/query', options);
  var json = JSON.parse(res.getContentText());
  if(res.getResponseCode() >= 400) return { error: 'Notion検索失敗: ' + res.getContentText() };

  var results = json.results || [];
  var done = [], failed = [];

  results.forEach(function(page){
    var titleProp = page.properties['NAME'];
    var urlProp = page.properties['userDefined:URL'] || page.properties['URL'];
    var title = (titleProp && titleProp.title && titleProp.title[0]) ? titleProp.title[0].plain_text : '';
    var url = (urlProp && urlProp.url) ? urlProp.url : '';

    var r = handleScoreArticle({ pageId: page.id, title: title, url: url });
    if(r.error) failed.push({ pageId: page.id, title: title, error: r.error });
    else done.push({ pageId: page.id, title: title, score: r.score });

    Utilities.sleep(300); // Claude APIのレート制限に配慮
  });

  return {
    processed: done.length,
    failed: failed.length,
    hasMore: results.length === limit,
    done: done,
    errors: failed
  };
}

// ================================================================
// handleReq(e) の switch 文に以下2行を追加すること:
//
//   case 'scoreArticle':
//     return jsonResponse(handleScoreArticle(payload));
//   case 'scoreUnprocessed':
//     return jsonResponse(handleScoreUnprocessed(payload));
//
// ================================================================
