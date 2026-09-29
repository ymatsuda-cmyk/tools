/**
 * export-wbs.ts — Power Automate「スクリプトの実行」用 Office Script
 * wbsシートをカンバン(addin/kanban/kanban.js)と同じ列ルールで読み、ダッシュボード用JSONを返す。
 * 出力は OneDrive に置き、Mac の jsonbin_sync.py が暗号化して JSONBin に登録する（api/kanban/README.md）。
 * テーブル化不要・見出し名に依存しない（列位置で取得）。
 * 日付はExcelシリアル値のまま返す（ダッシュボード側で変換）。
 */
function main(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet("wbs");
  if (!sheet) return JSON.stringify({ error: "wbs sheet not found" });

  const used = sheet.getUsedRange();
  const values = used ? used.getValues() : [];
  const rows = values.slice(10); // 11行目以降（kanban.jsと同じ）

  const tasks = [];
  rows.forEach((r, i) => {
    const title = r[25];            // Z
    if (!title || r[19] === "-") return; // T が "-" は除外
    tasks.push({
      id: r[24],                    // Y
      title: String(title),
      category: r[0] || "",         // A
      classification: r[1] || "",   // B
      user: r[13] || "",            // N
      note: String(r[14] || ""),    // O（★/▲判定に使用）
      start: r[15] || "",           // P 予定開始
      end: r[16] || "",             // Q 予定完了
      actualStart: r[17] || "",     // R 実績開始
      actualEnd: r[18] || "",       // S 実績完了
      row: i + 11
    });
  });

  return JSON.stringify({
    schema: "wbs-tasks/v1",
    updatedAt: new Date().toISOString(),
    tasks
  });
}
