/* 互換用：共有ランタイムを使わない環境でリボンの「グループシート」ボタンを受ける */
/* global Office, Excel */
function openGroupSheet(event) {
  Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject("グループ");
    await ctx.sync();
    if (!ws.isNullObject) { ws.activate(); await ctx.sync(); }
  }).catch((e) => console.error(e)).then(() => event.completed());
}
Office.onReady(() => {
  if (Office.actions && Office.actions.associate) Office.actions.associate("openGroupSheet", openGroupSheet);
});
self.openGroupSheet = openGroupSheet;
