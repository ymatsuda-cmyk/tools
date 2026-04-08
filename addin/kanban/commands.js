Office.onReady(() => {});
function openKanban(event) {
  Office.context.ui.displayDialogAsync(
    "https://ymatsuda-cmyk.github.io/tools/addin/kanban/dialog.html",
    { height: 70, width: 80 }
  );
  event.completed();
}
