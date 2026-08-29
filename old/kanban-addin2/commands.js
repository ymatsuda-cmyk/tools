function openKanbanDialog() {
  Office.context.ui.displayDialogAsync(
    "https://ymatsuda-cmyk.github.io/tools/kanban-addin2/dialog.html",
    {
      width: 80,
      height: 80,
      displayInIframe: true
    }
  );
}

Office.onReady();