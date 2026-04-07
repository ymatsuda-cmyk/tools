function openKanbanDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost/dialog.html",
    {
      width: 80,
      height: 80,
      displayInIframe: true
    }
  );
}

Office.onReady();