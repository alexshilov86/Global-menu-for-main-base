function openAsExcel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = ss.getActiveSheet().getSheetId();
  var url = ss.getUrl().replace(/\/edit.*$/, "/export?format=xlsx&gid=" + sheetId);
  
  // HTML-вставка для автоматического скачивания
  var html = "<script>window.open('" + url + "'); google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(200)
    .setHeight(50);
    
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Подготовка файла...");
}


