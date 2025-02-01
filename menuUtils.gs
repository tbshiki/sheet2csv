function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("add-on")
    .addItem("CSVをエクスポート", "exportFilteredCSV")
    .addToUi();
}
