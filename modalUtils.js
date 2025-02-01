/**
 * モーダルを表示する関数
 */
function showModal(message) {
    var html = HtmlService.createHtmlOutput(
      `<div style="padding: 20px; text-align: center;">
        <p>${message}</p>
        <button onclick="google.script.host.close()" style="padding: 10px;">閉じる</button>
      </div>`
    )
    .setWidth(400)
    .setHeight(200);

    SpreadsheetApp.getUi().showModalDialog(html, "CSVエクスポート");
  }
