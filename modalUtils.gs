/**
 * モーダルを表示する関数
 */
function showModal(message, downloadUrl = "", fileName = "") {
  var buttonHtml = downloadUrl
    ? `<button onclick="window.open('${downloadUrl}', '_blank')"
         style="padding:10px 20px; background:#007bff; color:#fff; border:none; border-radius:5px; cursor:pointer; font-size:14px;">
         CSVをダウンロード
       </button>`
    : "";

  var html = HtmlService.createHtmlOutput(
    `<div style="padding: 20px; text-align: center;">
      <p>${message}</p>
      ${buttonHtml}
      <br><br>
      <a href="#" onclick="google.script.host.close()" style="color:#007bff; text-decoration:none;">閉じる</a>
    </div>`
  )
  .setWidth(350)
  .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, "CSVエクスポート");
}
