function exportFilteredCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // === シート名（管理しやすく変更可能） ===
  var sheetNames = {
    input: "InputData",        // 入力用シート
    backup: "ExportLog",       // バックアップ用シート
    history: "ExportLogHistory" // CSV出力ログ用シート
  };

  // === データの処理対象列 ===
  var targetColumnIndex = 2; // B列を基準（1-indexed）
  var protectedColumns = [1, 3, 7, 21, 22]; // A, C, G, U, V列（1-indexed）

  // === 保存フォルダ名 ===
  var folderName = "CSV_Exports";

  var sheet = ss.getSheetByName(sheetNames.input);
  if (!sheet) {
    showModal(`シート "${sheetNames.input}" が見つかりません。`);
    return;
  }

  var backupSheet = ss.getSheetByName(sheetNames.backup);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(sheetNames.backup);
  }

  var historySheet = ss.getSheetByName(sheetNames.history);
  if (!historySheet) {
    historySheet = ss.insertSheet(sheetNames.history);
    historySheet.appendRow(["Download URL", "Timestamp"]); // ヘッダーを追加
  }

  var data = sheet.getDataRange().getValues();
  var header = data[0]; // シートのヘッダー（1行目）

  // 2行目以降で対象列（B列）に値がある行をフィルタリング
  var filteredData = data.filter((row, index) => index !== 0 && row[targetColumnIndex - 1]);
  if (filteredData.length === 0) {
    return showModal("エクスポートするデータがありません。");
  }

  // **バックアップ用シートへデータを最初に移動**
  if (backupSheet.getLastRow() === 0) {
    backupSheet.appendRow(header);
  }

  backupSheet.insertRows(2, filteredData.length);
  var insertedRange = backupSheet.getRange(2, 1, filteredData.length, filteredData[0].length);
  insertedRange.setValues(filteredData);

  // **2行目の背景色を取得**
  var firstRowBackground = backupSheet.getRange(2, 1, 1, 1).getBackground();

  // **背景色の適用**
  insertedRange.setBackground(firstRowBackground === "#ffffff" || firstRowBackground === "" ? "#f3f3f3" : "#ffffff");

  // **CSVデータ作成（バックアップ完了後に実施）**
  var csvData = [header].concat(filteredData);
  var csvContent = csvData.map(row => row.join(",")).join("\n");

  // タイムスタンプ付きファイル名
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  var fileName = `filtered_export_${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}.csv`;

  // Google Drive に保存
  var folder = getOrCreateFolder(folderName);
  var file = folder.createFile(fileName, csvContent, MimeType.CSV);
  var fileId = file.getId();
  var downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

  // **ログシートへ記録**
  historySheet.insertRowBefore(2);
  historySheet.getRange(2, 1, 1, 2).setValues([[downloadUrl, timestamp]]);
  historySheet.getRange(2, 2).setNumberFormat("yyyy/MM/dd HH:mm:ss");

  // **元データのクリア（バックアップと CSV 生成後に実行）**
  var rowsToDelete = filteredData.map(row => data.indexOf(row) + 1);
  rowsToDelete.forEach(rowNum => {
    for (var col = 1; col <= sheet.getLastColumn(); col++) {
      if (!protectedColumns.includes(col)) {
        sheet.getRange(rowNum, col).clearContent();
      }
    }
  });

  // **ダウンロードリンクをモーダルで表示**
  showModal(`CSVファイルが作成されました。<br>
    <a href="${downloadUrl}" target="_blank" download="${fileName}">[CSVをダウンロード]</a>`);
}
