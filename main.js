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
  
  // シートのヘッダー（1行目）
  var header = data[0];

  // 2行目以降で対象列（B列）に値がある行をフィルタリング
  var filteredData = data.filter((row, index) => index !== 0 && row[targetColumnIndex - 1]); // 1-indexedのため-1
  if (filteredData.length === 0) {
    return showModal("エクスポートするデータがありません。");
  }

  // **CSVデータ作成（ヘッダーを先頭に追加）**
  var csvData = [header].concat(filteredData); // ヘッダーを追加
  var csvContent = csvData.map(row => row.join(",")).join("\n");

  // タイムスタンプ付きファイル名
  var timestamp = new Date();
  var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  var fileName = `filtered_export_${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}.csv`;
  
  // Google Drive の特定フォルダに保存
  var folder = getOrCreateFolder(folderName);
  var file = folder.createFile(fileName, csvContent, MimeType.CSV);
  var fileId = file.getId();
  var downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

  // **バックアップ用シートへデータを2行目に挿入**
  if (backupSheet.getLastRow() === 0) {
    backupSheet.appendRow(header); // ヘッダーをコピー（最初の行がない場合）
  }

  backupSheet.insertRows(2, filteredData.length); // 2行目に新規行を追加
  var insertedRange = backupSheet.getRange(2, 1, filteredData.length, filteredData[0].length);
  insertedRange.setValues(filteredData);

  // **2行目の背景色を取得**
  var firstRowBackground = backupSheet.getRange(2, 1, 1, 1).getBackground();
  
  // **背景色の適用**
  if (firstRowBackground === "#ffffff" || firstRowBackground === "") {
    // 2行目が白または背景色なし → 追加データの背景色を「明るいグレー1」にする
    insertedRange.setBackground("#f3f3f3");
  } else {
    // 2行目が白以外の色 → 追加データの背景色をクリアする（白に戻す）
    insertedRange.setBackground("#ffffff");
  }

  // **ログシート `ExportLogHistory` へ記録（2行目に挿入）**
  historySheet.insertRowBefore(2); // 2行目に新規行を追加
  historySheet.getRange(2, 1, 1, 2).setValues([[downloadUrl, timestamp]]); // A列=URL, B列=タイムスタンプ（日時形式）

  // B列のタイムスタンプを自動フォーマット
  historySheet.getRange(2, 2).setNumberFormat("yyyy/MM/dd HH:mm:ss");

  // **元データのクリア（対象列を除外）**
  var rowsToDelete = filteredData.map(row => data.indexOf(row) + 1); // 行番号を取得
  rowsToDelete.forEach(rowNum => {
    var totalColumns = sheet.getLastColumn();
    for (var col = 1; col <= totalColumns; col++) {
      if (!protectedColumns.includes(col)) {
        sheet.getRange(rowNum, col).clearContent();
      }
    }
  });

  // ダウンロードリンクをモーダルで表示
  showModal(`CSVファイルが作成されました。<br>
    <a href="${downloadUrl}" target="_blank" download="${fileName}">[CSVをダウンロード]</a>`);
}

/**
 * 指定したフォルダを取得、なければ作成する関数
 */
function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/**
 * モーダルを表示する関数 (サイズ 400x200)
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
