/**
 * シートのデータを効率よくクリアする関数
 *
 * @param {Sheet} sheet - 操作対象のシート
 * @param {Array} filteredData - 削除対象となるデータ行の配列
 * @param {Array} protectedColumns - 保護対象の列（1-indexed の列番号）
 */
function clearSheetData(sheet, filteredData, protectedColumns) {
  var totalColumns = sheet.getLastColumn();
  var allValues = sheet.getDataRange().getValues(); // シート全体のデータを取得
  var rowNumbers = []; // 削除対象の行番号リスト（1-indexed）

  // **行番号を正しく取得**
  var data = sheet.getDataRange().getValues();
  filteredData.forEach(row => {
    for (var i = 1; i < data.length; i++) { // 1行目（ヘッダー）は除外
      if (JSON.stringify(data[i]) === JSON.stringify(row)) {
        rowNumbers.push(i + 1); // 1-based index
        break;
      }
    }
  });

  if (rowNumbers.length === 0) return;

  // **データのクリア（高速化）**
  rowNumbers.forEach(rowNum => {
    for (var col = 1; col <= totalColumns; col++) {
      if (!protectedColumns.includes(col)) {
        allValues[rowNum - 1][col - 1] = ""; // セルを空にする
      }
    }
  });

  // **setValues で一括更新**
  sheet.getDataRange().setValues(allValues);
}
