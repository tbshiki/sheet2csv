/**
 * `Config` シートから設定を取得する関数
 * @param {Spreadsheet} ss - スプレッドシートのオブジェクト
 * @returns {Object} 設定データ
 */
function getConfigSettings(ss) {
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
      throw new Error("Config シートが見つかりません。");
    }

    var configData = configSheet.getRange("A:B").getValues(); // C列（備考）は無視
    var configMap = {};

    for (var i = 0; i < configData.length; i++) {
      var key = configData[i][0]; // 設定キー（A列）
      var value = configData[i][1]; // 設定値（B列）

      if (!key) continue; // 空白のキーをスキップ

      switch (key) {
        case "inputSheet":
        case "backupSheet":
        case "historySheet":
        case "folderName":
          configMap[key] = value;
          break;
        case "targetColumnIndex":
          configMap[key] = columnLetterToNumber(value); // アルファベット → 数値変換
          break;
        case "protectedColumns":
          configMap[key] = value.split(",").map(col => columnLetterToNumber(col.trim()));
          break;
      }
    }

    return configMap;
  }

  /**
   * 列のアルファベットを数値に変換（例: "B" → 2, "AA" → 27）
   * @param {string} columnLetter - 列のアルファベット
   * @returns {number} 列番号（1-indexed）
   */
  function columnLetterToNumber(columnLetter) {
    var column = columnLetter.toUpperCase();
    var number = 0;
    for (var i = 0; i < column.length; i++) {
      number = number * 26 + (column.charCodeAt(i) - 64);
    }
    return number;
  }
