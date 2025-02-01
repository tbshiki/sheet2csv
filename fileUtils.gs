/**
 * 指定したフォルダを取得、なければ作成する関数
 */
function getOrCreateFolder(folderName) {
    var folders = DriveApp.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  }
