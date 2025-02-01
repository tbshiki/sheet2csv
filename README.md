# Google Apps Script - CSV Export System

## 概要
本プロジェクトは、Google スプレッドシートから特定のデータを CSV 形式でエクスポートし、Google Drive に保存するシステムです。設定は `Config` シートで管理し、スクリプトを修正せずに運用できます。

---

## 📂 ファイル構成

```
📂 Google Apps Script Project
 ├── exportFilteredCSV.gs  （CSV エクスポートのメイン処理）
 ├── configUtils.gs         （Config シートの設定を取得）
 ├── fileUtils.gs           （Google Drive 操作）
 ├── modalUtils.gs          （モーダル管理）
 ├── dataUtils.gs           （データ処理関数）
 ├── menuUtils.gs           （実行メニュー追加）
```

---

## 📜 設定管理 (`Config` シート)
スクリプトの動作設定は `Config` シートで管理されます。

### **📝 Config シートのフォーマット**
| 設定キー          | 設定値 | 備考 |
|-----------------|----|------------------|
| inputSheet      | InputData | 入力用シート名 |
| backupSheet     | ExportLog | バックアップ用シート名 |
| historySheet    | ExportLogHistory | CSV出力ログ用シート名 |
| folderName      | CSV_Exports | Google Drive の保存フォルダ名 |
| targetColumnIndex | B  | データの処理対象列（アルファベット表記） |
| protectedColumns | A,C,G,U,V | 保護対象の列（アルファベット表記） |

---

## 🔧 各スクリプトの説明

### **1️⃣ exportFilteredCSV.gs**（CSV エクスポートのメイン処理）
- `Config` シートの設定を取得し、データをフィルタリングして CSV を作成。
- `ExportLog` シートにバックアップを保存。
- Google Drive に CSV をアップロード。
- `ExportLogHistory` にダウンロード URL を記録。
- 対象データの内容をクリア。
- `showModal()` でユーザーに通知。

### **2️⃣ configUtils.gs**（設定取得）
- `Config` シートから各設定を取得する。
- `targetColumnIndex` や `protectedColumns` は **アルファベットを列番号に変換** する。

### **3️⃣ fileUtils.gs**（Google Drive 操作）
- `getOrCreateFolder(folderName)`
  - 指定したフォルダが存在しない場合は作成。

### **4️⃣ modalUtils.gs**（モーダル管理）
- `showModal(message, downloadUrl, fileName)`
  - CSV ダウンロードボタン付きのモーダルを表示。
  - `閉じる` ボタンはシンプルなテキストリンクに。

### **5️⃣ dataUtils.gs**（データ処理）
- `clearSheetData(sheet, filteredData, protectedColumns)`
  - 保護列を除外しながら、対象データを一括クリア。

### **6️⃣ menuUtils.gs**（スプレッドシートメニュー追加）
- `onOpen()`
  - Google スプレッドシートの UI に `add-on` メニューを追加。
  - `CSVをエクスポート` メニューから `exportFilteredCSV()` を実行可能。

---

## 🚀 設定変更方法
スクリプトを編集せずに、`Config` シートの値を変更するだけで以下の設定が可能です。
- **データの取得元のシートを変更** → `inputSheet`
- **バックアップの保存先を変更** → `backupSheet`
- **ログ記録のシートを変更** → `historySheet`
- **Google Drive の保存フォルダを変更** → `folderName`
- **対象列を変更（例: B → C）** → `targetColumnIndex`
- **編集不可の列を変更（例: A,C,G,U,V → D,E,F）** → `protectedColumns`

---

## ⚡ 実行手順
1. **Google スプレッドシートを開く**
2. **必要に応じて `Config` シートの設定を変更**
3. **スクリプトを実行（`exportFilteredCSV()` を実行）**
4. **モーダルでダウンロードリンクが表示される**
5. **Google Drive に CSV が保存される**

---

## 🛠 よくある質問（FAQ）
### **Q1: 設定を変更しても反映されません。**
- `Config` シートのキー名が正しく入力されているか確認してください。
- 設定を更新後、スクリプトを再実行してください。

### **Q2: CSV ダウンロードのボタンが反応しません。**
- Google Apps Script のポップアップブロックが原因の可能性があります。
- `window.open(downloadUrl, '_blank')` を使用して開く設計になっています。

### **Q3: 設定を増やしたい場合は？**
- `Config` シートに新しい設定を追加。
- `configUtils.gs` に新しい設定の読み取り処理を追加。

---

## 🎯 まとめ
✅ **設定は `Config` シートで管理し、スクリプトの編集不要！**
✅ **Google Drive に CSV を保存し、モーダルでダウンロードリンクを表示！**
✅ **カラム指定はアルファベット表記で直感的に変更可能！**
✅ **保護列を指定してデータクリアの安全性を確保！**
✅ **スプレッドシートのメニューから簡単に CSV エクスポートを実行可能！**

---
