# sheet2csv

## 概要
本プロジェクトは、Google スプレッドシートから特定のデータを CSV ファイルとしてエクスポートし、Google Drive に保存するシステムです。設定は `Config` シートで管理し、スクリプトを修正せずに運用できます。

---

## 📂 ファイル構成

```
📂 Google Apps Script Project
 ├── exportFilteredCSV.gs   （CSV エクスポートのメイン処理）
 ├── configUtils.gs         （Config シートの設定を取得）
 ├── fileUtils.gs           （Google Drive 操作）
 ├── modalUtils.gs          （モーダル管理）
 ├── dataUtils.gs           （データ処理関数）
 ├── menuUtils.gs           （実行メニュー追加）
 ├── .clasp.json            （Clasp 設定ファイル）
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
- `Config` シートの設定を取得し、データをフィルタリングして CSV ファイルを作成。
- `ExportLog` シートにバックアップを保存。
- Google Drive に CSV ファイルをアップロード。
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

### **7️⃣ .clasp.json**（Clasp 設定ファイル）
- Google Apps Script をローカル環境で管理し、Git 連携やデプロイを可能にする設定ファイル。
- `scriptId` は Google Apps Script のプロジェクト ID。
- `rootDir` はローカルのスクリプトディレクトリを指定。

---

## 🚀 Clasp を使用したデプロイ方法
Google Apps Script のローカル管理には Clasp が必要です。
**Node.js のインストール** が必要なので、事前に [公式サイト](https://nodejs.org/) からダウンロード＆インストールしてください。

### **Clasp のセットアップ**
1. **Node.js をインストール**
   - [公式サイト](https://nodejs.org/) から最新版をダウンロード＆インストール

2. **Clasp をインストール**
   ```sh
   npm install -g @google/clasp
   ```

3. **Google にログイン**
   ```sh
   clasp login
   ```

4. **プロジェクトをクローン**
   - `[[scriptId]]` の部分は **Google Apps Script の ID** に置き換えてください：
   ```sh
   clasp clone [[scriptId]]
   ```

### **スクリプトの更新 & デプロイ**
1. **ローカルで変更を反映**
   - `.clasp.json` があるディレクトリで実行：
   ```sh
   clasp push
   ```

2. **Apps Script エディタで変更を取得**
   ```sh
   clasp pull
   ```

---

## 🎯 まとめ
 - 設定は `Config` シートで管理し、スクリプトの編集不要
 - Google Drive に CSV ファイルを保存し、モーダルでダウンロードリンクを表示
 - カラム指定はアルファベット表記で直感的に変更可能
 - 保護列を指定してデータクリアの安全性を確保
 - スプレッドシートのメニューから簡単に CSV エクスポートを実行可能
 - Clasp を使用してスクリプトをバージョン管理し、デプロイが可能

---
