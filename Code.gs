// ========================================
// 【設定】ここだけ変更してください
// ========================================
// あなたのスプレッドシートIDをここに貼り付け
// (スプレッドシートのURL: https://docs.google.com/spreadsheets/d/【ここの文字列】/edit)
const SHEET_ID = 'ここにあなたのスプレッドシートIDを貼り付けてください';

// ========================================
// 【初回設定】最初に1回だけ実行してください
// ========================================
// バーコード列を文字列形式に設定し、既存データの'を削除
function initialSetup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. 製品マスタのバーコード列を文字列形式に
  const masterSheet = ss.getSheetByName('製品マスタ');
  masterSheet.getRange("A:A").setNumberFormat("@STRING@");
  masterSheet.getRange("B:B").setNumberFormat("@STRING@");
  
  // 2. 在庫表のバーコード列を文字列形式に
  const inventorySheet = ss.getSheetByName('在庫表');
  inventorySheet.getRange("A:A").setNumberFormat("@STRING@");
  inventorySheet.getRange("B:B").setNumberFormat("@STRING@");
  
  // 3. 既存データの先頭の'を削除
  cleanExistingData(masterSheet);
  cleanExistingData(inventorySheet);
  
  return "初期設定完了！これで普通にバーコードを入力できます";
}

// 既存データの'を削除する内部関数
function cleanExistingData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const dataA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const dataB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  // 'で始まるデータを修正
  const cleanedA = dataA.map(row => {
    let value = String(row[0]);
    return [value.startsWith("'") ? value.substring(1) : value];
  });
  
  const cleanedB = dataB.map(row => {
    let value = String(row[0]);
    return [value.startsWith("'") ? value.substring(1) : value];
  });
  
  sheet.getRange(2, 1, lastRow - 1, 1).setValues(cleanedA);
  sheet.getRange(2, 2, lastRow - 1, 1).setValues(cleanedB);
}

// ========================================
// 基本設定(変更不要)
// ========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('バーコード在庫管理システム');
}

// ========================================
// 品名取得
// ========================================
// 製品マスタからバーコードに対応する品名を取得
// 【製品マスタの構成】
// A列: JANコード(バーコード)
// B列: 商品コード(任意)
// C列: 品名
function getProductName(barcode) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const masterSheet = ss.getSheetByName('製品マスタ');
  
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return 'マスタなし';
  
  const data = masterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const barcodeStr = String(barcode).trim();
  
  for (let i = 0; i < data.length; i++) {
    const cellValue = String(data[i][0]).trim();
    // 'で始まる場合は削除して比較
    const cleanValue = cellValue.startsWith("'") ? cellValue.substring(1) : cellValue;
    
    if (cleanValue === barcodeStr) {
      return data[i][2]; // C列(品名)を返す
    }
  }
  
  return 'バーコードデータがありません';
}

// ========================================
// 在庫更新(1件)
// ========================================
// 在庫表の在庫数を更新
// 【在庫表の構成】
// A列: JANコード(バーコード)
// B列: 商品コード(任意)
// C列: 品名(任意)
// D列: その他(任意)
// E列: 在庫数 ← ここを更新
function updateInventory(barcode, quantity) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const inventorySheet = ss.getSheetByName('在庫表');
  const masterSheet = ss.getSheetByName('製品マスタ');
  
  const lastRow = inventorySheet.getLastRow();
  const barcodeStr = String(barcode).trim();
  
  // 在庫表にデータがある場合は検索
  if (lastRow >= 2) {
    const inventoryData = inventorySheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    for (let i = 0; i < inventoryData.length; i++) {
      const cellValue = String(inventoryData[i][0]).trim();
      const cleanValue = cellValue.startsWith("'") ? cellValue.substring(1) : cellValue;
      
      if (cleanValue === barcodeStr) {
        // 見つかったら在庫更新
        const currentStock = inventoryData[i][4] || 0;
        const newStock = currentStock + quantity;
        inventorySheet.getRange(i + 2, 5).setValue(newStock);
        return;
      }
    }
  }
  
  // 在庫表に見つからなかった→製品マスタから情報取得して自動追加
  const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues();
  
  for (let i = 0; i < masterData.length; i++) {
    const masterBarcode = String(masterData[i][0]).trim();
    const cleanMasterBarcode = masterBarcode.startsWith("'") ? masterBarcode.substring(1) : masterBarcode;
    
    if (cleanMasterBarcode === barcodeStr) {
      // 製品マスタから情報取得
      const productCode = masterData[i][1];
      const productName = masterData[i][2];
      
      // 在庫表に新規追加(初期在庫=quantity)
      const newRow = inventorySheet.getLastRow() + 1;
      inventorySheet.getRange(newRow, 1, 1, 5).setNumberFormats([["@STRING@", "@STRING@", "@", "@", "0"]]);
      inventorySheet.getRange(newRow, 1, 1, 5).setValues([[barcode, productCode, productName, "", quantity]]);
      return;
    }
  }
}

// ========================================
// 在庫更新(一括)
// ========================================
function updateInventoryBatch(barcodeList) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const inventorySheet = ss.getSheetByName('在庫表');
  
  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) return 0;
  
  const inventoryData = inventorySheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  // バーコードごとに数量を集計
  const summary = {};
  barcodeList.forEach(item => {
    const barcode = String(item.barcode).trim();
    summary[barcode] = (summary[barcode] || 0) + item.quantity;
  });
  
  // 在庫表を更新
  let updateCount = 0;
  inventoryData.forEach((row, index) => {
    const cellValue = String(row[0]).trim();
    // 'で始まる場合は削除して比較
    const cleanValue = cellValue.startsWith("'") ? cellValue.substring(1) : cellValue;
    
    if (summary[cleanValue]) {
      const currentStock = row[4] || 0;
      const newStock = currentStock + summary[cleanValue];
      inventorySheet.getRange(index + 2, 5).setValue(newStock);
      updateCount++;
    }
  });
  
  return updateCount;
}

// ========================================
// 1件登録
// ========================================
// スキャン履歴に追記 + 在庫表を更新
function addBarcodeWithQuantity(barcode, quantity) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('スキャン履歴');
  const now = new Date();
  
  const newRow = sheet.getLastRow() + 1;
  const formula = `=IFERROR(VLOOKUP(B${newRow},'製品マスタ'!A:C,3,0),"バーコードデータがありません")`;
  
  // スキャン履歴に追記(A列を文字列形式に設定)
  sheet.getRange(newRow, 1, 1, 4).setNumberFormats([["@", "@STRING@", "0", "@"]]);
  sheet.getRange(newRow, 1, 1, 4).setValues([[now, barcode, quantity, formula]]);
  
  // 在庫表を更新
  updateInventory(barcode, quantity);
  
  return "登録完了";
}

// ========================================
// 一括登録
// ========================================
function addBarcodesBatch(barcodeList) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('スキャン履歴');
  const now = new Date();
  const startRow = sheet.getLastRow() + 1;
  
  const rows = barcodeList.map((item, index) => {
    const rowNum = startRow + index;
    const formula = `=IFERROR(VLOOKUP(B${rowNum},'製品マスタ'!A:C,3,0),"バーコードデータがありません")`;
    return [now, item.barcode, item.quantity, formula];
  });
  
  // スキャン履歴に一括追記
  if (rows.length > 0) {
    // 書式を文字列に設定
    const formats = rows.map(() => ["@", "@STRING@", "0", "@"]);
    sheet.getRange(startRow, 1, rows.length, 4).setNumberFormats(formats);
    sheet.getRange(startRow, 1, rows.length, 4).setValues(rows);
  }
  
  // 在庫表を一括更新
  const updateCount = updateInventoryBatch(barcodeList);
  
  return "一括登録完了: " + rows.length + "件 (在庫更新: " + updateCount + "件)";
}
