/**
 * CSVインポート・在庫管理自動化機能
 * 送料データのCSVインポート、在庫管理の自動採番、データ検証機能
 */

/**
 * CSVファイルをアップロードして送料データをインポート
 */
function importShippingCSV() {
  const ui = SpreadsheetApp.getUi();
  
  // ファイル選択ダイアログを表示
  const result = ui.prompt(
    '送料CSVインポート',
    'CSVファイルの内容を貼り付けてください（カンマ区切り）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const csvContent = result.getResponseText();
    processShippingCSV(csvContent);
  }
}

/**
 * 送料CSVデータを処理
 */
function processShippingCSV(csvContent) {
  const shippingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('送料管理シート');
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('販売管理シート（利益計算）');
  
  // CSVをパース
  const lines = csvContent.split('\n');
  let newRowIndex = shippingSheet.getLastRow() + 1;
  
  for (let i = 1; i < lines.length; i++) { // ヘッダーをスキップ
    const line = lines[i].trim();
    if (!line) continue;
    
    const columns = line.split(',');
    if (columns.length >= 4) {
      const orderId = columns[0].trim();
      const shippingCost = parseFloat(columns[1]) || 0;
      const shippingMethod = columns[2].trim();
      const shippingDate = columns[3].trim();
      
      // 商品IDと商品名を販売管理シートから取得
      const productInfo = findProductByOrderId(orderId, salesSheet);
      
      const rowData = [
        orderId, // 注文ID
        productInfo ? productInfo.productId : '', // 商品ID
        productInfo ? productInfo.productName : '', // 商品名
        shippingCost, // 送料
        shippingMethod, // 配送方法
        shippingDate ? new Date(shippingDate) : '', // 配送日
        '配送中' // 配送状況
      ];
      
      shippingSheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
      newRowIndex++;
      
      // 販売管理シートの送料を更新
      if (productInfo) {
        updateSalesSheetShipping(orderId, shippingCost, salesSheet);
      }
    }
  }
  
  console.log(`${newRowIndex - shippingSheet.getLastRow() - 1}件の送料データをインポートしました。`);
}

/**
 * 注文IDで販売管理シートから商品情報を検索
 */
function findProductByOrderId(orderId, salesSheet) {
  const data = salesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) { // 販売IDが一致
      return {
        productId: data[i][1], // 商品ID
        productName: data[i][2] // 商品名
      };
    }
  }
  
  return null;
}

/**
 * 販売管理シートの送料を更新
 */
function updateSalesSheetShipping(orderId, shippingCost, salesSheet) {
  const data = salesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) { // 販売IDが一致
      salesSheet.getRange(i + 1, 8).setValue(shippingCost); // 送料
      break;
    }
  }
}

/**
 * 在庫管理シートに新しい商品を追加
 */
function addNewInventoryItem() {
  const ui = SpreadsheetApp.getUi();
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  
  // 商品情報入力ダイアログ
  const productName = ui.prompt('商品名', '商品名を入力してください:').getResponseText();
  if (!productName) return;
  
  const purchasePrice = ui.prompt('仕入れ価格', '仕入れ価格を入力してください（円）:').getResponseText();
  if (!purchasePrice) return;
  
  const purchaseDate = ui.prompt('仕入れ日', '仕入れ日を入力してください（YYYY-MM-DD）:').getResponseText();
  if (!purchaseDate) return;
  
  const supplier = ui.prompt('仕入れ先', '仕入れ先を入力してください（メルカリ/ヤフオク/Amazon等）:').getResponseText();
  if (!supplier) return;
  
  const supplierName = ui.prompt('仕入れ先のユーザー名/店舗名', '仕入れ先のユーザー名または店舗名を入力してください:').getResponseText();
  if (!supplierName) return;
  
  const condition = ui.prompt('商品の状態', '商品の状態を入力してください:').getResponseText();
  
  const notes = ui.prompt('備考', '備考があれば入力してください:').getResponseText();
  
  const researchFlag = ui.alert('リサーチ商品', 'この商品はリサーチ商品ですか？', ui.ButtonSet.YES_NO);
  const isResearch = researchFlag === ui.Button.YES;
  
  let researchStaff = '';
  if (isResearch) {
    researchStaff = ui.prompt('リサーチ担当者', 'リサーチ担当者名を入力してください:').getResponseText();
  }
  
  // 商品IDを自動採番
  const productId = generateProductId(inventorySheet);
  
  // 仕入れ先区分を自動判定
  const supplierType = determineSupplierType(supplier);
  
  // 還付率を自動設定
  const refundRate = supplierType === '個人' ? 8 : 10;
  
  // 新しい行を追加
  const newRowIndex = inventorySheet.getLastRow() + 1;
  const rowData = [
    productId, // 商品ID
    productName, // 商品名
    parseFloat(purchasePrice), // 仕入れ価格
    new Date(purchaseDate), // 仕入れ日
    supplier, // 仕入れ先
    supplierName, // 仕入れ先のユーザー名/店舗名
    condition, // 商品の状態
    notes, // 備考
    '在庫', // 販売状況
    '', // 販売日
    '', // 販売価格
    '', // 利益
    isResearch ? 'Y' : 'N', // リサーチ商品フラグ
    researchStaff, // リサーチ担当者
    supplierType, // 仕入れ先区分
    refundRate // 還付率
  ];
  
  inventorySheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
  
  console.log(`新しい商品「${productName}」を在庫管理シートに追加しました。商品ID: ${productId}`);
}

/**
 * 商品IDを自動採番
 */
function generateProductId(inventorySheet) {
  const data = inventorySheet.getDataRange().getValues();
  let maxId = 0;
  
  for (let i = 1; i < data.length; i++) {
    const currentId = data[i][0];
    if (currentId && typeof currentId === 'string' && currentId.startsWith('PROD')) {
      const idNumber = parseInt(currentId.replace('PROD', ''));
      if (idNumber > maxId) {
        maxId = idNumber;
      }
    }
  }
  
  return `PROD${String(maxId + 1).padStart(6, '0')}`;
}

/**
 * 仕入れ先から仕入れ先区分を自動判定
 */
function determineSupplierType(supplier) {
  const personalSuppliers = ['メルカリ', 'ヤフオク', 'ヤフオク!', 'ヤフオク！', 'mercari'];
  const storeSuppliers = ['Amazon', '楽天市場', '楽天', 'amazon', 'rakuten'];
  
  const lowerSupplier = supplier.toLowerCase();
  
  for (const personal of personalSuppliers) {
    if (lowerSupplier.includes(personal.toLowerCase())) {
      return '個人';
    }
  }
  
  for (const store of storeSuppliers) {
    if (lowerSupplier.includes(store.toLowerCase())) {
      return '商店';
    }
  }
  
  // デフォルトは個人
  return '個人';
}

/**
 * 在庫管理シートのデータを検証
 */
function validateInventoryData() {
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  const data = inventorySheet.getDataRange().getValues();
  const errors = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    
    // 必須項目のチェック
    if (!data[i][1]) { // 商品名
      errors.push(`行${row}: 商品名が入力されていません`);
    }
    
    if (!data[i][2] || isNaN(data[i][2])) { // 仕入れ価格
      errors.push(`行${row}: 仕入れ価格が正しく入力されていません`);
    }
    
    if (!data[i][3]) { // 仕入れ日
      errors.push(`行${row}: 仕入れ日が入力されていません`);
    }
    
    if (!data[i][4]) { // 仕入れ先
      errors.push(`行${row}: 仕入れ先が入力されていません`);
    }
    
    if (!data[i][5]) { // 仕入れ先のユーザー名/店舗名
      errors.push(`行${row}: 仕入れ先のユーザー名/店舗名が入力されていません`);
    }
    
    // 仕入れ先区分と還付率の整合性チェック
    const supplierType = data[i][14];
    const refundRate = data[i][15];
    
    if (supplierType === '個人' && refundRate !== 8) {
      errors.push(`行${row}: 個人仕入れの還付率が8%になっていません`);
    }
    
    if (supplierType === '商店' && refundRate !== 10) {
      errors.push(`行${row}: 商店仕入れの還付率が10%になっていません`);
    }
  }
  
  if (errors.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('データ検証エラー', errors.join('\n'), ui.ButtonSet.OK);
  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert('データ検証完了', 'すべてのデータが正常です。', ui.ButtonSet.OK);
  }
  
  return errors;
}

/**
 * 在庫管理シートの仕入れ先区分と還付率を自動更新
 */
function updateSupplierTypeAndRefundRate() {
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  const data = inventorySheet.getDataRange().getValues();
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const supplier = data[i][4]; // 仕入れ先
    const currentSupplierType = data[i][14]; // 現在の仕入れ先区分
    const currentRefundRate = data[i][15]; // 現在の還付率
    
    const newSupplierType = determineSupplierType(supplier);
    const newRefundRate = newSupplierType === '個人' ? 8 : 10;
    
    // 変更がある場合のみ更新
    if (currentSupplierType !== newSupplierType || currentRefundRate !== newRefundRate) {
      inventorySheet.getRange(i + 1, 15).setValue(newSupplierType); // 仕入れ先区分
      inventorySheet.getRange(i + 1, 16).setValue(newRefundRate); // 還付率
      updatedCount++;
    }
  }
  
  console.log(`${updatedCount}件の仕入れ先区分と還付率を更新しました。`);
}

/**
 * 在庫管理シートに一括で商品を追加（CSV形式）
 */
function bulkAddInventoryItems() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '一括商品追加',
    'CSV形式で商品情報を入力してください（商品名,仕入れ価格,仕入れ日,仕入れ先,仕入れ先ユーザー名,商品状態,備考,リサーチフラグ,リサーチ担当者）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const csvContent = result.getResponseText();
    processBulkInventoryCSV(csvContent);
  }
}

/**
 * 一括商品追加のCSVデータを処理
 */
function processBulkInventoryCSV(csvContent) {
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  const lines = csvContent.split('\n');
  let newRowIndex = inventorySheet.getLastRow() + 1;
  let addedCount = 0;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    
    const columns = line.split(',');
    if (columns.length >= 9) {
      const productName = columns[0].trim();
      const purchasePrice = parseFloat(columns[1]) || 0;
      const purchaseDate = columns[2].trim();
      const supplier = columns[3].trim();
      const supplierName = columns[4].trim();
      const condition = columns[5].trim();
      const notes = columns[6].trim();
      const researchFlag = columns[7].trim().toUpperCase();
      const researchStaff = columns[8].trim();
      
      if (productName && purchasePrice > 0) {
        const productId = generateProductId(inventorySheet);
        const supplierType = determineSupplierType(supplier);
        const refundRate = supplierType === '個人' ? 8 : 10;
        
        const rowData = [
          productId, // 商品ID
          productName, // 商品名
          purchasePrice, // 仕入れ価格
          new Date(purchaseDate), // 仕入れ日
          supplier, // 仕入れ先
          supplierName, // 仕入れ先のユーザー名/店舗名
          condition, // 商品の状態
          notes, // 備考
          '在庫', // 販売状況
          '', // 販売日
          '', // 販売価格
          '', // 利益
          researchFlag === 'Y' ? 'Y' : 'N', // リサーチ商品フラグ
          researchStaff, // リサーチ担当者
          supplierType, // 仕入れ先区分
          refundRate // 還付率
        ];
        
        inventorySheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
        newRowIndex++;
        addedCount++;
      }
    }
  }
  
  console.log(`${addedCount}件の商品を一括追加しました。`);
}

/**
 * 在庫管理シートの自動採番を修正
 */
function fixProductIds() {
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  const data = inventorySheet.getDataRange().getValues();
  let fixedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const currentId = data[i][0];
    
    // 商品IDが空または不正な形式の場合
    if (!currentId || typeof currentId !== 'string' || !currentId.startsWith('PROD')) {
      const newId = generateProductId(inventorySheet);
      inventorySheet.getRange(i + 1, 1).setValue(newId);
      fixedCount++;
    }
  }
  
  console.log(`${fixedCount}件の商品IDを修正しました。`);
}

/**
 * カスタムメニューに在庫管理機能を追加
 */
function addInventoryMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // 既存のメニューを取得
  const menu = ui.createMenu('在庫管理');
  
  menu
    .addItem('新商品追加', 'addNewInventoryItem')
    .addItem('一括商品追加', 'bulkAddInventoryItems')
    .addSeparator()
    .addItem('送料CSVインポート', 'importShippingCSV')
    .addSeparator()
    .addItem('データ検証', 'validateInventoryData')
    .addItem('仕入れ先区分・還付率自動更新', 'updateSupplierTypeAndRefundRate')
    .addItem('商品ID修正', 'fixProductIds')
    .addToUi();
}

/**
 * メニューを更新（onOpen関数で呼び出し）
 */
function updateMenus() {
  addInventoryMenu();
} 