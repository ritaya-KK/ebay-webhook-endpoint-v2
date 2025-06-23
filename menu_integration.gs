/**
 * メニュー統合機能
 * すべての機能を統合したカスタムメニューを作成
 */

/**
 * スプレッドシートを開いた時に実行される関数
 * カスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // eBay管理メニュー
  const ebayMenu = ui.createMenu('eBay管理');
  ebayMenu
    .addItem('eBay注文データ取得', 'fetchEbayOrders')
    .addItem('eBay返金データ取得', 'fetchEbayRefunds')
    .addItem('月次レポート生成', 'generateMonthlyReport')
    .addSeparator()
    .addItem('全データ更新', 'updateAllData')
    .addToUi();
  
  // 在庫管理メニュー
  const inventoryMenu = ui.createMenu('在庫管理');
  inventoryMenu
    .addItem('新商品追加', 'addNewInventoryItem')
    .addItem('一括商品追加', 'bulkAddInventoryItems')
    .addSeparator()
    .addItem('送料CSVインポート', 'importShippingCSV')
    .addSeparator()
    .addItem('データ検証', 'validateInventoryData')
    .addItem('仕入れ先区分・還付率自動更新', 'updateSupplierTypeAndRefundRate')
    .addItem('商品ID修正', 'fixProductIds')
    .addToUi();
  
  // システム管理メニュー
  const systemMenu = ui.createMenu('システム管理');
  systemMenu
    .addItem('スプレッドシート初期化', 'initializeSpreadsheet')
    .addItem('設定値リセット', 'resetSettings')
    .addItem('データバックアップ', 'backupData')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * スプレッドシートの初期化
 */
function initializeSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'スプレッドシート初期化',
    'スプレッドシートを初期化しますか？\n既存のデータは削除されます。',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // 既存のシートを削除
      const sheets = spreadsheet.getSheets();
      sheets.forEach(sheet => {
        spreadsheet.deleteSheet(sheet);
      });
      
      // 新しいシートを作成
      createInventorySheet(spreadsheet);
      createSalesSheet(spreadsheet);
      createFixedCostSheet(spreadsheet);
      createOutsourcingSheet(spreadsheet);
      createShippingSheet(spreadsheet);
      createSettingsSheet(spreadsheet);
      
      ui.alert('初期化完了', 'スプレッドシートの初期化が完了しました。', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('エラー', '初期化中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * 設定値のリセット
 */
function resetSettings() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '設定値リセット',
    '設定値をデフォルトにリセットしますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    try {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
      
      // デフォルト設定値を設定
      const defaultSettings = [
        ['設定項目', '設定値', '説明'],
        ['eBay API Key', '', 'eBay APIキー'],
        ['eBay API Secret', '', 'eBay APIシークレット'],
        ['為替レート（USD/JPY）', '150', '現在の為替レート'],
        ['利益率目標', '30%', '目標利益率'],
        ['リサーチ利益配分率', '5%', '外注スタッフへの配分率'],
        ['個人仕入れ還付率', '8%', '個人仕入れの還付率'],
        ['商店仕入れ還付率', '10%', '商店仕入れの還付率'],
        ['eBayストア料金', '5000', '月額ストア料金（円）'],
        ['最後に取得した日時', '', '最後にeBayデータを取得した日時']
      ];
      
      settingsSheet.clear();
      settingsSheet.getRange(1, 1, defaultSettings.length, 3).setValues(defaultSettings);
      
      // ヘッダーのスタイルを設定
      const headerRange = settingsSheet.getRange(1, 1, 1, 3);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#666666');
      headerRange.setFontColor('white');
      
      ui.alert('リセット完了', '設定値のリセットが完了しました。', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('エラー', 'リセット中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * データのバックアップ
 */
function backupData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const backupName = `eBay管理システム_バックアップ_${new Date().toISOString().split('T')[0]}`;
    
    // バックアップ用のスプレッドシートを作成
    const backupSpreadsheet = SpreadsheetApp.create(backupName);
    
    // 各シートをコピー
    const sheets = spreadsheet.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const backupSheet = backupSpreadsheet.insertSheet(sheetName);
      
      // データをコピー
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      
      // 列幅をコピー
      const columnCount = sheet.getLastColumn();
      for (let i = 1; i <= columnCount; i++) {
        backupSheet.setColumnWidth(i, sheet.getColumnWidth(i));
      }
    });
    
    // デフォルトシートを削除
    const defaultSheet = backupSpreadsheet.getSheetByName('シート1');
    if (defaultSheet) {
      backupSpreadsheet.deleteSheet(defaultSheet);
    }
    
    const backupUrl = backupSpreadsheet.getUrl();
    ui.alert(
      'バックアップ完了',
      `バックアップが完了しました。\nURL: ${backupUrl}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('エラー', 'バックアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ヘルプの表示
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  
  const helpText = `
eBay販売管理・古物台帳システム ヘルプ

【基本操作】
1. 新商品追加: 在庫管理シートに新しい商品を追加
2. 一括商品追加: CSV形式で複数の商品を一括追加
3. eBay注文データ取得: 最新の注文データを自動取得
4. 送料CSVインポート: 送料データをCSVからインポート

【利益計算】
- 一時販売利益 = 販売価格 - 仕入れ価格 - 手数料 - 送料
- 還付金額 = 仕入れ価格 × 還付率（個人8%/商店10%）
- 還付後最終利益 = 一時販売利益 + 還付金額 - 返金金額

【データ検証】
- データ検証: 入力データの整合性をチェック
- 仕入れ先区分・還付率自動更新: 仕入れ先に基づいて自動更新

【レポート】
- 月次レポート生成: 当月の売上・利益・費用の集計

【設定】
- 設定シートでAPI Key、為替レート、還付率などを設定
- 定期的に為替レートを更新してください

詳細な使用方法はREADME.mdを参照してください。
  `;
  
  ui.alert('ヘルプ', helpText, ui.ButtonSet.OK);
}

/**
 * システムの状態をチェック
 */
function checkSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const status = [];
  
  try {
    // 必要なシートの存在確認
    const requiredSheets = [
      '在庫管理シート（古物台帳）',
      '販売管理シート（利益計算）',
      '固定費管理シート',
      '外注費管理シート',
      '送料管理シート',
      '設定シート'
    ];
    
    requiredSheets.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        status.push(`✓ ${sheetName}: 正常`);
      } else {
        status.push(`✗ ${sheetName}: 見つかりません`);
      }
    });
    
    // 設定値の確認
    const settingsSheet = spreadsheet.getSheetByName('設定シート');
    if (settingsSheet) {
      const apiKey = settingsSheet.getRange('B2').getValue();
      const apiSecret = settingsSheet.getRange('B3').getValue();
      
      if (apiKey && apiSecret) {
        status.push('✓ eBay API設定: 完了');
      } else {
        status.push('✗ eBay API設定: 未設定');
      }
      
      const exchangeRate = settingsSheet.getRange('B4').getValue();
      if (exchangeRate) {
        status.push(`✓ 為替レート: ${exchangeRate}`);
      } else {
        status.push('✗ 為替レート: 未設定');
      }
    }
    
    // データ件数の確認
    const inventorySheet = spreadsheet.getSheetByName('在庫管理シート（古物台帳）');
    if (inventorySheet) {
      const inventoryCount = inventorySheet.getLastRow() - 1;
      status.push(`✓ 在庫商品数: ${inventoryCount}件`);
    }
    
    const salesSheet = spreadsheet.getSheetByName('販売管理シート（利益計算）');
    if (salesSheet) {
      const salesCount = salesSheet.getLastRow() - 1;
      status.push(`✓ 販売記録数: ${salesCount}件`);
    }
    
  } catch (error) {
    status.push(`✗ エラー: ${error.toString()}`);
  }
  
  const statusText = status.join('\n');
  ui.alert('システム状態', statusText, ui.ButtonSet.OK);
}

/**
 * クイックスタートガイド
 */
function showQuickStartGuide() {
  const ui = SpreadsheetApp.getUi();
  
  const guideText = `
【クイックスタートガイド】

Step 1: 初期設定
1. 設定シートでeBay API KeyとSecretを入力
2. 為替レートを現在のレートに更新

Step 2: 商品データの入力
1. 在庫管理メニュー→新商品追加
2. 商品情報を入力（仕入れ価格、仕入れ先など）
3. リサーチ商品の場合は担当者も入力

Step 3: eBayデータの取得
1. eBay管理メニュー→eBay注文データ取得
2. 最新の注文データが自動で取得される
3. 利益計算が自動で実行される

Step 4: 月次レポートの確認
1. eBay管理メニュー→月次レポート生成
2. 当月の売上・利益・費用を確認

【重要なポイント】
- 仕入れ先は正確に入力（メルカリ/ヤフオク→個人、Amazon→商店）
- 為替レートは定期的に更新
- データ検証機能で入力データをチェック
  `;
  
  ui.alert('クイックスタートガイド', guideText, ui.ButtonSet.OK);
}

/**
 * メニューに追加の機能を追加
 */
function addAdditionalMenus() {
  const ui = SpreadsheetApp.getUi();
  
  // システム管理メニューに追加機能を追加
  const systemMenu = ui.createMenu('システム管理');
  systemMenu
    .addItem('スプレッドシート初期化', 'initializeSpreadsheet')
    .addItem('設定値リセット', 'resetSettings')
    .addItem('データバックアップ', 'backupData')
    .addSeparator()
    .addItem('システム状態チェック', 'checkSystemStatus')
    .addItem('クイックスタートガイド', 'showQuickStartGuide')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * 完全なメニュー設定（onOpen関数の代替）
 */
function setupCompleteMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // eBay管理メニュー
  const ebayMenu = ui.createMenu('eBay管理');
  ebayMenu
    .addItem('eBay注文データ取得', 'fetchEbayOrders')
    .addItem('eBay返金データ取得', 'fetchEbayRefunds')
    .addItem('月次レポート生成', 'generateMonthlyReport')
    .addSeparator()
    .addItem('全データ更新', 'updateAllData')
    .addToUi();
  
  // 在庫管理メニュー
  const inventoryMenu = ui.createMenu('在庫管理');
  inventoryMenu
    .addItem('新商品追加', 'addNewInventoryItem')
    .addItem('一括商品追加', 'bulkAddInventoryItems')
    .addSeparator()
    .addItem('送料CSVインポート', 'importShippingCSV')
    .addSeparator()
    .addItem('データ検証', 'validateInventoryData')
    .addItem('仕入れ先区分・還付率自動更新', 'updateSupplierTypeAndRefundRate')
    .addItem('商品ID修正', 'fixProductIds')
    .addToUi();
  
  // システム管理メニュー
  const systemMenu = ui.createMenu('システム管理');
  systemMenu
    .addItem('スプレッドシート初期化', 'initializeSpreadsheet')
    .addItem('設定値リセット', 'resetSettings')
    .addItem('データバックアップ', 'backupData')
    .addSeparator()
    .addItem('システム状態チェック', 'checkSystemStatus')
    .addItem('クイックスタートガイド', 'showQuickStartGuide')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
} 