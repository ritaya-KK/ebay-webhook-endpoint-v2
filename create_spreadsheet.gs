/**
 * eBay管理用スプレッドシート作成スクリプト
 * 在庫管理、販売管理、固定費管理、外注費管理、送料管理、設定の6つのシートを作成
 */

function createEbayManagementSpreadsheet() {
  // スプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create('eBay販売管理・古物台帳システム');
  const spreadsheetId = spreadsheet.getId();
  
  console.log('スプレッドシートが作成されました。ID: ' + spreadsheetId);
  
  // 既存のシートを削除（デフォルトのシート1を除く）
  const sheets = spreadsheet.getSheets();
  if (sheets.length > 1) {
    for (let i = 1; i < sheets.length; i++) {
      spreadsheet.deleteSheet(sheets[i]);
    }
  }
  
  // シート1の名前を変更
  const sheet1 = spreadsheet.getSheets()[0];
  sheet1.setName('在庫管理シート（古物台帳）');
  
  // 各シートを作成
  createInventorySheet(spreadsheet);
  createSalesSheet(spreadsheet);
  createFixedCostSheet(spreadsheet);
  createOutsourcingSheet(spreadsheet);
  createShippingSheet(spreadsheet);
  createSettingsSheet(spreadsheet);
  createNotificationLogSheet(spreadsheet);
  
  // スプレッドシートのURLを取得
  const url = spreadsheet.getUrl();
  console.log('スプレッドシートのURL: ' + url);
  
  return {
    spreadsheetId: spreadsheetId,
    url: url
  };
}

/**
 * 在庫管理シート（古物台帳）を作成
 */
function createInventorySheet(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('在庫管理シート（古物台帳）');
  
  // ヘッダーを設定
  const headers = [
    '商品ID',
    '商品名',
    '仕入れ価格',
    '仕入れ日',
    '仕入れ先',
    '仕入れ先のユーザー名/店舗名',
    '商品の状態',
    '備考',
    '販売状況',
    '販売日',
    '販売価格',
    '利益',
    'リサーチ商品フラグ',
    'リサーチ担当者',
    '仕入れ先区分',
    '還付率'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 80);  // 商品ID
  sheet.setColumnWidth(2, 200); // 商品名
  sheet.setColumnWidth(3, 100); // 仕入れ価格
  sheet.setColumnWidth(4, 100); // 仕入れ日
  sheet.setColumnWidth(5, 120); // 仕入れ先
  sheet.setColumnWidth(6, 150); // 仕入れ先のユーザー名/店舗名
  sheet.setColumnWidth(7, 100); // 商品の状態
  sheet.setColumnWidth(8, 150); // 備考
  sheet.setColumnWidth(9, 100); // 販売状況
  sheet.setColumnWidth(10, 100); // 販売日
  sheet.setColumnWidth(11, 100); // 販売価格
  sheet.setColumnWidth(12, 100); // 利益
  sheet.setColumnWidth(13, 120); // リサーチ商品フラグ
  sheet.setColumnWidth(14, 120); // リサーチ担当者
  sheet.setColumnWidth(15, 120); // 仕入れ先区分
  sheet.setColumnWidth(16, 80);  // 還付率
  
  // データ検証を設定
  // 仕入れ先区分のドロップダウン
  const supplierTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['個人', '商店'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O2:O1000').setDataValidation(supplierTypeValidation);
  
  // リサーチ商品フラグのドロップダウン
  const researchFlagValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('M2:M1000').setDataValidation(researchFlagValidation);
  
  // 販売状況のドロップダウン
  const salesStatusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['在庫', '販売済み'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(salesStatusValidation);
}

/**
 * 販売管理シート（利益計算）を作成
 */
function createSalesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('販売管理シート（利益計算）');
  
  // ヘッダーを設定
  const headers = [
    '販売ID',
    '商品ID',
    '商品名',
    '販売日',
    '販売価格',
    'eBay手数料',
    'PayPal手数料',
    '送料',
    'その他経費',
    '仕入れ価格',
    '一時販売利益',
    '一時利益率',
    '顧客情報',
    '配送状況',
    '返金金額',
    '返金理由',
    '返金日',
    '最終利益（返金後）',
    'リサーチ商品フラグ',
    '仕入れ先区分',
    '還付率',
    '還付金額',
    '還付後最終利益'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 120); // 販売ID
  sheet.setColumnWidth(2, 80);  // 商品ID
  sheet.setColumnWidth(3, 200); // 商品名
  sheet.setColumnWidth(4, 100); // 販売日
  sheet.setColumnWidth(5, 100); // 販売価格
  sheet.setColumnWidth(6, 100); // eBay手数料
  sheet.setColumnWidth(7, 100); // PayPal手数料
  sheet.setColumnWidth(8, 80);  // 送料
  sheet.setColumnWidth(9, 100); // その他経費
  sheet.setColumnWidth(10, 100); // 仕入れ価格
  sheet.setColumnWidth(11, 120); // 一時販売利益
  sheet.setColumnWidth(12, 100); // 一時利益率
  sheet.setColumnWidth(13, 150); // 顧客情報
  sheet.setColumnWidth(14, 100); // 配送状況
  sheet.setColumnWidth(15, 100); // 返金金額
  sheet.setColumnWidth(16, 150); // 返金理由
  sheet.setColumnWidth(17, 100); // 返金日
  sheet.setColumnWidth(18, 120); // 最終利益（返金後）
  sheet.setColumnWidth(19, 120); // リサーチ商品フラグ
  sheet.setColumnWidth(20, 120); // 仕入れ先区分
  sheet.setColumnWidth(21, 80);  // 還付率
  sheet.setColumnWidth(22, 100); // 還付金額
  sheet.setColumnWidth(23, 120); // 還付後最終利益
}

/**
 * 固定費管理シートを作成
 */
function createFixedCostSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('固定費管理シート');
  
  // ヘッダーを設定
  const headers = [
    '固定費ID',
    '固定費項目',
    '金額',
    '発生月',
    '支払日',
    '備考',
    'eBayストア料金フラグ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 80);  // 固定費ID
  sheet.setColumnWidth(2, 150); // 固定費項目
  sheet.setColumnWidth(3, 100); // 金額
  sheet.setColumnWidth(4, 100); // 発生月
  sheet.setColumnWidth(5, 100); // 支払日
  sheet.setColumnWidth(6, 150); // 備考
  sheet.setColumnWidth(7, 150); // eBayストア料金フラグ
  
  // eBayストア料金フラグのドロップダウン
  const storeFeeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('G2:G1000').setDataValidation(storeFeeValidation);
}

/**
 * 外注費管理シートを作成
 */
function createOutsourcingSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('外注費管理シート');
  
  // ヘッダーを設定
  const headers = [
    '外注ID',
    '外注先',
    '外注内容',
    '外注金額',
    '外注日',
    '支払日',
    '支払方法',
    '備考',
    'リサーチ利益配分'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 80);  // 外注ID
  sheet.setColumnWidth(2, 120); // 外注先
  sheet.setColumnWidth(3, 200); // 外注内容
  sheet.setColumnWidth(4, 100); // 外注金額
  sheet.setColumnWidth(5, 100); // 外注日
  sheet.setColumnWidth(6, 100); // 支払日
  sheet.setColumnWidth(7, 100); // 支払方法
  sheet.setColumnWidth(8, 150); // 備考
  sheet.setColumnWidth(9, 120); // リサーチ利益配分
}

/**
 * 送料管理シートを作成
 */
function createShippingSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('送料管理シート');
  
  // ヘッダーを設定
  const headers = [
    '注文ID',
    '商品ID',
    '商品名',
    '送料',
    '配送方法',
    '配送日',
    '配送状況'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 120); // 注文ID
  sheet.setColumnWidth(2, 80);  // 商品ID
  sheet.setColumnWidth(3, 200); // 商品名
  sheet.setColumnWidth(4, 80);  // 送料
  sheet.setColumnWidth(5, 120); // 配送方法
  sheet.setColumnWidth(6, 100); // 配送日
  sheet.setColumnWidth(7, 100); // 配送状況
}

/**
 * 設定シートを作成
 */
function createSettingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('設定シート');
  
  // ★ 新しい、よりシンプルなトークンに変更
  const verificationToken = 'ryujisebaysystemverificationtoken123456789';

  // 設定項目を設定
  const settings = [
    ['設定項目', '設定値', '説明'],
    ['eBay API Key', '', 'eBay APIキー'],
    ['eBay API Secret', '', 'eBay APIシークレット'],
    ['為替レート（USD/JPY）', '150', '現在の為替レート'],
    ['利益率目標', '30%', '目標利益率'],
    ['リサーチ利益配分率', '5%', '外注スタッフへの配分率'],
    ['個人仕入れ還付率', '8%', '個人仕入れの還付率'],
    ['商店仕入れ還付率', '10%', '商店仕入れの還付率'],
    ['eBayストア料金', '5000', '月額ストア料金（円）'],
    ['eBay Verification Token', verificationToken, 'eBayからの通知を検証するためのトークン。この値をコピーしてeBayの開発者ポータルに貼り付けてください。']
  ];
  
  sheet.getRange(1, 1, settings.length, 3).setValues(settings);
  
  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#666666');
  headerRange.setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 200); // 設定項目
  sheet.setColumnWidth(2, 150); // 設定値
  sheet.setColumnWidth(3, 300); // 説明
  
  // 設定値の列にデータ検証を設定
  // 為替レート
  const exchangeRateValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(100, 200)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B3').setDataValidation(exchangeRateValidation);
  
  // 利益率目標
  const profitRateValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B5').setDataValidation(profitRateValidation);
  
  // リサーチ利益配分率
  const researchRateValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B6').setDataValidation(researchRateValidation);
  
  // 還付率
  const refundRateValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B7').setDataValidation(refundRateValidation);
  sheet.getRange('B8').setDataValidation(refundRateValidation);
  
  // eBayストア料金
  const storeFeeValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100000)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B9').setDataValidation(storeFeeValidation);
}

/**
 * ★ 通知ログシートを作成する関数を追加
 */
function createNotificationLogSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('通知ログ');
  
  const headers = [
    '受信日時',
    '通知タイプ',
    'ペイロード（内容）'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ff9800');
  headerRange.setFontColor('white');
  
  sheet.setColumnWidth(1, 150); // 受信日時
  sheet.setColumnWidth(2, 150); // 通知タイプ
  sheet.setColumnWidth(3, 500); // ペイロード（内容）
}

/**
 * メイン実行関数
 */
function main() {
  try {
    const result = createEbayManagementSpreadsheet();
    console.log('スプレッドシートの作成が完了しました！');
    console.log('スプレッドシートID: ' + result.spreadsheetId);
    console.log('URL: ' + result.url);
  } catch (error) {
    console.error('エラーが発生しました: ' + error.toString());
  }
} 