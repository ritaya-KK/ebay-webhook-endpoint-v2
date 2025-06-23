/**
 * eBay API連携機能
 * 注文データの取得、返金データの取得、利益計算の自動化
 */

// eBay API設定
const EBAY_API_BASE_URL = 'https://api.ebay.com';
const EBAY_SANDBOX_BASE_URL = 'https://api.sandbox.ebay.com';

/**
 * eBay API認証トークンを取得
 */
function getEbayAuthToken() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
  const apiKey = settingsSheet.getRange('B2').getValue();
  const apiSecret = settingsSheet.getRange('B3').getValue();
  
  if (!apiKey || !apiSecret) {
    throw new Error('eBay API KeyまたはSecretが設定されていません。設定シートで確認してください。');
  }
  
  const credentials = Utilities.base64Encode(apiKey + ':' + apiSecret);
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: 'grant_type=client_credentials&scope=https://api.ebay.com/oauth/api_scope'
  };
  
  try {
    const response = UrlFetchApp.fetch(EBAY_API_BASE_URL + '/identity/v1/oauth2/token', options);
    const result = JSON.parse(response.getContentText());
    return result.access_token;
  } catch (error) {
    console.error('eBay API認証エラー: ' + error.toString());
    throw error;
  }
}

/**
 * eBay注文データを取得
 */
function fetchEbayOrders() {
  const authToken = getEbayAuthToken();
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('販売管理シート（利益計算）');
  
  // 最後に取得した日時を取得（設定シートから）
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
  let lastFetchDate = settingsSheet.getRange('B10').getValue(); // 最後に取得した日時
  
  if (!lastFetchDate) {
    // 初回実行時は30日前から取得
    lastFetchDate = new Date();
    lastFetchDate.setDate(lastFetchDate.getDate() - 30);
  }
  
  const currentDate = new Date();
  const formattedLastDate = Utilities.formatDate(lastFetchDate, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  const formattedCurrentDate = Utilities.formatDate(currentDate, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + authToken,
      'Content-Type': 'application/json'
    }
  };
  
  try {
    // 注文データを取得
    const ordersUrl = `${EBAY_API_BASE_URL}/sell/fulfillment/v1/order?filter=creationdate:[${formattedLastDate}..${formattedCurrentDate}]&limit=100`;
    const ordersResponse = UrlFetchApp.fetch(ordersUrl, options);
    const ordersData = JSON.parse(ordersResponse.getContentText());
    
    if (ordersData.orders && ordersData.orders.length > 0) {
      processOrders(ordersData.orders, salesSheet);
    }
    
    // 最後に取得した日時を更新
    settingsSheet.getRange('B10').setValue(currentDate);
    
    console.log(`${ordersData.orders ? ordersData.orders.length : 0}件の注文データを取得しました。`);
    
  } catch (error) {
    console.error('eBay注文データ取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 注文データを処理して販売管理シートに記録
 */
function processOrders(orders, salesSheet) {
  const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('在庫管理シート（古物台帳）');
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
  
  // 設定値を取得
  const exchangeRate = parseFloat(settingsSheet.getRange('B4').getValue()) || 150;
  const personalRefundRate = parseFloat(settingsSheet.getRange('B7').getValue()) || 8;
  const storeRefundRate = parseFloat(settingsSheet.getRange('B8').getValue()) || 10;
  
  let newRowIndex = salesSheet.getLastRow() + 1;
  
  orders.forEach(order => {
    order.lineItems.forEach(lineItem => {
      // 商品IDで在庫管理シートから情報を取得
      const productInfo = findProductByEbayItemId(lineItem.itemId, inventorySheet);
      
      if (productInfo) {
        const rowData = [
          order.orderId, // 販売ID
          productInfo.productId, // 商品ID
          lineItem.title, // 商品名
          new Date(order.creationDate), // 販売日
          parseFloat(lineItem.total.value) * exchangeRate, // 販売価格（円換算）
          calculateEbayFee(parseFloat(lineItem.total.value)), // eBay手数料
          calculatePaypalFee(parseFloat(lineItem.total.value)), // PayPal手数料
          0, // 送料（後で更新）
          0, // その他経費
          productInfo.purchasePrice, // 仕入れ価格
          0, // 一時販売利益（計算式で自動計算）
          0, // 一時利益率（計算式で自動計算）
          order.buyer.username, // 顧客情報
          order.fulfillmentStartInstructions[0]?.shippingStep?.shipTo?.contactAddress?.city || '', // 配送状況
          0, // 返金金額
          '', // 返金理由
          '', // 返金日
          0, // 最終利益（返金後）（計算式で自動計算）
          productInfo.researchFlag, // リサーチ商品フラグ
          productInfo.supplierType, // 仕入れ先区分
          productInfo.refundRate, // 還付率
          0, // 還付金額（計算式で自動計算）
          0 // 還付後最終利益（計算式で自動計算）
        ];
        
        salesSheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
        newRowIndex++;
        
        // 在庫管理シートの販売状況を更新
        updateInventorySalesStatus(productInfo.productId, inventorySheet, order.orderId, parseFloat(lineItem.total.value) * exchangeRate);
      }
    });
  });
  
  // 計算式を設定
  setCalculationFormulas(salesSheet, newRowIndex - 1);
}

/**
 * eBay商品IDで在庫管理シートから商品情報を検索
 */
function findProductByEbayItemId(ebayItemId, inventorySheet) {
  const data = inventorySheet.getDataRange().getValues();
  
  // 商品名にeBay商品IDが含まれているかチェック
  for (let i = 1; i < data.length; i++) {
    const productName = data[i][1]; // 商品名
    const purchasePrice = data[i][2]; // 仕入れ価格
    const supplierType = data[i][14]; // 仕入れ先区分
    const researchFlag = data[i][12]; // リサーチ商品フラグ
    
    if (productName && productName.includes(ebayItemId.toString())) {
      const refundRate = supplierType === '個人' ? 8 : 10;
      
      return {
        productId: data[i][0], // 商品ID
        purchasePrice: purchasePrice,
        supplierType: supplierType,
        refundRate: refundRate,
        researchFlag: researchFlag
      };
    }
  }
  
  return null;
}

/**
 * 在庫管理シートの販売状況を更新
 */
function updateInventorySalesStatus(productId, inventorySheet, orderId, salesPrice) {
  const data = inventorySheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      inventorySheet.getRange(i + 1, 9).setValue('販売済み'); // 販売状況
      inventorySheet.getRange(i + 1, 10).setValue(new Date()); // 販売日
      inventorySheet.getRange(i + 1, 11).setValue(salesPrice); // 販売価格
      break;
    }
  }
}

/**
 * eBay手数料を計算
 */
function calculateEbayFee(salesAmount) {
  // eBay手数料率（概算）
  const feeRate = 0.10; // 10%
  return salesAmount * feeRate;
}

/**
 * PayPal手数料を計算
 */
function calculatePaypalFee(salesAmount) {
  // PayPal手数料率（概算）
  const feeRate = 0.029; // 2.9%
  const fixedFee = 0.30; // $0.30
  return salesAmount * feeRate + fixedFee;
}

/**
 * 販売管理シートに計算式を設定
 */
function setCalculationFormulas(salesSheet, lastRow) {
  if (lastRow < 2) return;
  
  // 一時販売利益の計算式
  const tempProfitFormula = '=E{0}-J{0}-F{0}-G{0}-H{0}-I{0}';
  // 一時利益率の計算式
  const tempProfitRateFormula = '=IF(E{0}>0,K{0}/E{0},0)';
  // 最終利益（返金後）の計算式
  const finalProfitFormula = '=K{0}-O{0}';
  // 還付金額の計算式
  const refundAmountFormula = '=J{0}*U{0}/100';
  // 還付後最終利益の計算式
  const refundedFinalProfitFormula = '=R{0}+V{0}';
  
  for (let row = 2; row <= lastRow; row++) {
    salesSheet.getRange(row, 11).setFormula(tempProfitFormula.replace(/\{0\}/g, row)); // 一時販売利益
    salesSheet.getRange(row, 12).setFormula(tempProfitRateFormula.replace(/\{0\}/g, row)); // 一時利益率
    salesSheet.getRange(row, 18).setFormula(finalProfitFormula.replace(/\{0\}/g, row)); // 最終利益（返金後）
    salesSheet.getRange(row, 22).setFormula(refundAmountFormula.replace(/\{0\}/g, row)); // 還付金額
    salesSheet.getRange(row, 23).setFormula(refundedFinalProfitFormula.replace(/\{0\}/g, row)); // 還付後最終利益
  }
}

/**
 * 返金データを取得
 */
function fetchEbayRefunds() {
  const authToken = getEbayAuthToken();
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('販売管理シート（利益計算）');
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + authToken,
      'Content-Type': 'application/json'
    }
  };
  
  try {
    // 返金データを取得（過去30日分）
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const formattedDate = Utilities.formatDate(thirtyDaysAgo, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
    
    const refundsUrl = `${EBAY_API_BASE_URL}/sell/fulfillment/v1/order?filter=lastmodifieddate:[${formattedDate}]&limit=100`;
    const refundsResponse = UrlFetchApp.fetch(refundsUrl, options);
    const refundsData = JSON.parse(refundsResponse.getContentText());
    
    if (refundsData.orders && refundsData.orders.length > 0) {
      processRefunds(refundsData.orders, salesSheet);
    }
    
    console.log(`${refundsData.orders ? refundsData.orders.length : 0}件の返金データを取得しました。`);
    
  } catch (error) {
    console.error('eBay返金データ取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 返金データを処理
 */
function processRefunds(orders, salesSheet) {
  const data = salesSheet.getDataRange().getValues();
  
  orders.forEach(order => {
    if (order.pricingSummary && order.pricingSummary.adjustments) {
      order.pricingSummary.adjustments.forEach(adjustment => {
        if (adjustment.type === 'REFUND') {
          // 販売管理シートから該当する注文を検索
          for (let i = 1; i < data.length; i++) {
            if (data[i][0] === order.orderId) { // 販売IDが一致
              const refundAmount = parseFloat(adjustment.amount.value);
              const refundReason = adjustment.reason || '返金';
              const refundDate = new Date(adjustment.date);
              
              salesSheet.getRange(i + 1, 15).setValue(refundAmount); // 返金金額
              salesSheet.getRange(i + 1, 16).setValue(refundReason); // 返金理由
              salesSheet.getRange(i + 1, 17).setValue(refundDate); // 返金日
              break;
            }
          }
        }
      });
    }
  });
}

/**
 * 月次レポートを生成
 */
function generateMonthlyReport() {
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('販売管理シート（利益計算）');
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
  const fixedCostSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定費管理シート');
  
  const currentMonth = new Date().getMonth() + 1;
  const currentYear = new Date().getFullYear();
  
  // 販売データを取得
  const salesData = salesSheet.getDataRange().getValues();
  let totalSales = 0;
  let totalTempProfit = 0;
  let totalRefundedProfit = 0;
  let totalRefundAmount = 0;
  let researchProfit = 0;
  
  for (let i = 1; i < salesData.length; i++) {
    const salesDate = new Date(salesData[i][3]); // 販売日
    if (salesDate.getMonth() + 1 === currentMonth && salesDate.getFullYear() === currentYear) {
      totalSales += salesData[i][4] || 0; // 販売価格
      totalTempProfit += salesData[i][10] || 0; // 一時販売利益
      totalRefundedProfit += salesData[i][22] || 0; // 還付後最終利益
      totalRefundAmount += salesData[i][14] || 0; // 返金金額
      
      // リサーチ商品の利益を集計
      if (salesData[i][18] === 'Y') { // リサーチ商品フラグ
        researchProfit += salesData[i][22] || 0; // 還付後最終利益
      }
    }
  }
  
  // 固定費を取得
  const fixedCostData = fixedCostSheet.getDataRange().getValues();
  let totalFixedCost = 0;
  let ebayStoreFee = 0;
  
  for (let i = 1; i < fixedCostData.length; i++) {
    const costMonth = new Date(fixedCostData[i][3]).getMonth() + 1; // 発生月
    if (costMonth === currentMonth) {
      totalFixedCost += fixedCostData[i][2] || 0; // 金額
      if (fixedCostData[i][6] === 'Y') { // eBayストア料金フラグ
        ebayStoreFee += fixedCostData[i][2] || 0;
      }
    }
  }
  
  // 月次レポートシートを作成または更新
  let reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('月次レポート');
  if (!reportSheet) {
    reportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('月次レポート');
  }
  
  // レポートヘッダーを設定
  const reportHeaders = [
    ['月次レポート', `${currentYear}年${currentMonth}月`],
    ['項目', '金額（円）'],
    ['総販売額', totalSales],
    ['一時販売利益', totalTempProfit],
    ['返金金額', totalRefundAmount],
    ['還付後最終利益', totalRefundedProfit],
    ['固定費合計', totalFixedCost],
    ['eBayストア料金', ebayStoreFee],
    ['純利益', totalRefundedProfit - totalFixedCost],
    ['リサーチ商品利益', researchProfit],
    ['リサーチ利益配分（5%）', researchProfit * 0.05]
  ];
  
  reportSheet.clear();
  reportSheet.getRange(1, 1, reportHeaders.length, 2).setValues(reportHeaders);
  
  // ヘッダーのスタイルを設定
  reportSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  reportSheet.getRange(2, 1, 1, 2).setFontWeight('bold').setBackground('#f1f3f4');
  
  console.log(`${currentYear}年${currentMonth}月の月次レポートを生成しました。`);
}

/**
 * カスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('eBay管理')
    .addItem('eBay注文データ取得', 'fetchEbayOrders')
    .addItem('eBay返金データ取得', 'fetchEbayRefunds')
    .addItem('月次レポート生成', 'generateMonthlyReport')
    .addSeparator()
    .addItem('全データ更新', 'updateAllData')
    .addToUi();
}

/**
 * 全データを更新
 */
function updateAllData() {
  try {
    console.log('eBay注文データを取得中...');
    fetchEbayOrders();
    
    console.log('eBay返金データを取得中...');
    fetchEbayRefunds();
    
    console.log('月次レポートを生成中...');
    generateMonthlyReport();
    
    console.log('全データの更新が完了しました！');
  } catch (error) {
    console.error('データ更新中にエラーが発生しました: ' + error.toString());
  }
} 