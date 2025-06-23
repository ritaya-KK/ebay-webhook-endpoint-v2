/**
 * eBay Webhook Handler
 * Marketplace Account Deletion通知を受け取るためのエンドポイント
 */

/**
 * 設定シートから指定された設定値を取得
 * @param {string} settingName - 取得したい設定の名称
 * @return {string|null} - 設定値。見つからない場合はnull。
 */
function getSetting(settingName) {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定シート');
  if (!settingsSheet) return null;
  const data = settingsSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1]; // B列の値を返す
    }
  }
  return null;
}

/**
 * eBayがエンドポイントの所有権を確認するために送信するGETリクエストを処理
 * @param {object} e - GETリクエストのイベントオブジェクト
 * @return {ContentService.TextOutput} - 検証用のレスポンス
 */
function doGet(e) {
  try {
    const challengeCode = e.parameter.challenge_code;
    logNotification('Debug', `1. doGetが開始されました。`);

    if (challengeCode) {
      logNotification('Debug', `2. challengeCodeを受信しました: ${challengeCode}`);
      
      // 設定シートからVerification Tokenを取得
      const verificationToken = getSetting('eBay Verification Token');
      if (!verificationToken) {
          logNotification("Verification Error", "Verification tokenが設定シートで見つかりません。");
          return ContentService.createTextOutput("Token not configured").setMimeType(ContentService.MimeType.TEXT);
      }
      logNotification('Debug', `3. verificationTokenを取得しました: ${verificationToken}`);
      
      // エンドポイントURLを取得
      const endpointUrl = ScriptApp.getService().getUrl();
      logNotification('Debug', `4. エンドポイントURLを取得しました: ${endpointUrl}`);
      
      // eBayの仕様に従い、チャレンジコード、トークン、エンドポイントURLを連結
      const stringToHash = challengeCode + verificationToken + endpointUrl;
      logNotification('Debug', `5. ハッシュ化する文字列: ${stringToHash}`);
      
      // SHA256でハッシュ化
      const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, stringToHash, Utilities.Charset.UTF_8);
      const hashString = hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
      logNotification('Debug', `6. 計算されたハッシュ値: ${hashString}`);

      // レスポンスを作成
      const responsePayload = {
        challengeResponse: hashString
      };

      logNotification('Verification Handshake', `検証リクエストに応答します。`);
      // レスポンスをJSON形式で返す
      return ContentService.createTextOutput(JSON.stringify(responsePayload))
                           .setMimeType(ContentService.MimeType.JSON);
    } else {
      // 通常のGETリクエストへの応答
      logNotification('Info', 'challenge_codeなしでエンドポイントが直接アクセスされました。');
      return ContentService.createTextOutput("Webhook endpoint is active.");
    }
  } catch (err) {
      logNotification('doGet Error', `エラー: ${err.toString()}, スタック: ${err.stack}`);
      return ContentService.createTextOutput("Error processing request.").setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * eBayからの通知（POSTリクエスト）を処理
 * @param {object} e - POSTリクエストのイベントオブジェクト
 */
function doPost(e) {
  try {
    // ★★★ セキュリティ警告 ★★★
    // Google Apps Scriptの標準機能では、eBayが使用するECDSA署名の検証が困難です。
    // そのため、この実装では署名検証をスキップしています。
    // 通知は記録されますが、厳密なセキュリティが求められる場合は別の環境をご検討ください。
    
    // リクエストボディ（通知内容）を取得
    const notification = JSON.parse(e.postData.contents);
    const notificationType = notification.metadata.topic;

    // 通知をログシートに記録
    logNotification(notificationType, e.postData.contents);

  } catch (err) {
    logNotification('doPost Error', err.toString());
  }
}

/**
 * 通知ログシートに情報を記録
 * @param {string} type - 通知のタイプ
 * @param {string} payload - 通知の内容
 */
function logNotification(type, payload) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('通知ログ');
    if (logSheet) {
      logSheet.appendRow([new Date(), type, payload]);
    }
  } catch (err) {
    // ログ記録のエラーは無視
  }
} 