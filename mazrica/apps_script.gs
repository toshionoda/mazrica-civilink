/**
 * Mazrica → Google Sheets 同期用 Apps Script
 * 
 * このスクリプトをスプレッドシートの「拡張機能」→「Apps Script」に貼り付けてください。
 * デプロイ後、WebアプリのURLをGitHub Secretsの APPS_SCRIPT_URL に設定します。
 */

// 設定
const CONFIG = {
  SHEET_NAME: '案件一覧',  // データを書き込むシート名
  SECRET_KEY: ''  // オプション: セキュリティ用のキー（空の場合は認証なし）
};

/**
 * POSTリクエストを処理する関数
 * @param {Object} e - リクエストイベント
 * @returns {TextOutput} レスポンス
 */
function doPost(e) {
  try {
    // リクエストボディをパース
    const data = JSON.parse(e.postData.contents);
    
    // シークレットキーの検証（設定されている場合）
    if (CONFIG.SECRET_KEY && data.secret_key !== CONFIG.SECRET_KEY) {
      return createResponse(false, 'Unauthorized');
    }
    
    // アクションに応じて処理
    const action = data.action || 'write';
    
    switch (action) {
      case 'write':
        return writeData(data);
      case 'clear':
        return clearSheet(data);
      case 'ping':
        return createResponse(true, 'pong');
      default:
        return createResponse(false, 'Unknown action: ' + action);
    }
    
  } catch (error) {
    return createResponse(false, 'Error: ' + error.message);
  }
}

/**
 * GETリクエストを処理する関数（ヘルスチェック用）
 * @param {Object} e - リクエストイベント
 * @returns {TextOutput} レスポンス
 */
function doGet(e) {
  return createResponse(true, 'Mazrica Sync API is ready');
}

/**
 * データをシートに書き込む
 * @param {Object} data - リクエストデータ
 * @returns {TextOutput} レスポンス
 */
function writeData(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const rows = data.rows;
  const headers = data.headers;
  const clearBefore = data.clear_before !== false;  // デフォルトはtrue
  
  if (!rows || !Array.isArray(rows)) {
    return createResponse(false, 'Invalid data format: rows is required');
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // 既存データをクリア
  if (clearBefore) {
    sheet.clear();
  }
  
  // データを書き込み
  const allData = headers ? [headers, ...rows] : rows;
  
  if (allData.length > 0 && allData[0].length > 0) {
    const range = sheet.getRange(1, 1, allData.length, allData[0].length);
    range.setValues(allData);
    
    // ヘッダー行の書式設定
    if (headers) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      sheet.setFrozenRows(1);
    }
    
    // 列幅を自動調整
    for (let i = 1; i <= allData[0].length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  return createResponse(true, 'Written ' + rows.length + ' rows to ' + sheetName);
}

/**
 * シートをクリアする
 * @param {Object} data - リクエストデータ
 * @returns {TextOutput} レスポンス
 */
function clearSheet(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return createResponse(false, 'Sheet not found: ' + sheetName);
  }
  
  sheet.clear();
  return createResponse(true, 'Cleared sheet: ' + sheetName);
}

/**
 * レスポンスを作成する
 * @param {boolean} success - 成功かどうか
 * @param {string} message - メッセージ
 * @returns {TextOutput} レスポンス
 */
function createResponse(success, message) {
  const response = {
    success: success,
    message: message,
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * テスト用関数
 * Apps Scriptエディタから実行してテストできます
 */
function testWriteData() {
  const testData = {
    action: 'write',
    sheet_name: 'テスト',
    headers: ['ID', '名前', '金額'],
    rows: [
      [1, 'テスト案件A', 10000],
      [2, 'テスト案件B', 20000],
      [3, 'テスト案件C', 30000]
    ]
  };
  
  // テスト用のモックイベント
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}

