/**
 * Mazrica → Google Sheets 同期用 Apps Script
 * 
 * このスクリプトをスプレッドシートの「拡張機能」→「Apps Script」に貼り付けてください。
 * デプロイ後、WebアプリのURLをGitHub Secretsの APPS_SCRIPT_URL に設定します。
 */

// 設定
const CONFIG = {
  SHEET_NAME: '案件一覧',  // データを書き込むシート名
  SECRET_KEY: '',  // オプション: セキュリティ用のキー（空の場合は認証なし）
  ID_COLUMN: 1  // 案件IDの列番号（1始まり）
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
      case 'sync':
        return syncData(data);
      case 'get_existing_ids':
        return getExistingIds(data);
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
 * 既存の案件IDリストを取得
 * @param {Object} data - リクエストデータ
 * @returns {TextOutput} レスポンス
 */
function getExistingIds(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const idColumn = data.id_column || CONFIG.ID_COLUMN;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return createResponseWithData(true, 'Sheet not found, will be created', { ids: [] });
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    // ヘッダー行のみ、またはデータなし
    return createResponseWithData(true, 'No data in sheet', { ids: [] });
  }
  
  // 2行目から最終行までのIDを取得（1行目はヘッダー）
  const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
  const idValues = idRange.getValues();
  
  // IDをフラットな配列に変換（数値または文字列）
  const ids = idValues.map(row => row[0]).filter(id => id !== '' && id !== null);
  
  return createResponseWithData(true, 'Found ' + ids.length + ' existing IDs', { ids: ids });
}

/**
 * 差分同期を実行
 * @param {Object} data - リクエストデータ
 * @returns {TextOutput} レスポンス
 */
function syncData(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const headers = data.headers;
  const newRows = data.new_rows || [];  // 新規追加する行
  const deleteIds = data.delete_ids || [];  // 削除する案件ID
  const idColumn = data.id_column || CONFIG.ID_COLUMN;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダーを書き込み
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      sheet.setFrozenRows(1);
    }
  }
  
  let deletedCount = 0;
  let addedCount = 0;
  
  // 削除処理（下から上に削除して行番号のずれを防ぐ）
  if (deleteIds.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
      const idValues = idRange.getValues();
      
      // 削除対象のIDをSetに変換（高速検索のため）
      const deleteIdSet = new Set(deleteIds.map(id => String(id)));
      
      // 削除対象の行番号を収集（下から上の順）
      const rowsToDelete = [];
      for (let i = idValues.length - 1; i >= 0; i--) {
        const cellId = String(idValues[i][0]);
        if (deleteIdSet.has(cellId)) {
          rowsToDelete.push(i + 2);  // 行番号は2始まり（1行目はヘッダー）
        }
      }
      
      // 行を削除
      for (const rowNum of rowsToDelete) {
        sheet.deleteRow(rowNum);
        deletedCount++;
      }
    }
  }
  
  // 新規行を追加
  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    const startRow = lastRow + 1;
    
    sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    addedCount = newRows.length;
  }
  
  // 列幅を自動調整
  if (addedCount > 0) {
    const colCount = headers ? headers.length : (newRows[0] ? newRows[0].length : 0);
    for (let i = 1; i <= colCount; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  return createResponse(true, 'Sync completed: added ' + addedCount + ' rows, deleted ' + deletedCount + ' rows');
}

/**
 * データをシートに書き込む（全件置換）
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
 * データ付きレスポンスを作成する
 * @param {boolean} success - 成功かどうか
 * @param {string} message - メッセージ
 * @param {Object} data - 追加データ
 * @returns {TextOutput} レスポンス
 */
function createResponseWithData(success, message, data) {
  const response = {
    success: success,
    message: message,
    timestamp: new Date().toISOString(),
    ...data
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * テスト用関数
 * Apps Scriptエディタから実行してテストできます
 */
function testSyncData() {
  const testData = {
    action: 'sync',
    sheet_name: 'テスト',
    headers: ['ID', '名前', '金額'],
    new_rows: [
      [4, 'テスト案件D', 40000],
      [5, 'テスト案件E', 50000]
    ],
    delete_ids: [1, 2]
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

function testGetExistingIds() {
  const testData = {
    action: 'get_existing_ids',
    sheet_name: 'テスト'
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}
