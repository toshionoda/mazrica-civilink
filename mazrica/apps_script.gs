/**
 * Mazrica → Google Sheets 同期用 Apps Script
 *
 * このスクリプトをスプレッドシートの「拡張機能」→「Apps Script」に貼り付けてください。
 * デプロイ後、WebアプリのURLをGitHub Secretsの APPS_SCRIPT_URL に設定します。
 *
 * 機能:
 * - sync / write / get_existing_ids / clear / ping アクション対応
 * - 列18(R)に「エリア」、列19(S)に「契約方法」を案件名(B列)から自動抽出・上書き
 * - 新規案件を `顧客リスト` シートに自動同期
 * - `spreadsheet_id` パラメータ指定で他スプレッドシートにも書き込み可（openById切替）
 * - 既存シートでもヘッダー行を常時最新化（シェイプ変更に追従）
 */

// 設定
const CONFIG = {
  SHEET_NAME: '案件一覧_v2',  // データを書き込むシート名
  SECRET_KEY: '',             // オプション: セキュリティ用のキー（空の場合は認証なし）
  ID_COLUMN: 1,               // 案件IDの列番号（1始まり）
  B_COLUMN: 3,                // 抽出元の列番号（C列=案件名 / 25列版HEADERSでは3列目）
  AREA_COLUMN: 18,            // エリア出力先（R列）
  CONTRACT_TYPE_COLUMN: 19,   // 契約方法出力先（S列）

  // 顧客リスト連携設定
  CUSTOMER_LIST_SHEET: '顧客リスト',
  CUSTOMER_LIST_MAPPING: {
    mazricaIdColumn: 2,      // B列: mazricaID
    companyNameColumn: 1,    // A列: 会社名
    freeUserColumn: 13,      // M列: 無償ユーザー
    areaColumn: 3            // C列: エリア
  },
  MAZRICA_COLUMNS: {
    id: 1,          // A列: ID
    companyName: 2, // B列: 取引先名（25列版HEADERSでは2列目）
    user: 16,       // P列: 無料トライアル開始日 → ユーザー項目は未使用
    area: 18        // R列: エリア
  }
};

// 契約方法の種類を定義
const CONTRACT_TYPES = ['無料トライアル', 'サブスクリプション', '有償化クローズ'];

/**
 * spreadsheet_id が指定されていれば openById、なければ active スプレッドシートを返す
 */
function _getSpreadsheet(data) {
  return (data && data.spreadsheet_id)
    ? SpreadsheetApp.openById(data.spreadsheet_id)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * B列の値からエリアと契約方法を抽出する
 */
function extractFieldsFromB(bValue) {
  if (!bValue) return { area: '', contractType: '' };

  const parts = String(bValue).split('_');

  // エリア: 2番目の要素（インデックス1）
  const area = parts.length > 1 ? parts[1] : '';

  // 契約方法: 定義済みキーワードに一致する要素
  const contractType = parts.find(p => CONTRACT_TYPES.includes(p)) || '';

  return { area, contractType };
}

/**
 * 顧客リストシートから既存のmazricaIDを全取得
 */
function getCustomerListIds(ss) {
  const sheet = ss.getSheetByName(CONFIG.CUSTOMER_LIST_SHEET);

  if (!sheet) {
    return new Set();
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Set();
  }

  const idColumn = CONFIG.CUSTOMER_LIST_MAPPING.mazricaIdColumn;
  const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
  const idValues = idRange.getValues();

  const ids = new Set();
  idValues.forEach(row => {
    if (row[0] !== '' && row[0] !== null) {
      ids.add(String(row[0]));
    }
  });

  return ids;
}

/**
 * 顧客リストシートに新規行を追加
 */
function addToCustomerList(ss, rows) {
  if (!rows || rows.length === 0) {
    return 0;
  }

  let sheet = ss.getSheetByName(CONFIG.CUSTOMER_LIST_SHEET);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CUSTOMER_LIST_SHEET);
    const headers = [];
    headers[CONFIG.CUSTOMER_LIST_MAPPING.companyNameColumn - 1] = '会社名';
    headers[CONFIG.CUSTOMER_LIST_MAPPING.mazricaIdColumn - 1] = 'mazricaID';
    headers[CONFIG.CUSTOMER_LIST_MAPPING.areaColumn - 1] = 'エリア';
    headers[CONFIG.CUSTOMER_LIST_MAPPING.freeUserColumn - 1] = '無償ユーザー';

    const maxCol = Math.max(
      CONFIG.CUSTOMER_LIST_MAPPING.companyNameColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.mazricaIdColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.areaColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.freeUserColumn
    );
    while (headers.length < maxCol) {
      headers.push('');
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e0e0e0');
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;

  const customerRows = rows.map(mazricaRow => {
    const customerRow = [];
    const maxCol = Math.max(
      CONFIG.CUSTOMER_LIST_MAPPING.companyNameColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.mazricaIdColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.areaColumn,
      CONFIG.CUSTOMER_LIST_MAPPING.freeUserColumn
    );

    for (let i = 0; i < maxCol; i++) {
      customerRow.push('');
    }

    // mazrica A列(ID) → 顧客リスト B列(mazricaID)
    customerRow[CONFIG.CUSTOMER_LIST_MAPPING.mazricaIdColumn - 1] =
      mazricaRow[CONFIG.MAZRICA_COLUMNS.id - 1] || '';

    // mazrica B列(取引先名) → 顧客リスト A列(会社名)
    customerRow[CONFIG.CUSTOMER_LIST_MAPPING.companyNameColumn - 1] =
      mazricaRow[CONFIG.MAZRICA_COLUMNS.companyName - 1] || '';

    // mazrica P列(ユーザー=現行は無料トライアル開始日) → 顧客リスト M列(無償ユーザー)
    customerRow[CONFIG.CUSTOMER_LIST_MAPPING.freeUserColumn - 1] =
      mazricaRow[CONFIG.MAZRICA_COLUMNS.user - 1] || '';

    // mazrica R列(エリア) → 顧客リスト C列(エリア)
    customerRow[CONFIG.CUSTOMER_LIST_MAPPING.areaColumn - 1] =
      mazricaRow[CONFIG.MAZRICA_COLUMNS.area - 1] || '';

    return customerRow;
  });

  if (customerRows.length > 0) {
    sheet.getRange(startRow, 1, customerRows.length, customerRows[0].length).setValues(customerRows);
  }

  return customerRows.length;
}

/**
 * mazricaの新規行を顧客リストに同期
 */
function syncToCustomerList(ss, newRows) {
  if (!newRows || newRows.length === 0) {
    return { synced: 0, skipped: 0 };
  }

  const existingIds = getCustomerListIds(ss);

  const rowsToAdd = [];
  let skipped = 0;

  newRows.forEach(row => {
    const mazricaId = String(row[CONFIG.MAZRICA_COLUMNS.id - 1] || '');

    if (mazricaId && !existingIds.has(mazricaId)) {
      rowsToAdd.push(row);
    } else {
      skipped++;
    }
  });

  const synced = addToCustomerList(ss, rowsToAdd);

  return { synced, skipped };
}

/**
 * 行データにエリアと契約方法を追加する
 */
function addExtractedFields(row, bColumnIndex, areaColumnIndex, contractTypeColumnIndex) {
  const maxColumn = Math.max(areaColumnIndex, contractTypeColumnIndex) + 1;
  while (row.length < maxColumn) {
    row.push('');
  }

  const bValue = row[bColumnIndex] || '';
  const extracted = extractFieldsFromB(bValue);

  row[areaColumnIndex] = extracted.area;
  row[contractTypeColumnIndex] = extracted.contractType;

  return row;
}

/**
 * POSTリクエストを処理する関数
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (CONFIG.SECRET_KEY && data.secret_key !== CONFIG.SECRET_KEY) {
      return createResponse(false, 'Unauthorized');
    }

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
 */
function doGet(e) {
  return createResponse(true, 'Mazrica Sync API is ready');
}

/**
 * 既存の案件IDリストを取得
 */
function getExistingIds(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const idColumn = data.id_column || CONFIG.ID_COLUMN;

  const ss = _getSpreadsheet(data);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return createResponseWithData(true, 'Sheet not found, will be created', { ids: [] });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return createResponseWithData(true, 'No data in sheet', { ids: [] });
  }

  const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
  const idValues = idRange.getValues();

  const ids = idValues.map(row => row[0]).filter(id => id !== '' && id !== null);

  return createResponseWithData(true, 'Found ' + ids.length + ' existing IDs', { ids: ids });
}

/**
 * 差分同期を実行
 */
function syncData(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const headers = data.headers;
  const newRows = data.new_rows || [];
  const updateRows = data.update_rows || [];
  const deleteIds = data.delete_ids || [];
  const idColumn = data.id_column || CONFIG.ID_COLUMN;

  const bColumnIndex = (data.b_column || CONFIG.B_COLUMN) - 1;
  const areaColumnIndex = (data.area_column || CONFIG.AREA_COLUMN) - 1;
  const contractTypeColumnIndex = (data.contract_type_column || CONFIG.CONTRACT_TYPE_COLUMN) - 1;

  // 顧客リスト同期を skip するフラグ（デフォルトは実施）
  const syncCustomerList = data.sync_customer_list !== false;

  const ss = _getSpreadsheet(data);
  let sheet = ss.getSheetByName(sheetName);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // ヘッダーを常に最新に保つ（シェイプ変更に追従）
  if (headers && headers.length > 0) {
    const extendedHeaders = [...headers];
    while (extendedHeaders.length < areaColumnIndex + 1) {
      extendedHeaders.push('');
    }
    extendedHeaders[areaColumnIndex] = 'エリア';
    while (extendedHeaders.length < contractTypeColumnIndex + 1) {
      extendedHeaders.push('');
    }
    extendedHeaders[contractTypeColumnIndex] = '契約方法';

    const currentLastCol = sheet.getLastColumn();
    const currentHeaders = currentLastCol > 0
      ? sheet.getRange(1, 1, 1, currentLastCol).getValues()[0]
      : [];
    const needsUpdate = JSON.stringify(currentHeaders) !== JSON.stringify(extendedHeaders);
    if (needsUpdate) {
      const clearCols = Math.max(currentLastCol, extendedHeaders.length);
      if (clearCols > 0) {
        sheet.getRange(1, 1, 1, clearCols).clearContent();
      }
      sheet.getRange(1, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);
      const headerRange = sheet.getRange(1, 1, 1, extendedHeaders.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      sheet.setFrozenRows(1);
    }
  }

  let deletedCount = 0;
  let addedCount = 0;
  let updatedCount = 0;

  // 削除処理
  if (deleteIds.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
      const idValues = idRange.getValues();

      const deleteIdSet = new Set(deleteIds.map(id => String(id)));

      const rowsToDelete = [];
      for (let i = idValues.length - 1; i >= 0; i--) {
        const cellId = String(idValues[i][0]);
        if (deleteIdSet.has(cellId)) {
          rowsToDelete.push(i + 2);
        }
      }

      for (const rowNum of rowsToDelete) {
        sheet.deleteRow(rowNum);
        deletedCount++;
      }
    }
  }

  // 既存行の上書き処理（A〜Y の同期列のみ、Z列以降は触らない）
  // - update_rows を案件ID で既存行に突合
  // - addExtractedFields でエリア(R)/契約方法(S)を案件名から再算出
  // - 既存データの該当範囲を読み出してメモリ上で差し替え、1回の setValues で書き戻し
  if (updateRows.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 同期列数: ヘッダー長を基準にしつつ R/S 列まで必ず含める
      const headerLen = headers ? headers.length : 0;
      const syncColCount = Math.max(
        headerLen,
        contractTypeColumnIndex + 1,
        areaColumnIndex + 1
      );

      const existingRange = sheet.getRange(2, 1, lastRow - 1, syncColCount);
      const existingData = existingRange.getValues();

      const idToIdx = new Map();
      for (let i = 0; i < existingData.length; i++) {
        const cellId = String(existingData[i][idColumn - 1]);
        if (cellId) idToIdx.set(cellId, i);
      }

      for (const row of updateRows) {
        const id = String(row[idColumn - 1] || '');
        if (!id) continue;
        const idx = idToIdx.get(id);
        if (idx === undefined) continue;

        const processedRow = addExtractedFields(
          [...row], bColumnIndex, areaColumnIndex, contractTypeColumnIndex
        );

        // 同期列数に正規化（短ければ '' で埋め、長ければ切り捨て）
        const fixedRow = new Array(syncColCount);
        for (let c = 0; c < syncColCount; c++) {
          fixedRow[c] = c < processedRow.length ? processedRow[c] : '';
        }
        existingData[idx] = fixedRow;
        updatedCount++;
      }

      if (updatedCount > 0) {
        sheet.getRange(2, 1, existingData.length, syncColCount).setValues(existingData);
      }
    }
  }

  // 新規行を追加（B列から抽出してR/S列に追加）
  let customerListResult = { synced: 0, skipped: 0 };

  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    const startRow = lastRow + 1;

    const processedRows = newRows.map(row => {
      return addExtractedFields([...row], bColumnIndex, areaColumnIndex, contractTypeColumnIndex);
    });

    const colCount = processedRows[0].length;
    sheet.getRange(startRow, 1, processedRows.length, colCount).setValues(processedRows);
    addedCount = processedRows.length;

    if (syncCustomerList) {
      customerListResult = syncToCustomerList(ss, processedRows);
    }
  }

  // 列幅を自動調整
  if (addedCount > 0) {
    const colCount = Math.max(
      headers ? headers.length : 0,
      newRows[0] ? newRows[0].length : 0,
      contractTypeColumnIndex + 1
    );
    for (let i = 1; i <= colCount; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  return createResponse(true,
    'Sync completed: added ' + addedCount + ' rows, updated ' + updatedCount + ' rows, deleted ' + deletedCount + ' rows. ' +
    'Customer list: added ' + customerListResult.synced + ', skipped ' + customerListResult.skipped
  );
}

/**
 * データをシートに書き込む（全件置換）
 */
function writeData(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;
  const rows = data.rows;
  const headers = data.headers;
  const clearBefore = data.clear_before !== false;

  const bColumnIndex = (data.b_column || CONFIG.B_COLUMN) - 1;
  const areaColumnIndex = (data.area_column || CONFIG.AREA_COLUMN) - 1;
  const contractTypeColumnIndex = (data.contract_type_column || CONFIG.CONTRACT_TYPE_COLUMN) - 1;

  const syncCustomerList = data.sync_customer_list !== false;

  if (!rows || !Array.isArray(rows)) {
    return createResponse(false, 'Invalid data format: rows is required');
  }

  const ss = _getSpreadsheet(data);
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (clearBefore) {
    sheet.clear();
  }

  let extendedHeaders = null;
  if (headers) {
    extendedHeaders = [...headers];
    while (extendedHeaders.length < areaColumnIndex + 1) {
      extendedHeaders.push('');
    }
    extendedHeaders[areaColumnIndex] = 'エリア';
    while (extendedHeaders.length < contractTypeColumnIndex + 1) {
      extendedHeaders.push('');
    }
    extendedHeaders[contractTypeColumnIndex] = '契約方法';
  }

  const processedRows = rows.map(row => {
    return addExtractedFields([...row], bColumnIndex, areaColumnIndex, contractTypeColumnIndex);
  });

  const allData = extendedHeaders ? [extendedHeaders, ...processedRows] : processedRows;

  if (allData.length > 0 && allData[0].length > 0) {
    const range = sheet.getRange(1, 1, allData.length, allData[0].length);
    range.setValues(allData);

    if (extendedHeaders) {
      const headerRange = sheet.getRange(1, 1, 1, extendedHeaders.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e0e0e0');
      sheet.setFrozenRows(1);
    }

    for (let i = 1; i <= allData[0].length; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  let customerListResult = { synced: 0, skipped: 0 };
  if (syncCustomerList) {
    customerListResult = syncToCustomerList(ss, processedRows);
  }

  return createResponse(true,
    'Written ' + rows.length + ' rows to ' + sheetName + '. ' +
    'Customer list: added ' + customerListResult.synced + ', skipped ' + customerListResult.skipped
  );
}

/**
 * シートをクリアする
 */
function clearSheet(data) {
  const sheetName = data.sheet_name || CONFIG.SHEET_NAME;

  const ss = _getSpreadsheet(data);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return createResponse(false, 'Sheet not found: ' + sheetName);
  }

  sheet.clear();
  return createResponse(true, 'Cleared sheet: ' + sheetName);
}

/**
 * レスポンスを作成する
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
 * テスト用関数（Apps Scriptエディタから実行可能）
 */
function testPing() {
  const result = doPost({ postData: { contents: JSON.stringify({ action: 'ping' }) } });
  Logger.log(result.getContent());
}
