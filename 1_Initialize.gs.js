/**
 * 1_Initialize.gs
 * スプレッドシートの初期化関数
 */

/**
 * 見積管理マスタシートを初期化
 */
function initEstimateMasterSheet() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(config.SHEET_NAMES.ESTIMATE_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(config.SHEET_NAMES.ESTIMATE_MASTER);
  }
  
  // ヘッダー行
  const headers = [
    '見積ID',
    '登録日',
    '案件ID',
    '拠点コード',
    '拠点名',
    '設備ID',
    '設備名',
    '業者名',
    '見積日',
    '総額(税抜)',
    '消費税',
    '総額(税込)',
    'メモ',
    'PDFファイル名',
    'PDFリンク'
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // スタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);  // 見積ID
  sheet.setColumnWidth(2, 100);  // 登録日
  sheet.setColumnWidth(3, 120);  // 案件ID
  sheet.setColumnWidth(4, 80);   // 拠点コード
  sheet.setColumnWidth(5, 120);  // 拠点名
  sheet.setColumnWidth(6, 100);  // 設備ID
  sheet.setColumnWidth(7, 150);  // 設備名
  sheet.setColumnWidth(8, 120);  // 業者名
  sheet.setColumnWidth(9, 100);  // 見積日
  sheet.setColumnWidth(10, 100); // 総額(税抜)
  sheet.setColumnWidth(11, 100); // 消費税
  sheet.setColumnWidth(12, 100); // 総額(税込)
  sheet.setColumnWidth(13, 200); // メモ
  sheet.setColumnWidth(14, 250); // PDFファイル名
  sheet.setColumnWidth(15, 80);  // PDFリンク
  
  // 金額列に通貨書式
  sheet.getRange(2, 10, sheet.getMaxRows() - 1, 3)
    .setNumberFormat('#,##0');
  
  // 日付列に日付書式
  sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1)
    .setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1)
    .setNumberFormat('yyyy-mm-dd');
  
  // PDFリンク列に数式（ハイパーリンク）
  // 注意: 数式は動的に設定するため、ここでは設定しない
  // 代わりに、saveEstimateLink関数内でHYPERLINK関数を使用
  
  // ウィンドウ枠の固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ 見積管理マスタシート初期化完了');
  return sheet;
}

/**
 * 案件別見積比較シートを初期化（オプション）
 */
function initEstimateComparisonSheet() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(config.SHEET_NAMES.ESTIMATE_COMPARISON);
  if (!sheet) {
    sheet = ss.insertSheet(config.SHEET_NAMES.ESTIMATE_COMPARISON);
  }
  
  const headers = [
    '案件ID',
    '拠点名',
    '設備名',
    '作業内容',
    '業者A',
    '見積額A',
    'PDFリンクA',
    '業者B',
    '見積額B',
    'PDFリンクB',
    '業者C',
    '見積額C',
    'PDFリンクC',
    '選定業者',
    '選定理由'
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#2ECC71');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  
  Logger.log('✅ 案件別見積比較シート初期化完了');
  return sheet;
}

/**
 * すべての見積関連シートを初期化
 */
function initAllEstimateSheets() {
  try {
    initEstimateMasterSheet();
    initEstimateComparisonSheet();
    
    SpreadsheetApp.getUi().alert('✅ 見積管理シートの初期化が完了しました');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.message);
    Logger.log('Error: ' + e.stack);
  }
}
