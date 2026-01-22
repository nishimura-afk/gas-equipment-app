function debugNozzleCover() {
  Logger.log('=== debugNozzleCover 開始 ===');
  var targetStores = getNozzleCoverTargetStores();
  
  Logger.log('対象店舗数: ' + targetStores.length);
  Logger.log('対象店舗一覧:');
  
  for (var i = 0; i < targetStores.length; i++) {
    var store = targetStores[i];
    Logger.log((i + 1) + '. ' + store.name + ' (設置日: ' + store.installDate + ', 初回4月: ' + store.firstApril + ')');
  }
  
  Logger.log('=== getNozzleCoverInfo 確認 ===');
  var info = getNozzleCoverInfo();
  Logger.log('hasAlert: ' + info.hasAlert);
  Logger.log('targetCount: ' + info.targetCount);
  Logger.log('targetYear: ' + info.targetYear);
}

function showNozzleCoverFunction() {
  Logger.log(getNozzleCoverTargetStores.toString());
}

function testNozzleCover() {
  var stores = getNozzleCoverTargetStores();
  Logger.log('対象店舗数: ' + stores.length);
  stores.forEach(function(s) {
    Logger.log(s.code + ' ' + s.name + ' - ' + Utilities.formatDate(s.installDate, 'JST', 'yyyy/MM/dd'));
  });
}

/**
 * デバッグ用：ステータス集計シートのヘッダーを確認
 */
function debugStatusHeaders() {
  const config = getConfig();
  const statusSheet = getSheet(config.SHEET_NAMES.STATUS_SUMMARY);
  const headers = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
  
  Logger.log('=== ステータス集計シートのヘッダー ===');
  headers.forEach((header, index) => {
    Logger.log(`列${index + 1}: ${header}`);
  });
  
  return headers;
}