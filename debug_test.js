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