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

/**
 * ダッシュボードデータの詳細診断
 */
function diagnoseDashboardData() {
  Logger.log('========================================');
  Logger.log('ダッシュボードデータ診断');
  Logger.log('========================================');
  
  const dashData = getDashboardData();
  
  Logger.log('\n【総数】');
  Logger.log('アラート総数: ' + dashData.noticeCount);
  Logger.log('正常設備数: ' + dashData.normalCount);
  
  Logger.log('\n【カテゴリ別集計】');
  const categoryCount = {};
  dashData.noticeList.forEach(item => {
    const cat = item['カテゴリ'] || '不明';
    categoryCount[cat] = (categoryCount[cat] || 0) + 1;
  });
  
  for (const cat in categoryCount) {
    Logger.log(`  ${cat}: ${categoryCount[cat]}件`);
  }
  
  Logger.log('\n【ガソリン計量機関連】');
  const gasRelated = dashData.noticeList.filter(item => 
    item['設備名'] && (
      item['設備名'].includes('ガソリン計量機') ||
      item['設備名'].includes('計量機')
    )
  );
  
  Logger.log(`ガソリン計量機関連: ${gasRelated.length}件`);
  gasRelated.forEach(item => {
    Logger.log(`\n  ${item['拠点名']} - ${item['設備名']}`);
    Logger.log(`    設備ID: ${item['設備ID']}`);
    Logger.log(`    カテゴリ: ${item['カテゴリ']}`);
    Logger.log(`    本体: ${item['本体ステータス']}`);
    Logger.log(`    部品A: ${item['部品Aステータス']}`);
    Logger.log(`    部品B: ${item['部品Bステータス'] || '(空)'}`);
  });
  
  Logger.log('\n【部材更新カテゴリ】');
  const partsMaintenance = dashData.noticeList.filter(item => 
    item['カテゴリ'] === '部材更新'
  );
  
  Logger.log(`部材更新: ${partsMaintenance.length}件`);
  if (partsMaintenance.length > 0) {
    Logger.log('\n⚠️ 以下の部材更新がnoticeListに含まれています:');
    partsMaintenance.slice(0, 10).forEach(item => {
      Logger.log(`  ${item['拠点名']} - ${item['設備名']}`);
    });
  }
  
  Logger.log('\n========================================');
}

/**
 * ステータス集計シートの実データを確認
 */
function checkActualStatusSheet() {
  Logger.log('========================================');
  Logger.log('ステータス集計シート実データ確認');
  Logger.log('========================================');
  
  const config = getConfig();
  const statusSheet = getSheet(config.SHEET_NAMES.STATUS_SUMMARY);
  const data = statusSheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log('\n【ヘッダー】');
  headers.forEach((h, i) => {
    Logger.log(`列${i}: ${h}`);
  });
  
  Logger.log('\n【PARTS-PUMP-4Y の設備】');
  let found = 0;
  
  for (let i = 1; i < data.length; i++) {
    const eqId = data[i][headers.indexOf('設備ID')];
    
    if (eqId === 'PARTS-PUMP-4Y') {
      found++;
      const locName = data[i][headers.indexOf('拠点名')];
      const category = data[i][headers.indexOf('カテゴリ')];
      const bodyStatus = data[i][headers.indexOf('本体ステータス')];
      
      Logger.log(`\n${found}. ${locName}`);
      Logger.log(`   カテゴリ: "${category}"`);
      Logger.log(`   本体ステータス: "${bodyStatus}"`);
      
      if (found >= 5) break; // 最初の5件のみ
    }
  }
  
  Logger.log('\n========================================');
}

function debugPARTS_PUMP_4Y() {
  const config = getConfig();
  
  // Configの設定を確認
  Logger.log('=== Config設定 ===');
  const cycle = config.MAINTENANCE_CYCLES['PARTS_PUMP_4Y'];
  Logger.log('カテゴリ: ' + cycle.category);
  Logger.log('年数: ' + cycle.years);
  Logger.log('searchKey: ' + cycle.searchKey);
  Logger.log('suffix: ' + cycle.suffix);
  
  // 実際にマッチングできるか確認
  Logger.log('\n=== マッチングテスト ===');
  const testEqId = 'PARTS-PUMP-4Y';
  const testEqName = 'ガソリン計量機部品(4年)';
  
  const matched = findCycleByEquipmentId(testEqId, testEqName, config.MAINTENANCE_CYCLES);
  if (matched) {
    Logger.log('✅ マッチング成功');
    Logger.log('マッチしたカテゴリ: ' + matched.category);
  } else {
    Logger.log('❌ マッチング失敗');
  }
}

/**
 * フォルダIDとファイル検出のデバッグ
 */
function debugFolderAccess() {
  Logger.log('=== フォルダアクセステスト ===');
  
  const folderId = '1gUD2Z2N2-APYFXQcugu1fdex_YfoiyA2';
  Logger.log('テスト対象フォルダID: ' + folderId);
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('✅ フォルダにアクセス成功');
    Logger.log('フォルダ名: ' + folder.getName());
    
    // すべてのファイルを取得
    const allFiles = folder.getFiles();
    let fileCount = 0;
    
    Logger.log('\n【全ファイル一覧】');
    while (allFiles.hasNext()) {
      fileCount++;
      const file = allFiles.next();
      Logger.log(fileCount + '. ' + file.getName());
      Logger.log('   MimeType: ' + file.getMimeType());
      Logger.log('   ID: ' + file.getId());
    }
    
    if (fileCount === 0) {
      Logger.log('⚠️ フォルダは空です');
    } else {
      Logger.log('\n合計ファイル数: ' + fileCount);
    }
    
    // PDFのみ取得
    Logger.log('\n【PDFファイルのみ】');
    const pdfFiles = folder.getFilesByType(MimeType.PDF);
    let pdfCount = 0;
    
    while (pdfFiles.hasNext()) {
      pdfCount++;
      const pdf = pdfFiles.next();
      Logger.log(pdfCount + '. ' + pdf.getName());
    }
    
    Logger.log('\nPDFファイル数: ' + pdfCount);
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
  }
  
  Logger.log('\n=== テスト完了 ===');
}

/**
 * 抽出データの構造確認
 * （見積り機能は別システムに移管したため、このテストは無効化）
 */
function debugExtractedData() {
  Logger.log('=== 見積りPDF抽出機能は設備管理システムから削除されました ===');
  Logger.log('見積りの分類・管理は別システム（分類リネームGAS）で行います。');
}