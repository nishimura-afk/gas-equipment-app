/**
 * 見積りDB参照サービス
 * 外部の見積りDBスプレッドシートを読み取り専用で参照
 */

var ESTIMATE_DB_ID = '17FKM50xNHEcftYmZ2u1O2VrAtPKnKm5E5dG1tnIR6fo';
var ESTIMATE_SHEET_NAME = '見積りDB';

/**
 * 見積りDBから設備カテゴリで検索（部分一致）
 * 案件化している設備名・作業内容とDBの「設備カテゴリ」を部分一致で照合
 * @param {string} category - 設備名・作業内容（例: ガソリン計量機, 灯油計量機, 塗装, エコステージ）
 * @return {Array} 該当する見積り一覧
 */
function searchEstimatesByCategory(category) {
  try {
    if (!category || String(category).trim() === '') {
      return [];
    }
    var ss = SpreadsheetApp.openById(ESTIMATE_DB_ID);
    var sheet = ss.getSheetByName(ESTIMATE_SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var data = sheet.getDataRange().getValues();
    var results = [];
    var cat = String(category).trim();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dbCategory = String(row[3] || '').trim();  // 設備カテゴリ列
      // 部分一致: DBのカテゴリに選択値が含まれる、または選択値にDBのカテゴリが含まれる
      var match = dbCategory.indexOf(cat) >= 0 || cat.indexOf(dbCategory) >= 0;
      if (match) {
        results.push({
          registeredAt: row[0],
          estimateDate: row[1],
          location: row[2],
          category: row[3],
          vendor: row[4],
          amount: row[5],
          fileName: row[6],
          pdfUrl: row[7],
          folder: row[8]
        });
      }
    }

    return results;
  } catch (e) {
    Logger.log('見積りDB参照エラー: ' + e.message);
    return [];
  }
}

/**
 * 見積り比較データを取得（案件管理画面用）
 * 同じ設備カテゴリの全店舗見積りを取得
 * @param {string} category - 設備カテゴリ
 * @return {Object} 比較用データ
 */
function getEstimateComparison(category) {
  var estimates = searchEstimatesByCategory(category);

  // 金額でソート（安い順）
  estimates.sort(function(a, b) {
    return (a.amount || 0) - (b.amount || 0);
  });

  // 統計情報を計算
  var amounts = estimates.filter(function(e) { return e.amount > 0; })
                       .map(function(e) { return e.amount; });

  var stats = {
    count: amounts.length,
    min: amounts.length > 0 ? Math.min.apply(null, amounts) : 0,
    max: amounts.length > 0 ? Math.max.apply(null, amounts) : 0,
    average: amounts.length > 0 ? amounts.reduce(function(a, b) { return a + b; }, 0) / amounts.length : 0
  };

  return {
    category: category,
    estimates: estimates,
    stats: stats
  };
}

/**
 * Web API: 見積り比較データ取得
 */
function doGetEstimateComparison(e) {
  var category = e.parameter.category;
  var data = getEstimateComparison(category);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
