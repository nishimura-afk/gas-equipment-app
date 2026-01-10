/**
 * 1_Setup.gs v6.6
 * 4月実施一括発注対応（5種類）完全版
 */
function initialSetup() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  
  const sheetDefinitions = [
    { name: config.SHEET_NAMES.MASTER_EQUIPMENT, headers: ['拠点コード', '拠点名', '設備ID', '設備名', '型式・仕様', '設置日(前回実施)', '部品A交換日', '部品B最終交換日', '備考', '次回作業メモ'] },
    { name: config.SHEET_NAMES.MASTER_LOCATION, headers: ['拠点コード', '拠点名', 'オープン日', '担当者名', 'メールアドレス'] },
    { name: config.SHEET_NAMES.SCHEDULE, headers: ['ID', '拠点コード', '設備ID', '作業内容', '予定日', 'ステータス', 'カレンダーID', '発注先'] },
    { name: config.SHEET_NAMES.HISTORY, headers: ['拠点コード', '設備ID', '作業内容', '実施日', '補助金情報', '備考'] },
    { name: config.SHEET_NAMES.STATUS_SUMMARY, headers: ['拠点コード', '拠点名', '設備ID', '設備名', 'カテゴリ', '設置日(前回実施)', '部品Aステータス', '部品Bステータス', '本体ステータス', '部品B対象', 'monthDiffA', 'subsidyAlert', 'nextWorkMemo', 'spec', '次回予定日'] },
    { name: config.SHEET_NAMES.SYS_LOG, headers: ['タイムスタンプ', 'ユーザー', '操作種別', '詳細', 'ステータス'] },
    { name: config.SHEET_NAMES.CONFIG_MASTER, headers: ['設定キー', '分類', '設備名(表示用)', '基準年数', '検索キーワード(取込用)', 'ID接尾辞'] },
    // ★ここに追加
    { name: config.SHEET_NAMES.ESTIMATE_HEADER, headers: ['見積ID', '案件ID', '拠点コード', '拠点名', '設備ID', '設備名', '業者名', '見積日', '総額(税抜)', '消費税', '総額(税込)', '諸経費', 'PDFファイル名', 'PDFリンク', '登録日'] },
    { name: config.SHEET_NAMES.ESTIMATE_DETAIL, headers: ['見積ID', '行番号', '項目名', '単価', '数量', '単位', '小計', '備考'] }
  ];

  for (const def of sheetDefinitions) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) sheet = ss.insertSheet(def.name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#e2e8f0');
    }
  }

  importEquipmentData(ss, config);
  setupSystemTriggers();
  Logger.log('初期セットアップ完了');
}

function updateWebData() {
  try {
    const alertCount = refreshStatusSummaryFast();
    if (alertCount > 0) checkAndSendAlertMail();
    Logger.log('Webデータ更新完了');
  } catch (e) {
    Logger.log('更新エラー: ' + e.message);
  }
}

function refreshStatusSummaryFast() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const summarySheet = ss.getSheetByName(config.SHEET_NAMES.STATUS_SUMMARY);
  
  const masterValues = masterSheet.getDataRange().getValues();
  if (masterValues.length <= 1) return 0;

  const col = {};
  masterValues[0].forEach((h, i) => col[h] = i);
  const cycles = config.MAINTENANCE_CYCLES;
  const statusLabels = config.STATUS;
  const summaryRows = [];
  
  const storeHasWellPumpAlert = {};

  const calculatedData = masterValues.slice(1).map(row => {
    const locCode = row[col['拠点コード']];
    if (!locCode) return null;
    
    const eqId = String(row[col['設備ID']] || '');
    const eqName = String(row[col['設備名']] || '');
    const installDate = row[col['設置日(前回実施)']];
    const partADate = row[col['部品A交換日']];
    const partBDate = row[col['部品B最終交換日']];

    const res = calcStatusRow(installDate, partADate, partBDate, eqName, eqId, cycles, statusLabels, config.ALERT_THRESHOLDS);
    
    if (eqId.includes('WELL-P-01') && (res.partA !== statusLabels.NORMAL || res.body !== statusLabels.NORMAL)) {
      storeHasWellPumpAlert[locCode] = true;
    }
    
    return { row, res, locCode, eqId, eqName, installDate };
  }).filter(r => r !== null);

  let alertCount = 0;

  calculatedData.forEach(item => {
    const { row, res, locCode, eqId, eqName, installDate } = item;
    
    if (eqId.includes('MAINT-WELL-5Y') && storeHasWellPumpAlert[locCode]) {
      res.partA = statusLabels.NORMAL;
      res.partB = statusLabels.NORMAL;
      res.body = statusLabels.NORMAL;
    }

    if (res.partA !== statusLabels.NORMAL || res.partB !== statusLabels.NORMAL || res.body !== statusLabels.NORMAL) alertCount++;

    summaryRows.push([
      locCode, row[col['拠点名']], eqId, eqName, res.category, installDate,
      res.partA, res.partB, res.body, (res.partB !== statusLabels.NORMAL ? '対象' : ''), res.monthsA, row[col['備考']] || "", row[col['次回作業メモ']], row[col['型式・仕様']], res.nextDate
    ]);
  });

  summarySheet.clearContents();
  const headers = ['拠点コード', '拠点名', '設備ID', '設備名', 'カテゴリ', '設置日(前回実施)', '部品Aステータス', '部品Bステータス', '本体ステータス', '部品B対象', 'monthDiffA', 'subsidyAlert', 'nextWorkMemo', 'spec', '次回予定日'];
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (summaryRows.length > 0) {
    summarySheet.getRange(2, 1, summaryRows.length, headers.length).setValues(summaryRows);
    summarySheet.getRange(2, 6, summaryRows.length, 1).setNumberFormat('yyyy/MM/dd');
    summarySheet.getRange(2, 15, summaryRows.length, 1).setNumberFormat('yyyy/MM/dd');
  }
  return alertCount;
}

function calcStatusRow(installDate, partADate, partBDate, eqName, eqId, cycles, status, thresholds) {
  let partA = status.NORMAL, partB = status.NORMAL, body = status.NORMAL, monthsA = 0, nextDate = null, category = 'その他';
  const today = new Date();
  const isValidDate = (d) => d instanceof Date && !isNaN(d.getTime());

  // seasonal設備は通常判定をスキップ（ダッシュボードの一括発注アラートで表示）
  let matchedKeyForSeasonal = findCycleKey(eqId, eqName, cycles);
  if (matchedKeyForSeasonal && cycles[matchedKeyForSeasonal].seasonal) {
    category = cycles[matchedKeyForSeasonal].category;
    return { partA: status.NORMAL, partB: status.NORMAL, body: status.NORMAL, monthsA: 0, nextDate: null, category };
  }

  if (!isValidDate(installDate)) {
    let matchedKey = findCycleKey(eqId, eqName, cycles);
    if (matchedKey) category = cycles[matchedKey].category;
    if (category === '法定検査' && (eqName.includes('入替') || eqName.includes('更新'))) category = '本体更新';
    return { partA, partB, body, monthsA, nextDate, category };
  }

  let matchedKey = findCycleKey(eqId, eqName, cycles);

  if (matchedKey) {
    const c = cycles[matchedKey];
    category = c.category;
    
    let baseForNext;
    if (c.category === '部材交換' || c.category === '部材更新') {
      baseForNext = (isValidDate(partADate) ? partADate : installDate);
    } else if (c.category === 'メンテ') {
      baseForNext = (isValidDate(partBDate) ? partBDate : installDate);
    } else {
      baseForNext = installDate;
    }

    let tempNext = new Date(baseForNext);
    tempNext.setFullYear(tempNext.getFullYear() + c.years);
    if (!nextDate || tempNext < nextDate) nextDate = tempNext;

    const yearsBase = getYearsDiff(installDate, today);
    const yearsA = isValidDate(partADate) ? getYearsDiff(partADate, today) : yearsBase;
    monthsA = yearsA * 12;

    if (c.category === '本体更新') {
      if (yearsBase >= c.years + thresholds.BODY_PREPARE) body = status.PREPARE;
      else if (yearsBase >= c.years - thresholds.BODY_NOTICE) body = status.NOTICE;
    } 
    else if (c.category === '法定検査') {
      if (!eqName.includes('入替') && !eqName.includes('更新')) {
          if (yearsA >= c.years + thresholds.LEGAL_PREPARE) partA = status.PREPARE;
          else if (yearsA >= c.years - thresholds.LEGAL_NOTICE) partA = status.NOTICE;
      }
    } 
    else if (c.category === '美観') {
      if (yearsBase >= c.years - thresholds.BODY_NOTICE) body = status.NOTICE;
    } 
    else if (c.category === '部材交換' || c.category === '部材更新' || c.category === 'メンテ') {
      
      if (c.seasonal) {
        const lastDate = isValidDate(partADate) ? partADate : installDate;
        const yearsPassed = getYearsDiff(lastDate, today);
        const yearsToNext = c.years - yearsPassed;
        const currentMonth = today.getMonth() + 1;
        
        if (yearsPassed >= c.years) {
          partA = status.PREPARE;
        }
        else if (yearsToNext > 0 && yearsToNext < thresholds.SEASONAL_NOTICE && currentMonth >= 1) {
          partA = status.NOTICE;
        }
      } 
      else {
        if (yearsA >= c.years + thresholds.PARTS_PREPARE) {
          partA = status.PREPARE;
        } else if (yearsA >= c.years - thresholds.PARTS_NOTICE) {
          partA = status.NOTICE;
        }
      }
    }
  }

  if (eqName.includes('入替') || eqName.includes('更新')) {
    category = '本体更新';
  }

  return { partA, partB, body, monthsA, nextDate, category };
}

function findCycleKey(eqId, eqName, cycles) {
  for (const key in cycles) {
    const c = cycles[key];
    if (c.suffix && eqId === c.suffix) {
      return key;
    }
  }
  
  for (const key in cycles) {
    const c = cycles[key];
    if (c.suffix && eqId.includes(c.suffix)) {
      if (c.category === '法定検査' && (eqName.includes('入替') || eqName.includes('更新'))) {
        continue;
      }
      return key;
    }
  }
  
  for (const key in cycles) {
    const c = cycles[key];
    const searchWord = c.searchKey || c.label.replace(/[入替更新交換検定検査]/g,'').replace(/漏[洩え]い?/,'').replace(/\(.*\)/,'');
    
    if (eqName.includes(searchWord)) {
      if (c.category === '法定検査' && (eqName.includes('入替') || eqName.includes('更新'))) {
        continue;
      }
      return key;
    }
  }
  
  return null;
}

function getYearsDiff(d1, d2) {
  return (d2.getFullYear() - d1.getFullYear()) + ((d2.getMonth() - d1.getMonth()) / 12);
}

/**
 * ====================================================================
 * 4月実施一括発注関連（5種類）
 * ====================================================================
 */

/**
 * 4月実施一括発注の設備設定
 */
function getBulkOrderConfigs() {
  return [
    { id: 'PARTS-PUMP-1Y', name: 'ノズルカバー', cycle: 1, vendor: 'タツノ', emoji: '📦', searchKey: 'ノズルカバー' },
    { id: 'PARTS-SEAL-3Y', name: '釣銭機シール貼り替え', cycle: 3, vendor: 'シャープ', emoji: '🔧', searchKey: 'シール' },
    { id: 'CHG-01', name: '釣銭機カバー', cycle: 6, vendor: 'シャープ', emoji: '💳', searchKey: '釣銭機カバー' },
    { id: 'PARTS-PUMP-4Y', name: '計量機部品(4年)', cycle: 4, vendor: 'タツノ', emoji: '⛽', searchKey: '計量機部品' },
    { id: 'PARTS-K-PANEL-7Y', name: '灯油パネル更新', cycle: 7, vendor: 'タツノ', emoji: '🛢️', searchKey: '灯油パネル' }
  ];
}

/**
 * 設置日から最初の4月を計算
 */
function getFirstApril(installDate) {
  var firstApril = new Date(installDate.getFullYear(), 3, 1);
  if (installDate.getMonth() >= 3) {
    firstApril.setFullYear(firstApril.getFullYear() + 1);
  }
  return firstApril;
}

/**
 * ノズルカバー交換用: 設置日から実施可能な最初の4月を計算
 * 計量機更新後、2回目の4月を返す（更新後1回目の4月はスキップ）
 */
function getFirstAprilForNozzle(installDate) {
  var year = installDate.getFullYear();
  var month = installDate.getMonth(); // 0-11
  
  // 設置が1月〜3月なら翌年の4月、4月〜12月なら翌々年の4月（2回目の4月）
  var firstAprilYear = (month < 3) ? year + 1 : year + 2;
  return new Date(firstAprilYear, 3, 1); // 4月1日
}

/**
 * ノズルカバー交換の対象店舗を取得
 * 計量機を持つ全店舗が対象（計量機更新から1年未満は除外）
 */
function getNozzleCoverTargetStores() {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  var masterValues = masterSheet.getDataRange().getValues();
  
  if (masterValues.length <= 1) return [];
  
  var col = {};
  for (var i = 0; i < masterValues[0].length; i++) {
    col[masterValues[0][i]] = i;
  }
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  var targetApril = new Date(targetYear, 3, 1);
  
  var storeMap = {};
  
  for (var i = 1; i < masterValues.length; i++) {
    var row = masterValues[i];
    var locCode = row[col['拠点コード']];
    var locName = row[col['拠点名']];
    var eqId = String(row[col['設備ID']] || '');
    var installDate = row[col['設置日(前回実施)']];
    
    if (!locCode || !locName) continue;
    
    var isPump = eqId.includes('PUMP-G-01') || eqId.includes('PUMP-K-01');
    
    if (isPump && installDate instanceof Date && !isNaN(installDate.getTime())) {
      var firstApril = getFirstAprilForNozzle(installDate);
      
      if (targetApril >= firstApril) {
        // 店舗コードをキーにして重複を防ぐ
        if (!storeMap[locCode]) {
          storeMap[locCode] = {
            code: locCode,
            name: locName,
            installDate: installDate,
            firstApril: firstApril
          };
        }
      }
    }
  }
  
  var result = [];
  for (var key in storeMap) {
    result.push(storeMap[key]);
  }
  
  result.sort(function(a, b) {
    return a.code > b.code ? 1 : -1;
  });
  
  return result;
}

/**
 * ノズルカバー交換メール下書き作成
 */
function createNozzleCoverDraftEmail(targetStores) {
  if (targetStores.length === 0) return '現在、発注対象の店舗はありません。';
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var fiscalYear = (currentMonth >= 1 && currentMonth <= 3) ? today.getFullYear() : today.getFullYear() + 1;
  
  var body = '';
  body += 'お世話になっております。\n\n';
  body += fiscalYear + '年度のノズルカバー交換の発注をお願いいたします。\n\n';
  body += '【対象店舗: ' + targetStores.length + '店舗（全店）】\n\n';
  
  for (var i = 0; i < targetStores.length; i++) {
    var store = targetStores[i];
    body += '- ' + store.name + '\n';
  }
  
  body += '\n【実施予定】\n' + fiscalYear + '年4月\n\n';
  body += '【発注先】\nタツノ\n\n';
  body += 'よろしくお願いいたします。\n\n';
  body += '--------------------------------------------------\n';
  body += '日商有田株式会社\n';
  body += 'nishimura@selfix.jp\n';
  body += '--------------------------------------------------';
  
  return body;
}

/**
 * ノズルカバー一括発注情報を取得（ダッシュボード表示用）
 */
function getNozzleCoverInfo() {
  var targetStores = getNozzleCoverTargetStores();
  var emailDraft = createNozzleCoverDraftEmail(targetStores);
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1; // 1-12
  var currentYear = today.getFullYear();
  
  // 1月〜3月は今年4月、4月以降は来年4月を実施予定とする
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
  return {
    config: {
      id: 'PARTS-PUMP-1Y',
      name: 'ノズルカバー交換',
      emoji: '📦',
      vendor: 'タツノ'
    },
    hasAlert: targetStores.length > 0,
    targetCount: targetStores.length,
    targetStores: targetStores,
    emailDraft: emailDraft,
    targetYear: targetYear
  };
}

/**
 * 一括発注対象店舗を取得（汎用）
 */
function getBulkOrderTargetStores(equipmentId, cycleYears, searchKey) {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  var masterValues = masterSheet.getDataRange().getValues();
  
  if (masterValues.length <= 1) return [];
  
  var col = {};
  for (var i = 0; i < masterValues[0].length; i++) {
    col[masterValues[0][i]] = i;
  }
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  
  // 1月から3月は今年4月、4月以降は来年4月を実施予定とする
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  var targetApril = new Date(targetYear, 3, 1); // 4月1日
  
  var storeMap = {};
  
  for (var i = 1; i < masterValues.length; i++) {
    var row = masterValues[i];
    var locCode = row[col['拠点コード']];
    var locName = row[col['拠点名']];
    var eqId = String(row[col['設備ID']] || '');
    var eqName = String(row[col['設備名']] || '');
    var installDate = row[col['設置日(前回実施)']];
    var partADate = row[col['部品A交換日']];
    
    if (!locCode || !locName) continue;
    
    var isMatch = (eqId.indexOf(equipmentId) >= 0) || (searchKey && eqName.indexOf(searchKey) >= 0);
    
    if (isMatch && installDate instanceof Date && !isNaN(installDate.getTime())) {
      var baseDate = (partADate instanceof Date && !isNaN(partADate.getTime())) ? partADate : installDate;
      
      var firstApril = getFirstApril(baseDate);
      // 実施予定の4月時点で、cycleYears年以上経過している店舗を抽出
      var yearsUntilTargetApril = getYearsDiff(firstApril, targetApril);
      
      // 今年または来年の4月までに、cycleYears年以上経過する予定の店舗を抽出
      if (yearsUntilTargetApril >= cycleYears && !storeMap[locCode]) {
        var yearsSinceFirstApril = getYearsDiff(firstApril, today);
        storeMap[locCode] = {
          code: locCode,
          name: locName,
          equipmentName: eqName, // 設備名を追加
          lastDate: baseDate,
          firstApril: firstApril,
          yearsSinceFirstApril: yearsSinceFirstApril,
          yearsUntilTargetApril: yearsUntilTargetApril,
          targetApril: targetApril,
          hasHistory: (partADate instanceof Date && !isNaN(partADate.getTime()))
        };
      }
    }
  }
  
  var result = [];
  for (var key in storeMap) {
    result.push(storeMap[key]);
  }
  
  result.sort(function(a, b) {
    return a.code > b.code ? 1 : -1;
  });
  
  return result;
}

/**
 * 一括発注メール下書き作成（汎用）
 */
function createBulkOrderDraftEmail(configItem, targetStores, targetYear) {
  if (targetStores.length === 0) return '現在、発注対象の店舗はありません。';
  
  // targetYearが指定されていない場合は、現在の日付から計算
  if (!targetYear) {
    var today = new Date();
    var currentMonth = today.getMonth() + 1;
    var currentYear = today.getFullYear();
    targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  }
  var fiscalYear = targetYear; // 実施年度
  
  // 計量器設備かどうかを判定（PARTS-PUMP-1Y, PARTS-PUMP-4Y は計量器）
  var isMeasuringEquipment = (configItem.id === 'PARTS-PUMP-1Y' || configItem.id === 'PARTS-PUMP-4Y');
  
  var body = '';
  body += 'お世話になっております。\n\n';
  body += fiscalYear + '年度の' + configItem.name + 'の発注をお願いいたします。\n\n';
  body += '【対象店舗: ' + targetStores.length + '店舗】\n';
  
  for (var i = 0; i < targetStores.length; i++) {
    var store = targetStores[i];
    var lastYear = store.lastDate.getFullYear();
    var lastMonth = store.lastDate.getMonth() + 1;
    body += '- ' + store.name + '（前回: ' + lastYear + '年' + lastMonth + '月）';
    
    // 計量器設備の場合、設備名を記載（型式・仕様は記載しない）
    if (isMeasuringEquipment && store.equipmentName) {
      body += '\n  ' + store.equipmentName;
    }
    body += '\n';
  }
  
  body += '\n【実施予定】\n' + targetYear + '年4月\n\n';
  body += '【発注先】\n' + configItem.vendor + '\n\n';
  body += 'よろしくお願いいたします。\n\n';
  body += '--------------------------------------------------\n';
  body += '日商有田株式会社\n';
  body += 'nishimura@selfix.jp\n';
  body += '--------------------------------------------------';
  return body;
}

/**
 * 全ての一括発注情報を取得
 */
function getAllBulkOrderInfo() {
  var configs = getBulkOrderConfigs();
  var results = [];
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
  for (var i = 0; i < configs.length; i++) {
    var cfg = configs[i];
    var targetStores = getBulkOrderTargetStores(cfg.id, cfg.cycle, cfg.searchKey);
    var emailDraft = createBulkOrderDraftEmail(cfg, targetStores, targetYear);
    
    results.push({
      config: cfg,
      hasAlert: targetStores.length > 0,
      targetCount: targetStores.length,
      targetStores: targetStores,
      emailDraft: emailDraft,
      targetYear: targetYear // 実施予定年度を追加
    });
  }
  
  return results;
}

/**
 * ノズルカバー交換Gmail下書き作成
 */
function createNozzleCoverGmailDraft() {
  var config = getConfig();
  var targetStores = getNozzleCoverTargetStores();
  if (targetStores.length === 0) throw new Error('発注対象の店舗がありません');
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
  var body = createNozzleCoverDraftEmail(targetStores);
  var subject = '【' + targetYear + '年度】ノズルカバー交換 発注のご依頼';
  
  // ベンダーのメールアドレスを取得
  var vendorEmail = '';
  for (var key in config.VENDORS) {
    var vendorName = config.VENDORS[key].name;
    if (vendorName.includes('タツノ') || 'タツノ'.includes(vendorName.replace('株式会社', '').replace('有限会社', ''))) {
      vendorEmail = config.VENDORS[key].email || '';
      break;
    }
  }
  
  // Gmailの下書きを作成（送信元はnishimura@selfix.jp）
  GmailApp.createDraft(vendorEmail || '', subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  return {
    success: true,
    message: 'Gmailの下書きを作成しました',
    subject: subject,
    recipient: vendorEmail || '（送信先未設定）'
  };
}

/**
 * ノズルカバー交換案件作成
 */
function createNozzleCoverProject() {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var scheduleSheet = ss.getSheetByName(config.SHEET_NAMES.SCHEDULE);
  
  var targetStores = getNozzleCoverTargetStores();
  if (targetStores.length === 0) throw new Error('発注対象の店舗がありません');
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  var scheduledDate = new Date(targetYear, 3, 1); // 4月1日
  var projectId = 'PARTS-PUMP-1Y-' + targetYear + '-' + Utilities.formatDate(new Date(), 'JST', 'MMddHHmmss');
  
  var newRow = [
    projectId,
    '全店',
    'PARTS-PUMP-1Y',
    '【一括発注】ノズルカバー交換 ' + targetStores.length + '店舗',
    scheduledDate,
    '見積依頼中',
    '',
    'タツノ'
  ];
  
  scheduleSheet.appendRow(newRow);
  var lastRow = scheduleSheet.getLastRow();
  scheduleSheet.getRange(lastRow, 5).setNumberFormat('yyyy/MM/dd');
  
  return {
    success: true,
    projectId: projectId,
    equipmentName: 'ノズルカバー交換',
    targetCount: targetStores.length
  };
}

/**
 * 一括発注メール下書きをGmailに作成
 */
function createBulkOrderGmailDraft(equipmentId) {
  var config = getConfig();
  var configs = getBulkOrderConfigs();
  var cfg = null;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].id === equipmentId) {
      cfg = configs[i];
      break;
    }
  }
  
  if (!cfg) throw new Error('設備IDが見つかりません: ' + equipmentId);
  
  var targetStores = getBulkOrderTargetStores(cfg.id, cfg.cycle, cfg.searchKey);
  if (targetStores.length === 0) throw new Error('発注対象の店舗がありません');
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
  var body = createBulkOrderDraftEmail(cfg, targetStores, targetYear);
  var subject = '【' + targetYear + '年度】' + cfg.name + ' 発注のご依頼';
  
  // ベンダーのメールアドレスを取得（ベンダー名でマッチング）
  var vendorEmail = '';
  for (var key in config.VENDORS) {
    var vendorName = config.VENDORS[key].name;
    // 'タツノ' は '株式会社タツノ' に、'シャープ' は 'シャープ' にマッチ
    if (vendorName.includes(cfg.vendor) || cfg.vendor.includes(vendorName.replace('株式会社', '').replace('有限会社', ''))) {
      vendorEmail = config.VENDORS[key].email || '';
      break;
    }
  }
  
  // Gmailの下書きを作成（送信元はnishimura@selfix.jp）
  GmailApp.createDraft(vendorEmail || '', subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  return {
    success: true,
    message: 'Gmailの下書きを作成しました',
    subject: subject,
    recipient: vendorEmail || '（送信先未設定）'
  };
}

/**
 * 一括発注案件を作成（汎用）
 */
function createBulkOrderProject(equipmentId) {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var scheduleSheet = ss.getSheetByName(config.SHEET_NAMES.SCHEDULE);
  
  var configs = getBulkOrderConfigs();
  var cfg = null;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].id === equipmentId) {
      cfg = configs[i];
      break;
    }
  }
  
  if (!cfg) throw new Error('設備IDが見つかりません: ' + equipmentId);
  
  var targetStores = getBulkOrderTargetStores(cfg.id, cfg.cycle, cfg.searchKey);
  
  if (targetStores.length === 0) throw new Error('発注対象の店舗がありません');
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  // 1月から3月は今年4月、4月以降は来年4月を実施予定とする
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  var scheduledDate = new Date(targetYear, 3, 1); // 4月1日
  var projectId = cfg.id.replace(/[^A-Z0-9]/g, '') + '-' + targetYear + '-' + Utilities.formatDate(new Date(), 'JST', 'MMddHHmmss');
  
  var newRow = [
    projectId,
    '全店',
    cfg.id,
    '【一括発注】' + cfg.name + ' ' + targetStores.length + '店舗',
    scheduledDate,
    '見積依頼中',
    '',
    cfg.vendor
  ];
  
  scheduleSheet.appendRow(newRow);
  var lastRow = scheduleSheet.getLastRow();
  scheduleSheet.getRange(lastRow, 5).setNumberFormat('yyyy/MM/dd');
  
  return {
    success: true,
    projectId: projectId,
    equipmentName: cfg.name,
    targetCount: targetStores.length
  };
}

/**
 * ====================================================================
 * その他の既存関数
 * ====================================================================
 */

function checkAndSendAlertMail() {
  const config = getConfig();
  const summarySheet = getSheet(config.SHEET_NAMES.STATUS_SUMMARY);
  const data = summarySheet.getDataRange().getValues();
  if (data.length <= 1) return;
  let hasAlert = data.some((row, i) => i > 0 && (row[6] === '期限超過' || row[6] === '実施時期' || row[7] === '期限超過' || row[8] === '期限超過'));
  if (hasAlert) {
    const admin = config.ADMIN_MAIL || 'nishimura@selfix.jp';
    GmailApp.sendEmail(admin, '【SS設備管理】メンテナンスアラート', '設備管理システムを確認してください。\n' + ScriptApp.getService().getUrl());
  }
}

function runDailyBackup() { checkAndSendAlertMail(); }

function setupSystemTriggers() {
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'runDailyBackup')) {
    ScriptApp.newTrigger('runDailyBackup').timeBased().atHour(9).everyDays(1).create();
  }
}

function importEquipmentData(ss, config) { 
  const masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const importSheet = ss.getSheetByName('データ取込');
  if (!importSheet) return;

  const range = masterSheet.getDataRange();
  const currentValues = range.getValues();
  let validRows = [], existingKeys = new Set(), deletedCount = 0, isDirty = false;

  if (currentValues.length > 0) validRows.push(currentValues[0]);
  if (currentValues.length > 1) {
    for (let i = 1; i < currentValues.length; i++) {
      const row = currentValues[i];
      const locCode = String(row[0]||''), eqId = String(row[2]||''), eqName = String(row[3]||'');
      if (eqName.includes('移動ポンプ') || eqId.includes('M-PUMP') || eqId.includes('MOBILE-PUMP')) {
        deletedCount++; isDirty = true;
      } else {
        validRows.push(row);
        if (locCode && eqId) existingKeys.add(`${locCode}_${eqId}`);
      }
    }
  }

  if (isDirty) {
    masterSheet.clearContents();
    if (validRows.length > 0) {
      masterSheet.getRange(1, 1, validRows.length, validRows[0].length).setValues(validRows);
      if (validRows.length > 1) masterSheet.getRange(2, 6, validRows.length - 1, 3).setNumberFormat('yyyy/MM/dd');
    }
    SpreadsheetApp.flush();
  }

  const stores = getStoreList();
  const rowsToAdd = [], cycles = config.MAINTENANCE_CYCLES, templates = [];
  for (const key in cycles) {
    if (key.includes('MOBILE_PUMP') || key.includes('移動ポンプ')) continue;
    if (cycles[key].suffix) templates.push({ suffix: cycles[key].suffix, name: cycles[key].label, searchKey: cycles[key].searchKey || '' });
  }

  const lastRow = importSheet.getLastRow();
  let values = [], headerMap = {};
  if (lastRow > 1) {
    values = importSheet.getRange(1, 1, lastRow, importSheet.getLastColumn()).getValues();
    values[0].forEach((h, i) => {
      const tmpl = templates.find(t => t.searchKey && String(h).includes(t.searchKey));
      if (tmpl) headerMap[tmpl.suffix] = i;
    });
  }

  let importData = {};
  for (let i = 1; i < values.length; i++) {
    const rowStoreName = String(values[i][0]).trim();
    if (rowStoreName) {
      const matchedStore = stores.find(s => rowStoreName.includes(s.name) || s.name.includes(rowStoreName));
      if (matchedStore) {
        if (!importData[matchedStore.name]) importData[matchedStore.name] = {};
        for (const sfx in headerMap) {
          const val = values[i][headerMap[sfx]];
          if (val) importData[matchedStore.name][sfx] = parseCellData(val).text;
        }
      }
    }
  }

  stores.forEach(store => {
    const sCode = store.code || ('SS' + ('000' + (Math.random()*1000).toFixed(0)).slice(-3));
    const storeImport = importData[store.name] || {};
    templates.forEach(tmpl => {
      if (existingKeys.has(`${sCode}_${tmpl.suffix}`)) return;
      let spec = storeImport[tmpl.suffix] || "";
      if (tmpl.suffix === 'PUMP-K-01' || tmpl.suffix === 'PUMP-K-CHK') {
         if (!spec && storeImport['PUMP-K-01']) spec = storeImport['PUMP-K-01'];
      }
      
      let memo = '';
      if (tmpl.suffix === 'PARTS-SEAL-3Y') {
        memo = 'お願いシールとお札は1枚ずつのみ';
      }

      rowsToAdd.push([sCode, store.name, tmpl.suffix, tmpl.name, spec, '', '', '', '', memo]);
    });
  });

  if (rowsToAdd.length > 0) {
    const startRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(startRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    masterSheet.getRange(startRow, 6, rowsToAdd.length, 3).setNumberFormat('yyyy/MM/dd');
    Logger.log(`${rowsToAdd.length}件の設備を追加しました。`);
  } else {
    Logger.log('全て登録済みです。追加設備はありません。');
  }
}

function parseCellData(val) {
  if (!val) return { date: '', text: '' };
  const str = String(val).trim();
  let text = str.replace(/^(\d{4})[\.\/-](\d{1,2})(?:[\.\/-](\d{1,2}))?/, '').trim();
  text = text.replace(/^(\d{4})/, '').trim();
  return { date: '', text: text || str };
}

function getStoreList() {
  try {
    const config = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_LOCATION);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      return data.filter(r => r[1]).map(r => ({ code: r[0], name: r[1] }));
    }
  } catch (e) {}
  return [{ name: '糸我' }, { name: 'かつらぎ' }, { name: '和佐' }, { name: '熊野' }, { name: '貴志川' }, { name: 'りんくう泉南' }, { name: '御所' }, { name: '東和歌山' }, { name: '和歌山北インター' }, { name: '紀三井寺' }, { name: '天理' }, { name: '厚木' }, { name: '坂出' }, { name: '裾野' }, { name: '徳島石井' }, { name: '小松島' }, { name: '池田' }, { name: '倉吉' }, { name: '小山' }, { name: '岡南' }, { name: '牛久' }, { name: '土浦' }, { name: '岐阜東' }, { name: '太田' }, { name: '北名古屋' }, { name: 'ひたちなか' }].map((d, i) => ({ code: 'SS' + ('000' + (i + 1)).slice(-3), name: d.name }));
}

/**
 * テスト関数
 */
function testAllBulkOrders() {
  var allInfo = getAllBulkOrderInfo();
  
  allInfo.forEach(function(info) {
    Logger.log('=== ' + info.config.name + ' ===');
    Logger.log('対象店舗数: ' + info.targetCount);
    Logger.log('アラート: ' + info.hasAlert);
    
    if (info.targetStores.length > 0) {
      info.targetStores.forEach(function(s) {
        var type = s.hasHistory ? '[交換済み]' : '[未実施]';
        Logger.log('  ' + s.name + ' ' + type + ' / ' + s.yearsSinceFirstApril.toFixed(1) + '年経過');
      });
    }
    Logger.log('');
  });
}

/**
 * ノズルカバー対象店舗のデバッグ表示
 */
function debugNozzleCover() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const masterValues = masterSheet.getDataRange().getValues();
  
  Logger.log('=== ノズルカバー対象店舗デバッグ ===');
  Logger.log('今日の日付: ' + new Date());
  
  const col = {};
  masterValues[0].forEach((h, i) => { col[h] = i; });
  
  const today = new Date();
  const currentMonth = today.getMonth() + 1;
  const currentYear = today.getFullYear();
  const targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  const targetApril = new Date(targetYear, 3, 1);
  
  Logger.log('現在月: ' + currentMonth + '月');
  Logger.log('実施予定年: ' + targetYear + '年4月');
  Logger.log('---');
  
  let pumpCount = 0;
  let eligibleCount = 0;
  const eligibleStores = [];
  
  for (let i = 1; i < masterValues.length; i++) {
    const row = masterValues[i];
    const locCode = row[col['拠点コード']];
    const locName = row[col['拠点名']];
    const eqId = String(row[col['設備ID']] || '');
    const eqName = String(row[col['設備名']] || '');
    const installDate = row[col['設置日(前回実施)']];
    const partADate = row[col['部品A交換日']];
    
    if (!locCode || !locName) continue;
    
    const isPump = eqId.includes('PUMP-G-01') || eqId.includes('PUMP-K-01');
    
    if (isPump) {
      pumpCount++;
      Logger.log(`[${locName}] 設備ID: ${eqId}`);
      
      if (installDate instanceof Date && !isNaN(installDate.getTime())) {
        const baseDate = (partADate instanceof Date && !isNaN(partADate.getTime())) ? partADate : installDate;
        Logger.log(`  基準日: ${Utilities.formatDate(baseDate, 'JST', 'yyyy/MM/dd')}`);
        
        const year = baseDate.getFullYear();
        const month = baseDate.getMonth();
        const firstAprilYear = (month < 3) ? year + 1 : year + 2;
        const firstApril = new Date(firstAprilYear, 3, 1);
        
        Logger.log(`  初回実施可能日: ${firstAprilYear}年4月`);
        Logger.log(`  判定: ${targetYear}年4月 >= ${firstAprilYear}年4月 = ${targetYear >= firstAprilYear}`);
        
        if (targetYear >= firstAprilYear) {
          eligibleCount++;
          eligibleStores.push(locName);
          Logger.log(`  ✓ 対象に含まれます`);
        } else {
          Logger.log(`  × まだ対象外`);
        }
      } else {
        Logger.log(`  × 設置日なし`);
      }
      Logger.log('---');
    }
  }
  
  Logger.log('====================');
  Logger.log(`計量機設備数: ${pumpCount}`);
  Logger.log(`対象店舗数: ${eligibleCount}`);
  Logger.log(`対象店舗: ${eligibleStores.join(', ')}`);
  Logger.log('====================');
}