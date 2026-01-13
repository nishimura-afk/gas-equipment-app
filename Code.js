/**
 * Code.gs v7.5
 * 本番運用版
 */
function doGet() {
  console.log('doGet START v7.4');
  const t = HtmlService.createTemplateFromFile('index');
  t.include = function(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); };
  return t.evaluate()
    .setTitle('SS設備管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getEquipmentList() { return getEquipmentListCached(); }

function getDashboardData() {
  const data = getEquipmentListCached();
  const config = getConfig();
  const scheduleData = getSheet(config.SHEET_NAMES.SCHEDULE).getDataRange().getValues();
  const ignoreActions = scheduleData.slice(1)
    .filter(row => row[5] !== config.PROJECT_STATUS.COMPLETED && row[5] !== config.PROJECT_STATUS.CANCELLED)
    .map(row => `${row[1]}_${row[2]}`);

  const notices = data.filter(m => {
    if (ignoreActions.includes(`${m['拠点コード']}_${m['設備ID']}`)) return false;
    return m['本体ステータス'] !== config.STATUS.NORMAL || m['部品Aステータス'] !== config.STATUS.NORMAL || (m['部品Bステータス'] && m['部品Bステータス'] !== config.STATUS.NORMAL);
  });
  return { noticeCount: notices.length, normalCount: data.length - notices.length, noticeList: notices };
}

function getAllActiveProjects() {
  const config = getConfig();
  const data = getSheet(config.SHEET_NAMES.SCHEDULE).getDataRange().getValues();
  if (data.length <= 1) return [];
  const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
  const locData = locSheet.getDataRange().getValues();
  const locMap = {};
  locData.slice(1).forEach(r => { if(r[0]) locMap[r[0]] = r[1]; });
  const equipmentList = getEquipmentListCached();
  const eqMap = {};
  equipmentList.forEach(row => {
    eqMap[`${row['拠点コード']}_${row['設備ID']}`] = row['設備名'] || row['設備ID'];
  });
  return data.slice(1).map((r, i) => {
    const locCode = r[1];
    const eqId = r[2];
    const key = `${locCode}_${eqId}`;
    return {
      id: r[0],
      locCode: locCode,
      locName: locMap[locCode] || locCode, 
      equipmentId: eqId,
      equipmentName: eqMap[key] || eqId,   
      workType: r[3],
      date: (r[4] instanceof Date) ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[4],
      status: r[5],
      rowNumber: i + 2
    };
  }).filter(p => p.status !== config.PROJECT_STATUS.COMPLETED && p.status !== config.PROJECT_STATUS.CANCELLED);
}

function getExchangeTargetsForUI() {
  return getDashboardData().noticeList.map(m => ({
    locCode: m['拠点コード'], locName: m['拠点名'], equipmentId: m['設備ID'], equipmentName: m['設備名'] || m['設備ID'],
    exchangeTargets: [m['部品Aステータス']!=='正常'?'消耗品':null, m['本体ステータス']!=='正常'?'本体':null].filter(v=>v).join('/'),
    subsidyAlert: m['subsidyAlert'], nextWorkMemo: m['nextWorkMemo'], category: m['カテゴリ']
  }));
}

function updateProjectStatus(id, newStatus) {
  const sheet = getSheet(getConfig().SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(newStatus);
      return { success: true };
    }
  }
}

function cancelProject(id) {
  const sheet = getSheet(getConfig().SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(getConfig().PROJECT_STATUS.CANCELLED);
      return { success: true };
    }
  }
}

function createScheduleAndRecord(loc, eq, work, date, notes, existingId = null) {
  const config = getConfig();
  const r = createMaintenanceEvent(loc, eq, work, date, notes);
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  if (existingId && existingId !== 'DIRECT') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(existingId)) {
        sheet.getRange(i + 1, 5).setValue(date);
        sheet.getRange(i + 1, 6).setValue(config.PROJECT_STATUS.SCHEDULED);
        sheet.getRange(i + 1, 7).setValue(r.eventId);
        sheet.getRange(i + 1, 8).setValue(notes);
        return r;
      }
    }
  } else {
    const uniqueId = Utilities.getUuid();
    sheet.appendRow([uniqueId, loc, eq, work, date, config.PROJECT_STATUS.SCHEDULED, r.eventId, notes]);
  }
  return r;
}

function completeExchange(uniqueId, date, subsidy) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[0] == uniqueId);
  if (!row) throw new Error('案件不明');
  recordExchangeComplete(row[1], row[2], row[3], date, subsidy);
  markEventAsCompleted(row[6], date);
  sheet.getRange(data.indexOf(row) + 1, 6).setValue(config.PROJECT_STATUS.COMPLETED);
  return { message: '成功' };
}

function generateQuoteRequest(locName, eqName, workType) {
  let displayEqName = eqName;
  if (displayEqName.includes('釣銭機カバー')) displayEqName = displayEqName.replace('釣銭機カバー', '投入/取出し口のプラスチックカバー');
  if (displayEqName.includes('パネル')) displayEqName = displayEqName.replace('パネル', 'タッチパネル');
  return `見積依頼...`;
}

// =================================================================
// 4月実施一括発注ロジック
// =================================================================

function getBulkOrderConfigs() {
  return [
    { id: 'PARTS-PUMP-1Y', name: 'ノズルカバー', cycle: 1, vendor: 'タツノ', emoji: '📦', searchKey: 'ノズルカバー' },
    { id: 'PARTS-SEAL-3Y', name: '釣銭機シール貼り替え', cycle: 3, vendor: 'シャープ', emoji: '🔧', searchKey: 'シール' },
    { id: 'CHG-01', name: '釣銭機カバー', cycle: 6, vendor: 'シャープ', emoji: '💳', searchKey: '釣銭機カバー' },
    { id: 'PARTS-PUMP-4Y', name: 'ガソリン計量機部品(4年)', cycle: 4, vendor: 'タツノ', emoji: '⛽', searchKey: 'ガソリン計量機部品' },
    { id: 'PARTS-K-PANEL-7Y', name: '灯油パネル更新', cycle: 7, vendor: 'タツノ', emoji: '🛢️', searchKey: '灯油パネル' }
  ];
}

function getFiscalYear(date) {
  if (!date || isNaN(date.getTime())) return 0;
  return (date.getMonth() < 3) ? date.getFullYear() - 1 : date.getFullYear();
}

function getNozzleCoverTargetStores() {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  var masterValues = masterSheet.getDataRange().getValues();
  if (masterValues.length <= 1) return [];
  var col = {};
  for (var i = 0; i < masterValues[0].length; i++) { col[masterValues[0][i]] = i; }
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
  var storeDates = {};
  for (var i = 1; i < masterValues.length; i++) {
    var row = masterValues[i];
    var locCode = row[col['拠点コード']];
    var locName = row[col['拠点名']];
    var eqId = String(row[col['設備ID']] || '');
    var installDate = row[col['設置日(前回実施)']];
    
    if (!locCode || !locName) continue;
    if (!storeDates[locCode]) storeDates[locCode] = { code: locCode, name: locName, dates: [] };
    
    if (installDate instanceof Date && !isNaN(installDate.getTime()) && installDate <= today) {
      if (eqId === 'PARTS-PUMP-1Y' || eqId.includes('PUMP-G-01')) {
        storeDates[locCode].dates.push(installDate);
      }
    }
  }
  
  // targetYear の 4月1日を基準日とする
var targetApril = new Date(targetYear, 3, 1);

var result = [];
for (var locCode in storeDates) {
  var store = storeDates[locCode];
  if (store.dates.length === 0) continue;
  var latestDate = new Date(Math.max.apply(null, store.dates));
  
  // 前回実施日から1年後を計算
  var oneYearLater = new Date(latestDate);
  oneYearLater.setFullYear(oneYearLater.getFullYear() + 1);
  
  // 1年後が targetYear の 4月1日以前なら対象
  if (oneYearLater <= targetApril) {
    result.push({ code: store.code, name: store.name, installDate: latestDate, targetYear: targetYear });
  }
}
  result.sort(function(a, b) { return a.code > b.code ? 1 : -1; });
  return result;
}

function createNozzleCoverDraftEmail(targetStores) {
  if (!targetStores || targetStores.length === 0) return '現在、発注対象の店舗はありません。';
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var fiscalYear = (currentMonth >= 1 && currentMonth <= 3) ? today.getFullYear() : today.getFullYear() + 1;
  var body = 'お世話になっております。\n\n' + fiscalYear + '年度のノズルカバー交換の発注をお願いいたします。\n\n【対象店舗: ' + targetStores.length + '店舗（全店）】\n\n';
  for (var i = 0; i < targetStores.length; i++) { body += '- ' + targetStores[i].name + '\n'; }
  body += '\n【実施予定】\n' + fiscalYear + '年4月\n\n【発注先】\nタツノ\n\nよろしくお願いいたします。\n\n--------------------------------------------------\n日商有田株式会社\nnishimura@selfix.jp\n--------------------------------------------------';
  return body;
}

function getNozzleCoverInfo() {
  try {
    var targetStores = getNozzleCoverTargetStores();
    var today = new Date();
    var currentMonth = today.getMonth() + 1;
    var currentYear = today.getFullYear();
    var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;

    var emailDraft = createNozzleCoverDraftEmail(targetStores);
    
    // 日付オブジェクトを文字列に変換して返す(null化回避)
    var safeStores = targetStores.map(s => ({
      code: s.code,
      name: s.name,
      installDate: Utilities.formatDate(s.installDate, 'JST', 'yyyy/MM/dd'),
      targetYear: s.targetYear
    }));

    return {
      config: { id: 'PARTS-PUMP-1Y', name: 'ノズルカバー交換', emoji: '📦', vendor: 'タツノ' },
      hasAlert: safeStores.length > 0,
      targetCount: safeStores.length,
      targetStores: safeStores,
      emailDraft: emailDraft,
      targetYear: targetYear
    };
  } catch (e) {
    return { hasAlert: false, error: e.toString() };
  }
}

function getBulkOrderTargetStores(equipmentId, cycleYears, searchKey) {
  var config = getConfig();
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(config.SHEET_NAMES.MASTER_EQUIPMENT);
  var masterValues = masterSheet.getDataRange().getValues();
  if (masterValues.length <= 1) return [];
  var col = {};
  for (var i = 0; i < masterValues[0].length; i++) { col[masterValues[0][i]] = i; }
  
  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var currentYear = today.getFullYear();
  var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
  
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
      var installFY = getFiscalYear(baseDate);
      var targetFY = targetYear;
      var diffYears = targetFY - installFY;
      
      if (diffYears >= cycleYears && !storeMap[locCode]) {
        storeMap[locCode] = {
          code: locCode, name: locName, equipmentName: eqName,
          lastDate: baseDate, lastFY: installFY, targetFY: targetFY, diffYears: diffYears
        };
      }
    }
  }
  var result = [];
  for (var key in storeMap) { result.push(storeMap[key]); }
  return result;
}

function createBulkOrderDraftEmail(configItem, targetStores, targetYear) {
  if (!targetStores || targetStores.length === 0) return '対象なし';
  var fiscalYear = targetYear || ((new Date().getMonth() < 3) ? new Date().getFullYear() : new Date().getFullYear() + 1);
  var body = 'お世話になっております。\n\n' + fiscalYear + '年度の' + configItem.name + 'の発注をお願いいたします。\n\n【対象店舗: ' + targetStores.length + '店舗】\n';
  for (var i = 0; i < targetStores.length; i++) {
    var s = targetStores[i];
    body += '- ' + s.name + ' (前回: ' + s.lastDate.getFullYear() + '年' + (s.lastDate.getMonth()+1) + '月)\n';
    if ((configItem.id.includes('PUMP')) && s.equipmentName) body += '  ' + s.equipmentName + '\n';
  }
  body += '\n【実施予定】\n' + fiscalYear + '年4月\n\n【発注先】\n' + configItem.vendor + '\n\nよろしくお願いいたします。\n\n--------------------------------------------------\n日商有田株式会社\nnishimura@selfix.jp\n--------------------------------------------------';
  return body;
}

function getAllBulkOrderInfo() {
  try {
    var configs = getBulkOrderConfigs();
    var results = [];
    var today = new Date();
    var targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    for (var i = 0; i < configs.length; i++) {
      var cfg = configs[i];
      if (cfg.id === 'PARTS-PUMP-1Y') continue; 
      var targetStores = getBulkOrderTargetStores(cfg.id, cfg.cycle, cfg.searchKey);
      
      // 対象がない場合はスキップ
      if (targetStores.length === 0) continue;
      
      var emailDraft = createBulkOrderDraftEmail(cfg, targetStores, targetYear);

      // 日付の安全化
      var safeStores = targetStores.map(s => ({
        code: s.code, name: s.name, equipmentName: s.equipmentName,
        lastDate: Utilities.formatDate(s.lastDate, 'JST', 'yyyy/MM/dd'),
        diffYears: s.diffYears
      }));

      results.push({
        config: cfg,
        hasAlert: true,
        targetCount: safeStores.length,
        targetStores: safeStores,
        emailDraft: emailDraft,
        targetYear: targetYear
      });
    }
    return results;
  } catch (e) {
    return [];
  }
}

/**
 * 電話依頼＋案件作成（エアコンなど）
 */
function createPhoneCallProject(locCode, eqId, eqName, memo) {
  const config = getConfig();
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  
  const uniqueId = Utilities.getUuid();
  const workType = `${eqName}更新（電話依頼）`;
  
  scheduleSheet.appendRow([
    uniqueId,
    locCode,
    eqId,
    workType,
    '', // 日程は後で入力
    config.PROJECT_STATUS.ESTIMATE_REQ, // ステータス：見積依頼中
    '',
    memo // 備考欄にメモを記録
  ]);
  
  return { success: true, message: '案件を作成しました' };
}

/**
 * 個別案件のGmail下書き作成
 */
function createIndividualGmailDraft(locCode, eqId, locName, eqName, workType) {
  try {
    var subject = '【見積依頼】見積り依頼の件';
    var body = 'お世話になっております。\n\n';
    body += '以下の設備につきまして、見積もりをお願いしたく存じます。\n\n';
    body += '■ セルフィックス' + locName + '\n';
    body += '・設備: ' + eqName + '\n';
    body += '・作業内容: ' + workType + '\n\n';
    body += '--------------------------------------------------\n';
    body += '日商有田株式会社\n西村\n';
    body += '--------------------------------------------------';
    
    GmailApp.createDraft('', subject, body, {
      from: 'nishimura@selfix.jp'
    });
    
    return { success: true };
  } catch (e) {
    throw new Error('Gmail下書き作成エラー: ' + e.message);
  }
}

/**
 * 個別案件作成
 */
function createIndividualProject(locCode, eqId, locName, eqName, workType) {
  try {
    var config = getConfig();
    var scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    var uniqueId = Utilities.getUuid();
    
    scheduleSheet.appendRow([
      uniqueId,
      locCode,
      eqId,
      workType,
      '',
      config.PROJECT_STATUS.ESTIMATE_REQ,
      '',
      ''
    ]);
    
    return {
      success: true,
      projectId: uniqueId
    };
  } catch (e) {
    throw new Error('案件作成エラー: ' + e.message);
  }
}

/**
 * 一括発注のGmail下書き作成（汎用）
 */
function createBulkOrderGmailDraft(equipmentId) {
  try {
    var configs = getBulkOrderConfigs();
    var today = new Date();
    var targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    // ノズルカバーの場合
    if (equipmentId === 'PARTS-PUMP-1Y') {
      var targetStores = getNozzleCoverTargetStores();
      
      if (targetStores.length === 0) {
        return { success: false, message: '対象店舗がありません' };
      }
      
      var subject = '【見積依頼】見積り依頼の件';
      var body = 'お世話になっております。\n\n';
      body += targetYear + '年度のノズルカバー交換の発注をお願いいたします。\n\n';
      body += '【対象店舗: ' + targetStores.length + '店舗（全店）】\n\n';
      
      for (var i = 0; i < targetStores.length; i++) {
        body += '- セルフィックス' + targetStores[i].name + '\n';
      }
      
      body += '\n--------------------------------------------------\n';
      body += '日商有田株式会社\n西村\n';
      body += '--------------------------------------------------';
      
      GmailApp.createDraft('', subject, body, {
        from: 'nishimura@selfix.jp'
      });
      
      return {
        success: true,
        message: 'Gmail下書きを作成しました',
        subject: subject,
        recipient: 'タツノ宛て'
      };
    }
    
    // その他の一括発注
    var configItem = null;
    for (var i = 0; i < configs.length; i++) {
      if (configs[i].id === equipmentId) {
        configItem = configs[i];
        break;
      }
    }
    
    if (!configItem) {
      return { success: false, message: '設定が見つかりません' };
    }
    
    var targetStores = getBulkOrderTargetStores(configItem.id, configItem.cycle, configItem.searchKey);
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    // メール本文作成（店舗名に「セルフィックス」を付ける）
    var subject = '【見積依頼】見積り依頼の件';
    var body = 'お世話になっております。\n\n';
    body += targetYear + '年度の' + configItem.name + 'の発注をお願いいたします。\n\n';
    body += '【対象店舗: ' + targetStores.length + '店舗】\n';
    
    for (var i = 0; i < targetStores.length; i++) {
      var s = targetStores[i];
      body += '- セルフィックス' + s.name + '\n';
    }
    
    body += '\n--------------------------------------------------\n';
    body += '日商有田株式会社\n西村\n';
    body += '--------------------------------------------------';
    
    // Gmail下書き作成
    GmailApp.createDraft('', subject, body, {
      from: 'nishimura@selfix.jp'
    });
    
    return {
      success: true,
      message: 'Gmail下書きを作成しました',
      subject: subject,
      recipient: configItem.vendor + '宛て'
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * 一括発注の案件作成（汎用）
 */
function createBulkOrderProject(equipmentId) {
  try {
    var config = getConfig();
    var configs = getBulkOrderConfigs();
    var today = new Date();
    var targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    // ノズルカバーの場合
    if (equipmentId === 'PARTS-PUMP-1Y') {
      var targetStores = getNozzleCoverTargetStores();
      
      if (targetStores.length === 0) {
        return { success: false, message: '対象店舗がありません' };
      }
      
      var scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
      var uniqueId = Utilities.getUuid();
      
      scheduleSheet.appendRow([
        uniqueId,
        'BULK',
        'PARTS-PUMP-1Y',
        'ノズルカバー交換一括発注(' + targetYear + '年度)',
        '',
        config.PROJECT_STATUS.ESTIMATE_REQ,
        '',
        'タツノ'
      ]);
      
      return {
        success: true,
        projectId: uniqueId,
        equipmentName: 'ノズルカバー交換',
        targetCount: targetStores.length
      };
    }
    
    // その他の一括発注
    var configItem = null;
    for (var i = 0; i < configs.length; i++) {
      if (configs[i].id === equipmentId) {
        configItem = configs[i];
        break;
      }
    }
    
    if (!configItem) {
      return { success: false, message: '設定が見つかりません' };
    }
    
    var targetStores = getBulkOrderTargetStores(configItem.id, configItem.cycle, configItem.searchKey);
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    var scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    var uniqueId = Utilities.getUuid();
    
    // 案件作成（全店舗まとめて1案件）
    scheduleSheet.appendRow([
      uniqueId,
      'BULK', // 拠点コード（一括案件用）
      equipmentId,
      configItem.name + '一括発注(' + targetYear + '年度)',
      '', // 日程は後で入力
      config.PROJECT_STATUS.ESTIMATE_REQ,
      '',
      configItem.vendor
    ]);
    
    return {
      success: true,
      projectId: uniqueId,
      equipmentName: configItem.name,
      targetCount: targetStores.length
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * ノズルカバー交換のGmail下書き作成
 */
function createNozzleCoverGmailDraft() {
  try {
    var targetStores = getNozzleCoverTargetStores();
    var today = new Date();
    var currentMonth = today.getMonth() + 1;
    var fiscalYear = (currentMonth >= 1 && currentMonth <= 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    var subject = '【見積依頼】見積り依頼の件';
    var body = 'お世話になっております。\n\n';
    body += fiscalYear + '年度のノズルカバー交換の発注をお願いいたします。\n\n';
    body += '【対象店舗: ' + targetStores.length + '店舗（全店）】\n\n';
    
    for (var i = 0; i < targetStores.length; i++) {
      body += '- セルフィックス' + targetStores[i].name + '\n';
    }
    
    body += '\n--------------------------------------------------\n';
    body += '日商有田株式会社\n西村\n';
    body += '--------------------------------------------------';
    
    GmailApp.createDraft('', subject, body, {
      from: 'nishimura@selfix.jp'
    });
    
    return {
      success: true,
      message: 'Gmail下書きを作成しました',
      subject: subject,
      recipient: 'タツノ宛て'
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * ノズルカバー交換の案件作成
 */
function createNozzleCoverProject() {
  try {
    var config = getConfig();
    var targetStores = getNozzleCoverTargetStores();
    var today = new Date();
    var currentMonth = today.getMonth() + 1;
    var fiscalYear = (currentMonth >= 1 && currentMonth <= 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    var scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    var uniqueId = Utilities.getUuid();
    
    scheduleSheet.appendRow([
      uniqueId,
      'BULK',
      'PARTS-PUMP-1Y',
      'ノズルカバー交換一括発注(' + fiscalYear + '年度)',
      '',
      config.PROJECT_STATUS.ESTIMATE_REQ,
      '',
      'タツノ'
    ]);
    
    return {
      success: true,
      projectId: uniqueId,
      equipmentName: 'ノズルカバー交換',
      targetCount: targetStores.length
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
