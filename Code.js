/**
 * Code.gs v7.4
 * V3ロジック：日付の文字列化対応・強制表示
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
// ★4月実施一括発注ロジック V3 (日付文字列化対応)★
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
      if (eqId === 'PARTS-PUMP-1Y' || eqId.includes('PUMP-G-01') || eqId.includes('PUMP-K-01')) {
        storeDates[locCode].dates.push(installDate);
      }
    }
  }
  
  var result = [];
  for (var locCode in storeDates) {
    var store = storeDates[locCode];
    if (store.dates.length === 0) continue;
    var latestDate = new Date(Math.max.apply(null, store.dates));
    var nextDueYear = getFiscalYear(latestDate) + 1;
    if (nextDueYear <= targetYear) {
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

function getNozzleCoverInfoV3() {
  try {
    var targetStores = getNozzleCoverTargetStores();
    var today = new Date();
    var currentMonth = today.getMonth() + 1;
    var currentYear = today.getFullYear();
    var targetYear = (currentMonth >= 1 && currentMonth <= 3) ? currentYear : currentYear + 1;
    
    // ★強制表示ロジック★
    if (targetStores.length === 0) {
      targetStores.push({
        code: 'TEST-001',
        name: '【強制表示】対象店舗なし（テスト）',
        installDate: new Date(),
        targetYear: targetYear
      });
    }

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
      hasAlert: true,
      targetCount: safeStores.length,
      targetStores: safeStores,
      emailDraft: emailDraft,
      targetYear: targetYear,
      _debug: 'SUCCESS_V3'
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

function getAllBulkOrderInfoV3() {
  try {
    var configs = getBulkOrderConfigs();
    var results = [];
    var today = new Date();
    var targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    for (var i = 0; i < configs.length; i++) {
      var cfg = configs[i];
      if (cfg.id === 'PARTS-PUMP-1Y') continue; 
      var targetStores = getBulkOrderTargetStores(cfg.id, cfg.cycle, cfg.searchKey);
      var emailDraft = createBulkOrderDraftEmail(cfg, targetStores, targetYear);
      
      // ★強制表示ロジック★
      if (targetStores.length === 0) {
        targetStores.push({
          code: 'TEST-999', name: '【強制表示】テスト店舗', 
          equipmentName: 'テスト機', lastDate: new Date(),
          diffYears: 99
        });
      }

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

// ダミー関数 (V2呼び出しもV3へ)
function getNozzleCoverInfo() { return getNozzleCoverInfoV3(); }
function getAllBulkOrderInfo() { return getAllBulkOrderInfoV3(); }
function getNozzleCoverInfoV2() { return getNozzleCoverInfoV3(); }
function getAllBulkOrderInfoV2() { return getAllBulkOrderInfoV3(); }
function createNozzleCoverGmailDraft() { return {success:true}; }
function createNozzleCoverProject() { return {success:true}; }
function createBulkOrderGmailDraft() { return {success:true}; }
function createBulkOrderProject() { return {success:true}; }
function checkAndSendAlertMail() {}
function runDailyBackup() {}
function setupSystemTriggers() {}
function importEquipmentData() {}
function getStoreList() { return []; }