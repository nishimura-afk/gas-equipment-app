/**
 * Code.gs v8.0
 * 本番運用版
 * - 標準レスポンス形式対応
 * - 共通ユーティリティ関数使用
 * - 列インデックス定数使用
 */
function doGet() {
  logDebug('doGet START v8.0');
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

  // 本体入れ替え案件がある設備の消耗品アラートを除外
  const bodyReplacementProjects = scheduleData.slice(1)
    .filter(row => 
      row[5] !== config.PROJECT_STATUS.COMPLETED && 
      row[5] !== config.PROJECT_STATUS.CANCELLED &&
      (row[3].includes('本体') || row[3].includes('入替') || row[3].includes('更新'))
    )
    .map(row => `${row[1]}_${row[2]}`);

  // 本体更新が案件化されている拠点を抽出（定期部品交換の除外用）
  const bodyReplacements = getBodyReplacementLocations(scheduleData, config);
  const gasBodyReplacementLocations = bodyReplacements.gasLocations;
  const keroseneBodyReplacementLocations = bodyReplacements.keroseneLocations;

  const notices = data.filter(m => {
    const equipmentKey = `${m['拠点コード']}_${m['設備ID']}`;
    
    // 既に案件化されているものは除外
    if (ignoreActions.includes(equipmentKey)) return false;
    
    // ガソリン計量機部品(4年)の除外判定
    if (m['設備ID'] === 'PARTS-PUMP-4Y' && gasBodyReplacementLocations.has(m['拠点コード'])) {
      return false;
    }
    
    // 灯油パネル更新の除外判定
    if (m['設備ID'] === 'PARTS-K-PANEL-7Y' && keroseneBodyReplacementLocations.has(m['拠点コード'])) {
      return false;
    }
    
    // 本体入れ替え案件がある場合、消耗品アラートは除外
    if (bodyReplacementProjects.includes(equipmentKey) && 
        m['部品Aステータス'] !== config.STATUS.NORMAL) {
      return false;
    }
    
    // その他のアラート判定
    return m['本体ステータス'] !== config.STATUS.NORMAL || 
           m['部品Aステータス'] !== config.STATUS.NORMAL || 
           (m['部品Bステータス'] && m['部品Bステータス'] !== config.STATUS.NORMAL);
  });
  return { noticeCount: notices.length, normalCount: data.length - notices.length, noticeList: notices };
}

function getAllActiveProjects() {
  const config = getConfig();
  const data = getSheet(config.SHEET_NAMES.SCHEDULE).getDataRange().getValues();
  if (data.length <= 1) return [];
  const locMap = buildLocationMap();
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
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.STATUS).setValue(newStatus);
      return successResponse({ id: id, newStatus: newStatus });
    }
  }
  return errorResponse('指定されたIDが見つかりません');
}

function cancelProject(id) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.STATUS).setValue(config.PROJECT_STATUS.CANCELLED);
      return successResponse({ id: id, status: config.PROJECT_STATUS.CANCELLED });
    }
  }
  return errorResponse('指定されたIDが見つかりません');
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
    { id: 'PARTS-PUMP-4Y', name: 'ガソリン計量機部品(4年)', cycle: 4, vendor: 'タツノ', emoji: '⛽', searchKey: 'ガソリン計量機部品' }
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
  
  // 案件管理シートから既に案件化済みの拠点を取得
  var scheduleSheet = ss.getSheetByName(config.SHEET_NAMES.SCHEDULE);
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var existingProjects = new Set();
  scheduleData.slice(1).forEach(function(row) {
    if (row[5] !== config.PROJECT_STATUS.COMPLETED && 
        row[5] !== config.PROJECT_STATUS.CANCELLED &&
        row[2] === 'PARTS-PUMP-1Y') {
      existingProjects.add(row[1]); // 拠点コードを追加
    }
  });
  
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
    // 既に案件化済みの店舗は除外
    if (!existingProjects.has(store.code)) {
      result.push({ code: store.code, name: store.name, installDate: latestDate, targetYear: targetYear });
    }
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
  for (var i = 0; i < targetStores.length; i++) { 
    body += '- セルフィックス' + targetStores[i].name + '\n'; 
  }
  body += '\n--------------------------------------------------\n日商有田株式会社\n西村\n--------------------------------------------------';
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
  
  // 案件管理シートから本体更新が案件化されている拠点を取得
  var scheduleSheet = ss.getSheetByName(config.SHEET_NAMES.SCHEDULE);
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var bodyReplacements = getBodyReplacementLocations(scheduleData, config);
  var gasBodyReplacementLocations = bodyReplacements.gasLocations;
  var keroseneBodyReplacementLocations = bodyReplacements.keroseneLocations;
  
  // この設備IDで既に案件化済みの拠点を取得
  var existingProjects = new Set();
  scheduleData.slice(1).forEach(function(row) {
    if (row[5] !== config.PROJECT_STATUS.COMPLETED && 
        row[5] !== config.PROJECT_STATUS.CANCELLED &&
        row[2] === equipmentId) {
      existingProjects.add(row[1]); // 拠点コードを追加
    }
  });
  
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
    
    // ガソリン計量機部品(4年)の除外判定
    if (eqId === 'PARTS-PUMP-4Y' && gasBodyReplacementLocations.has(locCode)) {
      continue;
    }
    
    // 灯油パネル更新の除外判定
    if (eqId === 'PARTS-K-PANEL-7Y' && keroseneBodyReplacementLocations.has(locCode)) {
      continue;
    }
    
    var isMatch = (eqId.indexOf(equipmentId) >= 0) || (searchKey && eqName.indexOf(searchKey) >= 0);
    
    if (isMatch && installDate instanceof Date && !isNaN(installDate.getTime())) {
      var baseDate = (partADate instanceof Date && !isNaN(partADate.getTime())) ? partADate : installDate;
      var installFY = getFiscalYear(baseDate);
      var targetFY = targetYear;
      var diffYears = targetFY - installFY;
      
      if (diffYears > cycleYears && !storeMap[locCode] && !existingProjects.has(locCode)) {
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
    body += '- セルフィックス' + s.name + '\n';
  }
  body += '\n--------------------------------------------------\n日商有田株式会社\n西村\n--------------------------------------------------';
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
    const subject = '【見積依頼】見積り依頼の件';
    let body = 'お世話になっております。\n\n';
    body += '以下の設備につきまして、見積もりをお願いしたく存じます。\n\n';
    body += `■ セルフィックス${locName}\n`;
    body += `・設備: ${eqName}\n`;
    body += `・作業内容: ${workType}\n\n`;
    body += '--------------------------------------------------\n';
    body += '日商有田株式会社\n西村\n';
    body += '--------------------------------------------------';
    
    GmailApp.createDraft('', subject, body, {
      from: getConfig().ADMIN_MAIL
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
    const config = getConfig();
    const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    const uniqueId = Utilities.getUuid();
    
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
 * 一括発注用Gmail下書き作成
 */
function createBulkOrderGmailDraft(equipmentId) {
  try {
    const today = new Date();
    const targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    // ノズルカバーの場合
    if (equipmentId === 'PARTS-PUMP-1Y') {
      const targetStores = getNozzleCoverTargetStores();
      
      if (targetStores.length === 0) {
        return { success: false, message: '対象店舗がありません' };
      }
      
      const subject = '【見積依頼】見積り依頼の件';
      let body = 'お世話になっております。\n\n';
      body += targetYear + '年度のノズルカバー交換の発注をお願いいたします。\n\n';
      body += '【対象店舗: ' + targetStores.length + '店舗（全店）】\n\n';
      
      for (let i = 0; i < targetStores.length; i++) {
        body += '- セルフィックス' + targetStores[i].name + '\n';
      }
      
      body += '\n--------------------------------------------------\n';
      body += '日商有田株式会社\n西村\n';
      body += '--------------------------------------------------';
      
      GmailApp.createDraft('', subject, body, {
        from: getConfig().ADMIN_MAIL
      });

      return {
        success: true,
        message: 'Gmail下書きを作成しました',
        subject: subject,
        recipient: 'タツノ宛て'
      };
    }
    
    // その他の一括発注
    const configs = getBulkOrderConfigs();
    const configItem = configs.find(c => c.id === equipmentId);
    if (!configItem) {
      return { success: false, message: '設定が見つかりません' };
    }
    
    const targetStores = getBulkOrderTargetStores(configItem.id, configItem.cycle, configItem.searchKey);
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    const subject = '【見積依頼】見積り依頼の件';
    let body = 'お世話になっております。\n\n';
    body += targetYear + '年度の' + configItem.name + 'の発注をお願いいたします。\n\n';
    body += '【対象店舗: ' + targetStores.length + '店舗】\n';
    
    for (let i = 0; i < targetStores.length; i++) {
      const s = targetStores[i];
      body += '- セルフィックス' + s.name + '\n';
    }
    
    body += '\n--------------------------------------------------\n';
    body += '日商有田株式会社\n西村\n';
    body += '--------------------------------------------------';
    
    GmailApp.createDraft('', subject, body, {
      from: getConfig().ADMIN_MAIL
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
 * 一括発注用案件作成（店舗ごとに個別案件を作成）
 */
function createBulkOrderProject(equipmentId) {
  try {
    const config = getConfig();
    const today = new Date();
    const targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
    
    // ノズルカバーの場合
    if (equipmentId === 'PARTS-PUMP-1Y') {
      return createNozzleCoverProject();
    }
    
    // その他の一括発注
    const configs = getBulkOrderConfigs();
    const configItem = configs.find(c => c.id === equipmentId);
    if (!configItem) {
      return { success: false, message: '設定が見つかりません' };
    }
    
    const targetStores = getBulkOrderTargetStores(configItem.id, configItem.cycle, configItem.searchKey);
    
    if (targetStores.length === 0) {
      return { success: false, message: '対象店舗がありません' };
    }
    
    const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    
    // 対象店舗ごとに個別案件を作成
    targetStores.forEach(store => {
      const uniqueId = Utilities.getUuid();
      scheduleSheet.appendRow([
        uniqueId,
        store.code,
        equipmentId,
        configItem.name + '(' + targetYear + '年度)',
        '',
        config.PROJECT_STATUS.ESTIMATE_REQ,
        '',
        configItem.vendor
      ]);
    });
    
    return {
      success: true,
      equipmentName: configItem.name,
      targetCount: targetStores.length
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * ノズルカバー交換の案件を店舗ごとに作成
 */
function createNozzleCoverProject() {
  const config = getConfig();
  const targetStores = getNozzleCoverTargetStores();
  
  if (targetStores.length === 0) {
    return { success: false, message: '対象店舗がありません' };
  }
  
  const today = new Date();
  const targetYear = (today.getMonth() < 3) ? today.getFullYear() : today.getFullYear() + 1;
  
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  
  // 対象店舗ごとに個別案件を作成
  targetStores.forEach(store => {
    const uniqueId = Utilities.getUuid();
    scheduleSheet.appendRow([
      uniqueId,
      store.code,
      'PARTS-PUMP-1Y',
      'ノズルカバー交換(' + targetYear + '年度)',
      '',
      config.PROJECT_STATUS.ESTIMATE_REQ,
      '',
      'タツノ'
    ]);
  });
  
  return {
    success: true,
    equipmentName: 'ノズルカバー交換',
    targetCount: targetStores.length
  };
}
