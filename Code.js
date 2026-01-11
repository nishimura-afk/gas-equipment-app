/**
 * Code.gs v6.0
 * Webアプリのエントリーポイント & 不足関数の実装
 */
function doGet() {
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
  // 完了・取消以外の進行中案件IDリスト
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
  
  // 拠点マスタから拠点名を取得するためのマップ
  const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
  const locData = locSheet.getDataRange().getValues();
  const locMap = {};
  locData.slice(1).forEach(r => { if(r[0]) locMap[r[0]] = r[1]; });

  // 設備名を取得するためのマップ
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
      locName: locMap[locCode] || locCode, // 拠点名を付与
      equipmentId: eqId,
      equipmentName: eqMap[key] || eqId,   // 設備名を付与
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

// ★追加実装: 案件取り消し
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

// ★追加実装: 日程登録＆カレンダー連携
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

  return `いつもお世話になっております。\n日商有田株式会社西村です。\n\n` +
         `以下の設備につきまして、見積もりをお願いしたく存じます。\n\n` +
         `■ ${locName}\n` +
         `・対象設備: ${displayEqName}\n` +
         `\n` + 
         `--------------------------------------------------\n日商有田株式会社\n西村\n--------------------------------------------------`;
}

// --- Code.gs の末尾に追記 ---

// 強制的にダミーデータを返すデバッグ用関数
function getNozzleCoverInfo() {
  Logger.log('=== Code.gs: getNozzleCoverInfo FORCED DEBUG ===');
  return {
    hasAlert: true,
    targetCount: 999,
    targetYear: 2026,
    emailDraft: "これはCode.gsから強制的に返されたデバッグデータです。",
    config: {
      id: "PARTS-PUMP-1Y",
      name: "ノズルカバー（デバッグ）",
      emoji: "🐞",
      vendor: "タツノ"
    },
    targetStores: [{name: "デバッグ店A"}, {name: "デバッグ店B"}]
  };
}

function getAllBulkOrderInfo() {
  Logger.log('=== Code.gs: getAllBulkOrderInfo FORCED DEBUG ===');
  return [
    {
      hasAlert: true,
      targetCount: 123,
      targetYear: 2026,
      emailDraft: "これはCode.gsから強制的に返されたデバッグデータです。",
      config: {
        id: "DEBUG-BULK",
        name: "一括発注（デバッグ）",
        emoji: "🐛",
        vendor: "タツノ"
      },
      targetStores: [{name: "デバッグ店1"}]
    }
  ];
}