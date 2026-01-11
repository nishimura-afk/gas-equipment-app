/**
 * 3_DataUpdateService.gs v6.0
 * ベンダー自動振り分け・メール下書き作成
 * - Configからベンダー情報を取得し、Toアドレスを設定
 * - UUID使用
 */
function logSystemAction(actionType, detail, status = 'SUCCESS') {
  try {
    const config = getConfig();
    const sheet = getSheet(config.SHEET_NAMES.SYS_LOG);
    sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), actionType, detail, status]);
  } catch (e) {}
}

function recordExchangeComplete(locationCode, equipmentId, workType, workDate, subsidyInfo) {
  const config = getConfig();
  const masterSheet = getSheet(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const masterData = masterSheet.getDataRange().getValues();
  const rowIndex = masterData.findIndex(row => row[0] == locationCode && row[2] == equipmentId);
  if (rowIndex === -1) throw new Error('マスタに対象の設備が見つかりません');
 
  const headers = masterData[0];
  const rowRange = masterSheet.getRange(rowIndex + 1, 1, 1, headers.length);
  const rowValues = rowRange.getValues()[0];

  if (workType.includes('部品A') || workType.includes('消耗品')) rowValues[headers.indexOf('部品A交換日')] = workDate;
  else if (workType.includes('部品B') || workType.includes('メンテ')) rowValues[headers.indexOf('部品B最終交換日')] = workDate;
  else if (workType.includes('本体') || workType.includes('入替')) {
    rowValues[headers.indexOf('設置日(前回実施)')] = workDate;
    rowValues[headers.indexOf('部品A交換日')] = workDate;
    rowValues[headers.indexOf('部品B最終交換日')] = workDate;
  }
  rowValues[9] = ""; 
  rowRange.setValues([rowValues]);

  getSheet(config.SHEET_NAMES.HISTORY).appendRow([locationCode, equipmentId, workType, workDate, subsidyInfo || '-', '']);
  updateWebData();
  logSystemAction('UPDATE', `${locationCode}:${equipmentId} - ${workType}`);
  return { success: true };
}

function saveNextWorkMemo(shopCode, machineId, memo, spec) {
  const masterSheet = getSheet(getConfig().SHEET_NAMES.MASTER_EQUIPMENT);
  const data = masterSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] == shopCode && row[2] == machineId);
  if (rowIndex === -1) throw new Error('マスタが見つかりません');
  masterSheet.getRange(rowIndex + 1, 5).setValue(spec || "");
  masterSheet.getRange(rowIndex + 1, 10).setValue(memo || "");
  updateWebData();
  return { success: true };
}

function createVendorBatchDrafts() {
  const config = getConfig();
  const notices = getDashboardData().noticeList;
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  if (notices.length === 0) return { message: 'アラート対象はありません。' };

  const vendors = config.VENDORS;
  for (const key in vendors) { vendors[key].items = []; }

  notices.forEach(item => {
    const eqName = item['設備名'] || item['設備ID'];
    let assigned = false;
    for (const key in vendors) {
      if (key === 'OTHERS') continue;
      if (vendors[key].keywords.some(k => eqName.includes(k))) {
        vendors[key].items.push(item);
        assigned = true; break;
      }
    }
    if (!assigned) vendors['OTHERS'].items.push(item);
  });

  let log = [];
  for (const key in vendors) {
    const v = vendors[key];
    if (v.items.length === 0) continue;

    const subject = `【見積依頼】見積り依頼の件`;
    let body = `いつもお世話になっております。\n日商有田株式会社西村です。\n\n`;

    const isConsumableVendor = (key === 'SHARP' || key === 'TATSUNO');
    const hasConsumableWork = v.items.some(i => i['部品Aステータス'] !== '正常');

    if (isConsumableVendor && hasConsumableWork) {
        body += `以下の消耗品につきまして、交換をお願いいたします。\n`;
        body += `（価格に変更があればお知らせ下さい。）\n\n`;
    } else {
        body += `以下の設備につきまして、見積もりをお願いしたく存じます。\n\n`;
    }

    v.items.forEach(i => {
      let eqDisplayName = i['設備名'];
      if (eqDisplayName.includes('釣銭機カバー')) eqDisplayName = eqDisplayName.replace('釣銭機カバー', '投入/取出し口のプラスチックカバー');
      if (eqDisplayName.includes('パネル')) eqDisplayName = eqDisplayName.replace('パネル', 'タッチパネル');

      body += `■ ${i['拠点名']}\n`;
      body += `・設備: ${eqDisplayName}\n`;
      if (i['spec']) body += `・型式: ${i['spec']}\n`;
      if (i['nextWorkMemo']) body += `・備考: ${i['nextWorkMemo']}\n`;
      body += `\n`;

      const uniqueId = Utilities.getUuid();
      scheduleSheet.appendRow([uniqueId, i['拠点コード'], i['設備ID'], '発注', '', config.PROJECT_STATUS.ORDERED, '', v.name]);
    });

    body += `--------------------------------------------------\n日商有田株式会社\n西村\n--------------------------------------------------`;

    GmailApp.createDraft(v.email || '', subject, body, {
      from: 'nishimura@selfix.jp'
    });
    log.push(v.name);
  }
  return { message: `${log.join(', ')} の下書きを作成しました。` };
}

function createAlertDrafts() {
  const config = getConfig();
  const notices = getDashboardData().noticeList;
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  if (notices.length === 0) return { message: 'アラート対象はありません。' };

  let count = 0;
  notices.forEach(item => {
    let eqDisplayName = item['設備名'];
    if (eqDisplayName.includes('釣銭機カバー')) eqDisplayName = eqDisplayName.replace('釣銭機カバー', '投入/取出し口のプラスチックカバー');
    if (eqDisplayName.includes('パネル')) eqDisplayName = eqDisplayName.replace('パネル', 'タッチパネル');

    let workType = 'メンテナンス';
    if (item['本体ステータス'] !== '正常') workType = '本体更新・入替';
    else if (item['部品Aステータス'] !== '正常') workType = '消耗品交換';

    const subject = `【見積依頼】見積り依頼の件`;
    const body = `いつもお世話になっております。\n日商有田株式会社西村です。\n\n` +
                 `以下の設備につきまして、見積もりをお願いしたく存じます。\n\n` +
                 `■ ${item['拠点名']}\n` +
                 `・対象設備: ${eqDisplayName}\n` +
                 `・型式: ${item['spec'] || '不明'}\n\n` +
                 `--------------------------------------------------\n日商有田株式会社\n西村\n--------------------------------------------------`;
    
    GmailApp.createDraft('', subject, body, {
      from: 'nishimura@selfix.jp'
    });
    const uniqueId = Utilities.getUuid();
    scheduleSheet.appendRow([uniqueId, item['拠点コード'], item['設備ID'], workType, '', config.PROJECT_STATUS.ESTIMATE_REQ, '', '']);
    count++;
  });
  return { message: `${count}件の下書きを作成しました。` };
}

function createSingleDraftAndProject(locName, locCode, eqName, eqId, workType, body) {
  const config = getConfig();
  const subject = `【見積依頼】見積り依頼の件`;
  GmailApp.createDraft('', subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  const uniqueId = Utilities.getUuid();
  getSheet(config.SHEET_NAMES.SCHEDULE).appendRow([uniqueId, locCode, eqId, workType, '', config.PROJECT_STATUS.ESTIMATE_REQ, '', '']);
  return { success: true };
}

/**
 * ベンダー別にアラートをグループ化
 */
function getVendorGroupedAlerts() {
  const config = getConfig();
  const notices = getDashboardData().noticeList;
  const vendors = config.VENDORS;
  
  // 初期化
  for (const key in vendors) {
    vendors[key].items = [];
  }
  
  // 振り分け
  notices.forEach(item => {
    const eqName = item['設備名'] || item['設備ID'];
    let assigned = false;
    
    for (const key in vendors) {
      if (key === 'OTHERS') continue;
      if (vendors[key].keywords.some(k => eqName.includes(k))) {
        vendors[key].items.push(item);
        assigned = true;
        break;
      }
    }
    
    if (!assigned) vendors['OTHERS'].items.push(item);
  });
  
  return vendors;
}

/**
 * 個別案件のGmail下書きと案件作成
 */
function createIndividualDraftAndProject(locCode, locName, eqId, eqName, workType, spec) {
  const config = getConfig();
  
  // メール本文作成
  const subject = '【見積依頼】見積もり依頼の件';
  let body = 'いつもお世話になっております。\n日商有田株式会社西村です。\n\n';
  body += '以下の設備につきまして、見積もりをお願いしたく存じます。\n\n';
  body += `■ ${locName}\n`;
  body += `・対象設備: ${eqName}\n`;
  if (spec) body += `・型式: ${spec}\n`;
  body += `・作業内容: ${workType}\n\n`;
  body += '--------------------------------------------------\n';
  body += '日商有田株式会社\n西村\n';
  body += '--------------------------------------------------';
  
  // Gmail下書き作成
  GmailApp.createDraft('', subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  // 案件シートに登録
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
  
  return { success: true, message: 'Gmail下書きを作成し、案件を登録しました' };
}

/**
 * 選択された案件のみでベンダー別下書きを作成
 */
function createVendorDraftForSelected(vendorKey, selectedItemIds) {
  const config = getConfig();
  const notices = getDashboardData().noticeList;
  const vendor = config.VENDORS[vendorKey];
  
  if (!vendor) throw new Error('ベンダーが見つかりません');
  
  // 選択されたアイテムのみフィルタ
  const selectedItems = notices.filter(item => {
    const itemId = `${item['拠点コード']}_${item['設備ID']}`;
    return selectedItemIds.includes(itemId);
  });
  
  if (selectedItems.length === 0) {
    return { message: '対象案件がありません' };
  }
  
  // メール本文作成
  const subject = '【見積依頼】見積り依頼の件';
  let body = 'いつもお世話になっております。\n日商有田株式会社西村です。\n\n';
  body += '以下の設備につきまして、見積もりをお願いしたく存じます。\n\n';
  
  selectedItems.forEach(item => {
    let eqDisplayName = item['設備名'];
    if (eqDisplayName.includes('釣銭機カバー')) {
      eqDisplayName = eqDisplayName.replace('釣銭機カバー', '投入/取出し口のプラスチックカバー');
    }
    if (eqDisplayName.includes('パネル')) {
      eqDisplayName = eqDisplayName.replace('パネル', 'タッチパネル');
    }
    
    body += `■ ${item['拠点名']}\n`;
    body += `・設備: ${eqDisplayName}\n`;
    if (item['spec']) body += `・型式: ${item['spec']}\n`;
    if (item['nextWorkMemo']) body += `・備考: ${item['nextWorkMemo']}\n`;
    body += '\n';
  });
  
  body += '--------------------------------------------------\n';
  body += '日商有田株式会社\n西村\n';
  body += '--------------------------------------------------';
  
  // Gmail下書き作成
  GmailApp.createDraft(vendor.email || '', subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  // 案件シートに登録
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  selectedItems.forEach(item => {
    const uniqueId = Utilities.getUuid();
    scheduleSheet.appendRow([
      uniqueId,
      item['拠点コード'],
      item['設備ID'],
      '見積依頼',
      '',
      config.PROJECT_STATUS.ESTIMATE_REQ,
      '',
      vendor.name
    ]);
  });
  
  return {
    message: `${vendor.name}宛てに${selectedItems.length}件の下書きを作成しました`
  };
}

/**
 * 複数ベンダーの選択案件をまとめて処理
 */
function createMultipleVendorDrafts(vendorGroups) {
  const results = [];
  
  for (const vendorKey in vendorGroups) {
    const itemIds = vendorGroups[vendorKey];
    const result = createVendorDraftForSelected(vendorKey, itemIds);
    results.push(result.message);
  }
  
  return { message: results.join('\n') };
}

/**
 * カスタムGmail下書き作成
 */
function createCustomGmailDraft(to, subject, body) {
  GmailApp.createDraft(to, subject, body, {
    from: 'nishimura@selfix.jp'
  });
  
  return { success: true, message: 'Gmail下書きを作成しました' };
}