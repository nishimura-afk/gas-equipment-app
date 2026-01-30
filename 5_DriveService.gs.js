/**
 * 5_DriveService.gs v5.7
 * ID生成にUUIDを採用
 */
const USER_DRIVE_ID = '1gUD2Z2N2-APYFXQcugu1fdex_YfoiyA2';
const INBOX_FOLDER_NAME = 'SS見積_受信BOX'; 
const ARCHIVE_FOLDER_NAME = 'SS見積_処理済';

function ensureInboxFolder() {
  // USER_DRIVE_IDが既に受信BOXフォルダ自体を指している
  const inboxFolder = DriveApp.getFolderById(USER_DRIVE_ID);
  
  // 処理済フォルダは親フォルダ内に作成（02_見積り内）
  const parentFolders = inboxFolder.getParents();
  if (parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    const archiveFolders = parentFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
    if (!archiveFolders.hasNext()) {
      parentFolder.createFolder(ARCHIVE_FOLDER_NAME);
    }
  }
  
  return { id: inboxFolder.getId(), url: inboxFolder.getUrl(), name: inboxFolder.getName() };
}

function scanInboxFiles() {
  const folderInfo = ensureInboxFolder();
  const folder = DriveApp.getFolderById(folderInfo.id);
  const files = folder.getFilesByType(MimeType.PDF);
  const result = [];
  const activeProjects = getAllActiveProjects();
  const allEquipments = getEquipmentListCached();

  while (files.hasNext()) {
    const file = files.next();
    const suggestion = suggestProjectFromFileName(file.getName(), activeProjects, allEquipments);
    result.push({
      id: file.getId(), name: file.getName(), url: file.getUrl(),
      lastUpdated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
      suggestion: suggestion
    });
  }
  return { files: result, folderUrl: folderInfo.url };
}

function scanDriveFolder(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const result = [];
    const activeProjects = getAllActiveProjects();
    const allEquipments = getEquipmentListCached();
    while (files.hasNext()) {
      const file = files.next();
      const suggestion = suggestProjectFromFileName(file.getName(), activeProjects, allEquipments);
      result.push({
        id: file.getId(), name: file.getName(), url: file.getUrl(),
        lastUpdated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy/MM/dd'),
        suggestion: suggestion
      });
    }
    return { success: true, files: result, folderName: folder.getName() };
  } catch (e) { return { success: false, error: 'フォルダが見つかりません' }; }
}

function suggestProjectFromFileName(fileName, projects, equipments) {
  const normalized = fileName.normalize('NFKC').toUpperCase();
  let bestMatch = null;
  let maxScore = 0;

  projects.forEach(p => {
    let score = 0;
    if (p.locName && normalized.includes(p.locName)) score += 10;
    if (p.equipmentName && normalized.includes(p.equipmentName)) score += 5;
    if (p.equipmentId && normalized.includes(p.equipmentId)) score += 8;
    if (score > maxScore && score >= 10) {
      maxScore = score;
      // 案件IDの短縮版を作成（表示用）
      const shortId = p.id ? p.id.substring(0, 8) : '';
      bestMatch = { 
        type: 'EXISTING', 
        id: p.id, 
        label: `案件#${shortId}: ${p.locName} - ${p.workType}`,
        detailLabel: `${p.locName} - ${p.equipmentName || p.equipmentId} - ${p.workType}`,
        locCode: p.locCode, 
        eqId: p.equipmentId,
        equipmentName: p.equipmentName || p.equipmentId
      };
    }
  });
  if (bestMatch) return bestMatch;

  equipments.forEach(e => {
    let score = 0;
    if (e['拠点名'] && normalized.includes(e['拠点名'])) score += 10;
    const keywords = ['タンク', '計量機', 'POS', '洗車機', 'LED', 'ポンプ'];
    const eqName = (e['設備名']||'').toUpperCase();
    if (eqName && normalized.includes(eqName)) score += 5;
    keywords.forEach(k => { if (eqName.includes(k) && normalized.includes(k)) score += 3; });
    if (score > maxScore && score >= 10) {
      maxScore = score;
      bestMatch = { type: 'NEW', locCode: e['拠点コード'], eqId: e['設備ID'], label: `【新規】${e['拠点名']}-${e['設備名']}`, locName: e['拠点名'], eqName: e['設備名'] };
    }
  });
  return bestMatch;
}

/**
 * 拠点一覧を取得
 */
function getStoreList() {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({ 
    code: row[0], 
    name: row[1] 
  }));
}

/**
 * 設備名を取得
 */
function getEquipmentName(locCode, eqId) {
  try {
    const list = getEquipmentListCached();
    const equipment = list.find(e => 
      e['拠点コード'] === locCode && e['設備ID'] === eqId
    );
    return equipment ? equipment['設備名'] : eqId;
  } catch (e) {
    return eqId;
  }
}

/**
 * ファイル名から拠点・設備を推測（軽量処理）
 */
function suggestFromFileName(fileName, locList) {
  if (!fileName) return null;
  
  const normalized = fileName.normalize('NFKC').toUpperCase();
  
  // 拠点名を検索
  let locCode = null;
  let locName = null;
  
  for (const loc of locList) {
    if (normalized.includes(loc.name.toUpperCase())) {
      locCode = loc.code;
      locName = loc.name;
      break;
    }
  }
  
  if (!locCode) return null;
  
  // 設備名を推測
  const equipmentKeywords = {
    'PUMP-G-01': ['ガソリン', '計量機', 'PUMP'],
    'PUMP-K-01': ['灯油', '計量機'],
    'POS-01': ['POS', 'レジ'],
    'TANK-01': ['タンク', '漏洩', '漏えい'],
    'LED-C-01': ['LED', 'キャノピー'],
    'PAINT-01': ['塗装'],
    'AC-01': ['エアコン', '空調'],
    'COMP-01': ['コンプレッサー', '圧縮機']
  };
  
  let eqId = null;
  let eqName = null;
  
  for (const [id, keywords] of Object.entries(equipmentKeywords)) {
    if (keywords.some(kw => normalized.includes(kw))) {
      eqId = id;
      eqName = getEquipmentName(locCode, id);
      break;
    }
  }
  
  return eqId ? { locCode, locName, eqId, eqName } : null;
}

function executeImport(filesToImport) {
  const config = getConfig();
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  
  let rootFolder;
  try { 
    rootFolder = DriveApp.getFolderById(USER_DRIVE_ID); 
  } catch(e) { 
    rootFolder = DriveApp.getRootFolder(); 
  }
  
  // 処理済フォルダの取得
  const parentFolders = rootFolder.getParents();
  let parentFolder;
  if (parentFolders.hasNext()) {
    parentFolder = parentFolders.next();
  } else {
    parentFolder = rootFolder;
  }
  
  const archiveFolders = parentFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
  const archiveFolder = archiveFolders.hasNext() 
    ? archiveFolders.next() 
    : parentFolder.createFolder(ARCHIVE_FOLDER_NAME);
  
  // 拠点マスタ取得
  const locList = getStoreList();
  const locMap = {};
  locList.forEach(l => locMap[l.code] = l.name);
  
  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  filesToImport.forEach((item, index) => {
    try {
      Logger.log(`[${index + 1}/${filesToImport.length}] Processing: ${item.fileId}`);
      
      const file = DriveApp.getFileById(item.fileId);
      const originalName = file.getName();
      
      // ========================================
      // 【重要】ファイル名ベースの軽量解析
      // （Gemini APIは使わない）
      // ========================================
      
      // 案件情報を整理
      let locCode = item.locCode;
      let locName = locMap[locCode] || locCode;
      let eqId = item.eqId;
      let eqName = '';
      
      // 設備名を取得
      if (locCode && eqId) {
        eqName = getEquipmentName(locCode, eqId);
      }
      
      // もしユーザーが「新規案件作成」を選ばなかった場合、
      // ファイル名から推測を試みる
      if (!locCode || !eqId) {
        const suggestion = suggestFromFileName(originalName, locList);
        if (suggestion) {
          locCode = suggestion.locCode;
          locName = suggestion.locName;
          eqId = suggestion.eqId;
          eqName = suggestion.eqName;
        }
      }
      
      // 見積り登録は外部の分類リネームGAS・見積りDB側で行う（当システムでは記録しない）
      
      // 新規案件の場合のみ案件作成
      if (item.projectType === 'NEW' && locCode && eqId) {
        const uniqueId = Utilities.getUuid();
        scheduleSheet.appendRow([
          uniqueId, 
          locCode, 
          eqId, 
          '見積受領(インポート)', 
          '', 
          config.PROJECT_STATUS.ESTIMATE_RCV, 
          '', 
          ''
        ]);
      }
      
      // 履歴に記録
      if (locCode && eqId) {
        getSheet(config.SHEET_NAMES.HISTORY).appendRow([
          locCode, 
          eqId, 
          '見積書登録', 
          new Date(), 
          `File: ${originalName}\nUrl: ${file.getUrl()}`, 
          ''
        ]);
      }
      
      // ========================================
      // 【重要】リネーム処理
      // ========================================
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
      let newName;
      
      if (locCode && eqId) {
        newName = `[${locCode}]${eqId}_見積_${timestamp}.pdf`;
      } else {
        // 情報が不足している場合は元のファイル名を保持
        newName = `未分類_${originalName}`;
      }
      
      file.setName(newName);
      Logger.log(`✅ Renamed: ${originalName} → ${newName}`);
      
      // ========================================
      // 【重要】フォルダ振り分け
      // ========================================
      let targetFolder;
      
      if (locCode && locName) {
        // 店舗フォルダに移動
        const subFolders = archiveFolder.getFoldersByName(locName);
        targetFolder = subFolders.hasNext() 
          ? subFolders.next() 
          : archiveFolder.createFolder(locName);
      } else {
        // 未分類フォルダに移動
        const unclassifiedFolders = archiveFolder.getFoldersByName('未分類');
        targetFolder = unclassifiedFolders.hasNext()
          ? unclassifiedFolders.next()
          : archiveFolder.createFolder('未分類');
      }
      
      file.moveTo(targetFolder);
      Logger.log(`✅ Moved to: ${targetFolder.getName()}`);
      
      successCount++;
      
    } catch (e) {
      errorCount++;
      errors.push(`${item.fileId}: ${e.message}`);
      Logger.log(`❌ Error processing ${item.fileId}: ${e.message}`);
    }
  });
  
  // 結果メッセージ
  let message = `✅ ${successCount}件処理完了`;
  
  if (errorCount > 0) {
    message += `\n⚠️ ${errorCount}件でエラーが発生しました`;
    Logger.log('Errors: ' + JSON.stringify(errors));
  }
  
  return { 
    success: true, 
    message: message,
    successCount: successCount,
    errorCount: errorCount,
    folderUrl: archiveFolder.getUrl() 
  };
}

function uploadAndImport(data, fileName, mimeType, projectInfo) {
  const folderInfo = ensureInboxFolder();
  const folder = DriveApp.getFolderById(folderInfo.id);
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
  const file = folder.createFile(blob);
  const importItem = { fileId: file.getId(), projectType: projectInfo.type, projectId: projectInfo.id, locCode: projectInfo.locCode, eqId: projectInfo.eqId };
  return executeImport([importItem]);
}