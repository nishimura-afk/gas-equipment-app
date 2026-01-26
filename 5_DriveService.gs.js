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
      bestMatch = { type: 'EXISTING', id: p.id, label: `【既存】${p.locName}-${p.workType}`, locCode: p.locCode, eqId: p.equipmentId };
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

function executeImport(filesToImport) {
  const config = getConfig();
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  filesToImport.forEach(item => {
    try {
      const file = DriveApp.getFileById(item.fileId);
      
      // 1. PDF抽出（Gemini API）
      Logger.log('PDF抽出開始: ' + file.getName());
      const extractResult = extractEstimateFromPDF(item.fileId);
      
      if (!extractResult || !extractResult.success) {
        Logger.log('PDF抽出失敗: ' + (extractResult ? extractResult.message : '不明なエラー'));
        errors.push(file.getName() + ': 抽出失敗');
        errorCount++;
        return;
      }
      
      // 2. 案件情報を構築
      let projectInfo = null;
      if (item.projectType === 'NONE') {
        projectInfo = { type: 'NONE', id: null, locCode: null, eqId: null, locName: null, eqName: null };
      } else if (item.projectType === 'NEW') {
        // 新規案件を作成
        const uniqueId = Utilities.getUuid();
        scheduleSheet.appendRow([uniqueId, item.locCode, item.eqId, '見積受領(インポート)', '', config.PROJECT_STATUS.ESTIMATE_RCV, '', '']);
        
        const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
        const locData = locSheet.getDataRange().getValues();
        const locRow = locData.find(row => row[0] === item.locCode);
        const locName = locRow ? locRow[1] : item.locCode;
        
        const equipmentList = getEquipmentListCached();
        const eqRow = equipmentList.find(row => row['拠点コード'] === item.locCode && row['設備ID'] === item.eqId);
        const eqName = eqRow ? eqRow['設備名'] : item.eqId;
        
        projectInfo = { type: 'NEW', id: uniqueId, locCode: item.locCode, eqId: item.eqId, locName: locName, eqName: eqName };
      } else if (item.projectType === 'EXISTING') {
        // 既存案件の情報を取得
        const scheduleData = scheduleSheet.getDataRange().getValues();
        const projectRow = scheduleData.find(row => row[0] === item.projectId);
        if (projectRow) {
          const locCode = projectRow[1];
          const eqId = projectRow[2];
          
          const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
          const locData = locSheet.getDataRange().getValues();
          const locRow = locData.find(row => row[0] === locCode);
          const locName = locRow ? locRow[1] : locCode;
          
          const equipmentList = getEquipmentListCached();
          const eqRow = equipmentList.find(row => row['拠点コード'] === locCode && row['設備ID'] === eqId);
          const eqName = eqRow ? eqRow['設備名'] : eqId;
          
          projectInfo = { type: 'EXISTING', id: item.projectId, locCode: locCode, eqId: eqId, locName: locName, eqName: eqName };
        }
      }
      
      // 3. 案件ありの場合のみスプレッドシート保存
      if (projectInfo && projectInfo.type !== 'NONE') {
        Logger.log('スプレッドシート保存開始');
        const fileInfo = {
          id: item.fileId,
          name: file.getName(),
          url: file.getUrl()
        };
        
        try {
          saveEstimateToSheet(extractResult, fileInfo, projectInfo);
          Logger.log('スプレッドシート保存完了');
        } catch (e) {
          Logger.log('スプレッドシート保存エラー: ' + e.message);
          errors.push(file.getName() + ': 保存エラー - ' + e.message);
        }
      } else {
        Logger.log('案件なしのため、スプレッドシート保存をスキップ');
      }
      
      // 4. PDFをリネームして処理済フォルダに移動（案件の有無に関わらず実行）
      Logger.log('PDF移動開始');
      const moveResult = moveAndRenameEstimatePDF(item.fileId, extractResult, projectInfo);
      Logger.log('PDF移動完了: ' + moveResult.newName);
      
      // 5. 履歴に記録
      const locCode = projectInfo ? projectInfo.locCode : '';
      const eqId = projectInfo ? projectInfo.eqId : '';
      getSheet(config.SHEET_NAMES.HISTORY).appendRow([
        locCode, 
        eqId, 
        '見積書登録', 
        new Date(), 
        `File: ${moveResult.newName}\nUrl: ${moveResult.newUrl}`, 
        ''
      ]);
      
      successCount++;
    } catch (e) {
      Logger.log('Import Error: ' + e.message);
      Logger.log(e.stack);
      errors.push(DriveApp.getFileById(item.fileId).getName() + ': ' + e.message);
      errorCount++;
    }
  });
  
  let message = `${successCount}件処理完了`;
  if (errorCount > 0) {
    message += `（${errorCount}件エラー）`;
  }
  
  return { 
    success: true, 
    message: message,
    errors: errors.length > 0 ? errors : null
  };
}

/**
 * フォルダを取得または作成
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

/**
 * ファイル名から設備名を抽出
 */
function extractEquipmentNameFromFileName(fileName) {
  if (!fileName) return '設備';
  
  // キーワードを検索
  const keywords = ['トイレ', 'ポンプ', 'タンク', '計量機', '計量器', 'POS', 'LED', '洗車', '釣銭機', 'エアコン', '照明'];
  const normalized = fileName.normalize('NFKC');
  
  for (let i = 0; i < keywords.length; i++) {
    if (normalized.includes(keywords[i])) {
      return keywords[i];
    }
  }
  
  return '設備';
}

/**
 * 見積PDFをリネームして処理済フォルダに移動
 * @param {string} fileId - PDFファイルのID
 * @param {Object} estimateData - 抽出データ {success, data}
 * @param {Object} projectInfo - 案件情報（なしの場合はnull）
 * @return {Object} {newName, newUrl}
 */
function moveAndRenameEstimatePDF(fileId, estimateData, projectInfo) {
  try {
    const file = DriveApp.getFileById(fileId);
    const data = estimateData.data;
    
    // 処理済フォルダを取得
    const inboxFolder = DriveApp.getFolderById(USER_DRIVE_ID);
    const parentFolders = inboxFolder.getParents();
    if (!parentFolders.hasNext()) {
      throw new Error('親フォルダが見つかりません');
    }
    const parentFolder = parentFolders.next();
    const archiveFolders = parentFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
    const archiveFolder = archiveFolders.hasNext() ? archiveFolders.next() : parentFolder.createFolder(ARCHIVE_FOLDER_NAME);
    
    // 年月フォルダを取得または作成
    const estimateDate = data.estimateDate ? new Date(data.estimateDate) : new Date();
    const year = estimateDate.getFullYear() + '年';
    const month = String(estimateDate.getMonth() + 1).padStart(2, '0') + '月';
    
    const yearFolder = getOrCreateFolder(archiveFolder, year);
    const monthFolder = getOrCreateFolder(yearFolder, month);
    
    // ファイル名を生成
    let locationName = '';
    let equipmentName = '';
    
    if (projectInfo && projectInfo.type !== 'NONE') {
      // 案件情報から取得
      if (projectInfo.type === 'EXISTING') {
        const config = getConfig();
        const scheduleData = getSheet(config.SHEET_NAMES.SCHEDULE).getDataRange().getValues();
        const projectRow = scheduleData.find(row => row[0] === projectInfo.id);
        if (projectRow) {
          const locCode = projectRow[1];
          const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
          const locData = locSheet.getDataRange().getValues();
          const locRow = locData.find(row => row[0] === locCode);
          locationName = locRow ? locRow[1] : locCode;
          
          const eqId = projectRow[2];
          const equipmentList = getEquipmentListCached();
          const eqRow = equipmentList.find(row => row['拠点コード'] === locCode && row['設備ID'] === eqId);
          equipmentName = eqRow ? eqRow['設備名'] : eqId;
        }
      } else if (projectInfo.type === 'NEW') {
        const config = getConfig();
        const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
        const locData = locSheet.getDataRange().getValues();
        const locRow = locData.find(row => row[0] === projectInfo.locCode);
        locationName = locRow ? locRow[1] : projectInfo.locCode;
        
        const equipmentList = getEquipmentListCached();
        const eqRow = equipmentList.find(row => row['拠点コード'] === projectInfo.locCode && row['設備ID'] === projectInfo.eqId);
        equipmentName = eqRow ? eqRow['設備名'] : projectInfo.eqId;
      }
    }
    
    // 案件情報がない場合はファイル名から抽出
    if (!locationName) {
      locationName = extractLocationNameFromFileName(file.getName());
    }
    if (!equipmentName) {
      equipmentName = extractEquipmentNameFromFileName(file.getName());
    }
    
    // デフォルト値
    if (!locationName) locationName = '不明';
    if (!equipmentName) equipmentName = '設備';
    
    const vendor = data.vendor || '不明';
    const dateStr = data.estimateDate ? data.estimateDate.replace(/-/g, '') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    
    const newName = `${locationName}_${equipmentName}_${vendor}_見積_${dateStr}.pdf`;
    
    // ファイルを移動してリネーム
    file.moveTo(monthFolder);
    file.setName(newName);
    
    return {
      newName: newName,
      newUrl: file.getUrl()
    };
  } catch (e) {
    Logger.log('PDF移動エラー: ' + e.message);
    throw e;
  }
}

function uploadAndImport(data, fileName, mimeType, projectInfo) {
  const folderInfo = ensureInboxFolder();
  const folder = DriveApp.getFolderById(folderInfo.id);
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
  const file = folder.createFile(blob);
  const importItem = { fileId: file.getId(), projectType: projectInfo.type, projectId: projectInfo.id, locCode: projectInfo.locCode, eqId: projectInfo.eqId };
  return executeImport([importItem]);
}