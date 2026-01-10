/**
 * 5_DriveService.gs v5.7
 * ID生成にUUIDを採用
 */
const USER_DRIVE_ID = '1wWBNUfHcoK9AIffNLOzc5Xbl_CBYtWH-';
const INBOX_FOLDER_NAME = 'SS見積_受信BOX'; 
const ARCHIVE_FOLDER_NAME = 'SS見積_処理済';

function ensureInboxFolder() {
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(USER_DRIVE_ID); } catch (e) { rootFolder = DriveApp.getRootFolder(); }
  const inboxFolders = rootFolder.getFoldersByName(INBOX_FOLDER_NAME);
  const inboxFolder = inboxFolders.hasNext() ? inboxFolders.next() : rootFolder.createFolder(INBOX_FOLDER_NAME);
  const archiveFolders = rootFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
  if (!archiveFolders.hasNext()) rootFolder.createFolder(ARCHIVE_FOLDER_NAME);
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
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(USER_DRIVE_ID); } catch(e) { rootFolder = DriveApp.getRootFolder(); }
  const archiveFolders = rootFolder.getFoldersByName(ARCHIVE_FOLDER_NAME);
  const archiveFolder = archiveFolders.hasNext() ? archiveFolders.next() : rootFolder.createFolder(ARCHIVE_FOLDER_NAME);
  const locList = getStoreList();
  const locMap = {};
  locList.forEach(l => locMap[l.code] = l.name);
  let successCount = 0;

  filesToImport.forEach(item => {
    try {
      const file = DriveApp.getFileById(item.fileId);
      if (item.projectType === 'NEW') {
        const uniqueId = Utilities.getUuid();
        scheduleSheet.appendRow([uniqueId, item.locCode, item.eqId, '見積受領(インポート)', '', config.PROJECT_STATUS.ESTIMATE_RCV, '', '']);
      }
      getSheet(config.SHEET_NAMES.HISTORY).appendRow([item.locCode, item.eqId, '見積書登録', new Date(), `File: ${file.getName()}\nUrl: ${file.getUrl()}`, '']);
      
      const newName = `[${item.locCode}]${item.eqId}_見積_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}.pdf`;
      file.setName(newName);
      const shopName = locMap[item.locCode] || item.locCode;
      const subFolders = archiveFolder.getFoldersByName(shopName);
      const targetFolder = subFolders.hasNext() ? subFolders.next() : archiveFolder.createFolder(shopName);
      file.moveTo(targetFolder);
      successCount++;
    } catch (e) { console.error('Import Error', e); }
  });
  return { success: true, message: `${successCount}件処理完了`, folderUrl: archiveFolder.getUrl() };
}

function uploadAndImport(data, fileName, mimeType, projectInfo) {
  const folderInfo = ensureInboxFolder();
  const folder = DriveApp.getFolderById(folderInfo.id);
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
  const file = folder.createFile(blob);
  const importItem = { fileId: file.getId(), projectType: projectInfo.type, projectId: projectInfo.id, locCode: projectInfo.locCode, eqId: projectInfo.eqId };
  return executeImport([importItem]);
}