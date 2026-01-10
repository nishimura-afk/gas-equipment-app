/**
 * 2_DataService.gs v4.7
 * データの取得に特化
 */
function getEquipmentMasterData() {
  return getEquipmentListCached();
}

function getEquipmentListCached() {
  const sheet = getSheet(getConfig().SHEET_NAMES.STATUS_SUMMARY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = (row[i] instanceof Date) ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[i];
    });
    return obj;
  });
}

function getYearsDiff(fromDate, toDate = new Date()) {
  if (!fromDate || !(fromDate instanceof Date) || isNaN(fromDate.getTime())) return 0;
  return (toDate.getFullYear() - fromDate.getFullYear()) + ((toDate.getMonth() - fromDate.getMonth()) / 12);
}