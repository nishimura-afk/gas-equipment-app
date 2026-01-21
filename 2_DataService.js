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
  const config = getConfig();
  const seasonalCycles = Object.values(config.MAINTENANCE_CYCLES)
    .filter(c => c.seasonal && c.alertMonth && c.alertDay);
  const installDateIndex = headers.indexOf('設置日(前回実施)');
  const equipmentIdIndex = headers.indexOf('設備ID');
  const equipmentNameIndex = headers.indexOf('設備名');
  const today = new Date();
  
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = (row[i] instanceof Date) ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[i];
    });
    
    applySeasonalAlertOverride(obj, row, {
      seasonalCycles,
      installDateIndex,
      equipmentIdIndex,
      equipmentNameIndex,
      today,
      status: config.STATUS
    });

    return obj;
  });
}

function getYearsDiff(fromDate, toDate = new Date()) {
  if (!fromDate || !(fromDate instanceof Date) || isNaN(fromDate.getTime())) return 0;
  return (toDate.getFullYear() - fromDate.getFullYear()) + ((toDate.getMonth() - fromDate.getMonth()) / 12);
}

function applySeasonalAlertOverride(obj, row, context) {
  if (!context.seasonalCycles || context.seasonalCycles.length === 0) return;
  if (!obj || !obj['本体ステータス'] || obj['本体ステータス'] === context.status.NORMAL) return;
  if (context.equipmentIdIndex === -1 && context.equipmentNameIndex === -1) return;
  if (context.installDateIndex === -1) return;
  
  const equipmentId = String(row[context.equipmentIdIndex] || '');
  const equipmentName = String(row[context.equipmentNameIndex] || '');
  const installDate = parseDateValue(row[context.installDateIndex]);
  const cycle = findSeasonalCycle(context.seasonalCycles, equipmentId, equipmentName);
  const alertDate = getSeasonalAlertDate(installDate, cycle);
  
  if (!alertDate) return;
  if (context.today < alertDate) obj['本体ステータス'] = context.status.NORMAL;
}

function findSeasonalCycle(cycles, equipmentId, equipmentName) {
  if (!cycles || cycles.length === 0) return null;
  return cycles.find(cycle => {
    if (cycle.suffix && equipmentId && equipmentId.includes(cycle.suffix)) return true;
    if (cycle.searchKey && equipmentName && equipmentName.includes(cycle.searchKey)) return true;
    return false;
  }) || null;
}

function getSeasonalAlertDate(installDate, cycle) {
  if (!cycle || !installDate) return null;
  if (!cycle.alertMonth || !cycle.alertDay || !cycle.years) return null;
  if (!(installDate instanceof Date) || isNaN(installDate.getTime())) return null;
  const targetYear = installDate.getFullYear() + cycle.years;
  return new Date(targetYear, cycle.alertMonth - 1, cycle.alertDay);
}

function parseDateValue(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'string') {
    const parts = value.split('/');
    if (parts.length === 3) {
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      const day = parseInt(parts[2], 10);
      if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
        return new Date(year, month - 1, day);
      }
    }
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}