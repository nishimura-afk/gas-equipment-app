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

/**
 * ステータス集計シートを更新（完全版）
 */
function updateWebData() {
  Logger.log('ステータス集計を更新します...');
  
  try {
    const config = getConfig();
    const masterSheet = getSheet(config.SHEET_NAMES.MASTER_EQUIPMENT);
    const statusSheet = getSheet(config.SHEET_NAMES.STATUS_SUMMARY);
    const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
    
    // 各シートのデータを取得
    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    
    Logger.log('マスタ行数: ' + (masterData.length - 1) + '件');
    
    // 拠点マスタから拠点名マップを作成
    const locData = locSheet.getDataRange().getValues();
    const locMap = {};
    for (let i = 1; i < locData.length; i++) {
      if (locData[i][0]) {
        locMap[locData[i][0]] = locData[i][1];
      }
    }
    
    // 設備マスタの列インデックスマップを作成
    const masterColIdx = {};
    masterHeaders.forEach((header, idx) => {
      masterColIdx[header] = idx;
    });
    
    // ステータス集計用の新しいデータ配列
    const newData = [];
    const today = new Date();
    
    // 設備マスタの各行を処理
    for (let i = 1; i < masterData.length; i++) {
      const row = masterData[i];
      
      try {
        // 基本情報を取得
        const locCode = row[masterColIdx['拠点コード']];
        const eqId = row[masterColIdx['設備ID']];
        
        // 空行をスキップ
        if (!locCode || !eqId) {
          continue;
        }
        
        const locName = locMap[locCode] || locCode;
        const eqName = row[masterColIdx['設備名']] || '';
        const spec = row[masterColIdx['型式・仕様']] || '';
        const installDate = row[masterColIdx['設置日(前回実施)']];
        const partADate = row[masterColIdx['部品A交換日']];
        const partBDate = row[masterColIdx['部品B最終交換日']];
        const nextWorkMemo = row[masterColIdx['次回作業メモ']] || '';
        
        // メンテナンスサイクルを特定
        const cycle = findCycleByEquipmentId(eqId, eqName, config.MAINTENANCE_CYCLES);
        
        if (!cycle) {
          Logger.log(`警告: ${locCode} - ${eqId}(${eqName}) のサイクルが見つかりません`);
          continue;
        }
        
        // ステータスを計算
        const bodyStatus = calculateBodyStatus(installDate, cycle, today, config.STATUS);
        const partAStatus = calculatePartAStatus(partADate || installDate, cycle, today, config.STATUS);
        
        // 部品B: 本体更新かつ交換日が実際に記録されている場合のみ計算
        let partBStatus = config.STATUS.NORMAL;
        if (cycle.category === '本体更新' && partBDate && partBDate instanceof Date && !isNaN(partBDate.getTime())) {
          partBStatus = calculatePartBStatus(partBDate, cycle, today, config.STATUS);
        }
        
        // 次回予定日を計算
        const nextDate = calculateNextDate(installDate, partADate, cycle);
        
        // 部品B対象かどうか判定
        const hasPartB = (cycle.category === '本体更新' && partBDate) ? '対象' : '';
        
        // monthDiffAを計算（部品Aの経過月数）
        const monthDiffA = partADate 
          ? Math.floor(getYearsDiff(partADate, today) * 12) 
          : 0;
        
        // subsidyAlertの計算（必要に応じて）
        const subsidyAlert = '';
        
        // ステータス集計シートの列順に合わせて行データを作成
        // 列順: 拠点コード, 拠点名, 設備ID, 設備名, カテゴリ, 設置日, 部品Aステータス, 部品Bステータス, 
        //       本体ステータス, 部品B対象, monthDiffA, subsidyAlert, nextWorkMemo, spec, 次回予定日
        const newRow = [
          locCode,                    // 1. 拠点コード
          locName,                    // 2. 拠点名
          eqId,                       // 3. 設備ID
          eqName,                     // 4. 設備名
          cycle.category || '',       // 5. カテゴリ
          installDate,                // 6. 設置日(前回実施)
          partAStatus,                // 7. 部品Aステータス
          partBStatus,                // 8. 部品Bステータス
          bodyStatus,                 // 9. 本体ステータス
          hasPartB,                   // 10. 部品B対象
          monthDiffA,                 // 11. monthDiffA
          subsidyAlert,               // 12. subsidyAlert
          nextWorkMemo,               // 13. nextWorkMemo
          spec,                       // 14. spec
          nextDate                    // 15. 次回予定日
        ];
        
        newData.push(newRow);
        
      } catch (e) {
        Logger.log(`エラー: 行${i+1}の処理中: ${e.message}`);
      }
    }
    
    Logger.log('データ処理完了: ' + newData.length + '件');
    
    // ステータス集計シートをクリアして新しいデータを書き込み
    const lastRow = statusSheet.getLastRow();
    if (lastRow > 1) {
      statusSheet.getRange(2, 1, lastRow - 1, statusSheet.getLastColumn()).clearContent();
      Logger.log('既存データをクリア: ' + (lastRow - 1) + '行');
    }
    
    if (newData.length > 0) {
      statusSheet.getRange(2, 1, newData.length, 15).setValues(newData);
      Logger.log(`✅ ステータス集計シートに${newData.length}件のデータを書き込みました`);
    } else {
      Logger.log('⚠️ 書き込むデータがありません');
    }
    
    Logger.log('updateWebData完了');
    
  } catch (e) {
    Logger.log('❌ エラー: ' + e.message);
    Logger.log(e.stack);
    throw e;
  }
}

/**
 * 設備IDまたは設備名からメンテナンスサイクルを特定
 * 優先順位: suffix完全一致 > suffix部分一致 > searchKey
 */
function findCycleByEquipmentId(eqId, eqName, cycles) {
  let suffixMatch = null;
  let searchKeyMatch = null;
  
  for (const key in cycles) {
    const cycle = cycles[key];
    
    // 1. suffixで完全一致判定（最優先）
    if (cycle.suffix && eqId === cycle.suffix) {
      return cycle;
    }
    
    // 2. suffixで部分一致判定（2番目の優先度）
    if (cycle.suffix && eqId && eqId.includes(cycle.suffix)) {
      if (!suffixMatch) {
        suffixMatch = cycle;
      }
    }
    
    // 3. searchKeyで判定（最後の優先度）
    if (cycle.searchKey && eqName && eqName.includes(cycle.searchKey)) {
      if (!searchKeyMatch) {
        searchKeyMatch = cycle;
      }
    }
  }
  
  // suffixマッチを優先
  if (suffixMatch) {
    return suffixMatch;
  }
  
  // searchKeyマッチ
  if (searchKeyMatch) {
    return searchKeyMatch;
  }
  
  return null;
}

/**
 * 本体ステータスを計算
 */
function calculateBodyStatus(installDate, cycle, today, statusEnum) {
  if (!installDate || !(installDate instanceof Date) || isNaN(installDate.getTime())) {
    return statusEnum.NORMAL;
  }
  
  // 季節性のサイクルの場合
  if (cycle.seasonal && cycle.alertMonth && cycle.alertDay) {
    // 期限年を計算
    const deadlineYear = installDate.getFullYear() + cycle.years;
    const deadlineDate = new Date(deadlineYear, cycle.alertMonth - 1, cycle.alertDay);
    
    // 前年度のアラート日を計算（期限の1年前の指定月日）
    const alertYear = deadlineYear - 1;
    const alertDate = new Date(alertYear, cycle.alertMonth - 1, cycle.alertDay);
    
    // 期限超過判定
    if (today >= deadlineDate) {
      return statusEnum.PREPARE;
    }
    
    // アラート判定（前年度の指定日以降）
    if (today >= alertDate) {
      return statusEnum.NOTICE;
    }
    
    return statusEnum.NORMAL;
  }
  
  // 通常のサイクルの場合
  const yearsDiff = getYearsDiff(installDate, today);
  const alertTiming = cycle.alertTiming || { prepare: 0, notice: 0.5 };
  
  if (yearsDiff >= cycle.years + alertTiming.prepare) {
    return statusEnum.PREPARE;
  } else if (yearsDiff >= cycle.years - alertTiming.notice) {
    return statusEnum.NOTICE;
  }
  
  return statusEnum.NORMAL;
}

/**
 * 部品Aステータスを計算
 */
function calculatePartAStatus(partADate, cycle, today, statusEnum) {
  // 部品Aが関係ないサイクルの場合
  if (cycle.category !== '部材更新') {
    return statusEnum.NORMAL;
  }
  
  if (!partADate || !(partADate instanceof Date) || isNaN(partADate.getTime())) {
    return statusEnum.NORMAL;
  }
  
  const yearsDiff = getYearsDiff(partADate, today);
  
  // サイクルの年数を基準に判定
  if (yearsDiff >= cycle.years) {
    return statusEnum.NOTICE;
  }
  
  return statusEnum.NORMAL;
}

/**
 * 部品Bステータスを計算
 */
function calculatePartBStatus(partBDate, cycle, today, statusEnum) {
  // 部品Bが関係ないサイクルの場合
  if (cycle.category !== '本体更新') {
    return statusEnum.NORMAL;
  }
  
  // 部品B交換日が記録されていない場合は正常扱い（必須チェック）
  if (!partBDate || !(partBDate instanceof Date) || isNaN(partBDate.getTime())) {
    return statusEnum.NORMAL;
  }
  
  const yearsDiff = getYearsDiff(partBDate, today);
  const partBCycle = 5; // 部品Bの標準サイクル（年）
  
  if (yearsDiff >= partBCycle) {
    return statusEnum.NOTICE;
  }
  
  return statusEnum.NORMAL;
}

/**
 * 次回予定日を計算
 */
function calculateNextDate(installDate, partADate, cycle) {
  if (!installDate || !(installDate instanceof Date) || isNaN(installDate.getTime())) {
    return '';
  }
  
  const baseDate = partADate && partADate instanceof Date && !isNaN(partADate.getTime()) 
    ? partADate 
    : installDate;
  
  const nextDate = new Date(baseDate);
  nextDate.setFullYear(nextDate.getFullYear() + cycle.years);
  
  return Utilities.formatDate(nextDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}