/**
 * 4_CalendarService.gs
 * カレンダー連携
 */
function createMaintenanceEvent(locationCode, equipmentId, workType, dateString, notes) {
  try {
    const config = getConfig();
    const list = getEquipmentListCached();
    const target = list.find(d => d['拠点コード'] == locationCode && d['設備ID'] == equipmentId);
    const locName = target ? target['拠点名'] : '拠点不明';
    const eqName = target ? target['設備名'] : equipmentId;
   
    const event = (CalendarApp.getCalendarById(config.CALENDAR_ID) || CalendarApp.getDefaultCalendar())
      .createAllDayEvent(`【設備メンテ】${locName} ${eqName} - ${workType}`, new Date(dateString), { description: `備考:${notes}` });
    return { message: `登録完了`, eventId: event.getId() };
  } catch (e) {
    throw new Error('カレンダー登録失敗: ' + e.message);
  }
}

function markEventAsCompleted(eventId, completionDate) {
  try {
    const event = CalendarApp.getEventById(eventId);
    if(event) event.setTitle(`【完了】${event.getTitle().replace('【設備メンテ】', '')}`);
    return { message: '完了マーク済み' };
  } catch (e) {
    return { message: 'スキップ' };
  }
}