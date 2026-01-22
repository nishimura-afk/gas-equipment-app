/**
 * 0_Config.gs v6.1
 * 設定一元管理版
 * - ベンダー情報（メールアドレス含む）を集約
 * - アラート閾値を定義
 * - 釣銭機カバーをseasonal: falseに修正
 */
function getConfig() {
  const SPREADSHEET_ID = '1FKMS0xNHEcftYmZ2u1Q2VrAtPKnKm5E5dG1tnlR6fo';
  
  // システム管理者（エラー通知やCC用）
  const ADMIN_MAIL = 'nishimura@selfix.jp'; 
 
const SHEET_NAMES = {
  MASTER_EQUIPMENT: '設備マスタ',
  MASTER_LOCATION: '拠点マスタ',
  SCHEDULE: '案件管理',
  HISTORY: '履歴',
  STATUS_SUMMARY: 'ステータス集計',
  SYS_LOG: 'システムログ',
  CONFIG_MASTER: '設定マスタ',
  ESTIMATE_HEADER: '見積比較',      // ★追加
  ESTIMATE_DETAIL: '見積明細'       // ★追加
};

  // アラートを出す時期（年数基準に対する前倒し期間）
  // 設備ごとに個別設定したため、この全体設定は非推奨
  // const ALERT_THRESHOLDS = {
  //   BODY_PREPARE: 0,    // 期限到達
  //   BODY_NOTICE:  1.0,  // 1年前から注意
  //   LEGAL_PREPARE: 0,
  //   LEGAL_NOTICE: 0.2,  // 約2.4ヶ月前
  //   PARTS_PREPARE: 0,
  //   PARTS_NOTICE: 0.2,
  //   SEASONAL_NOTICE: 0.4 // 季節性は少し早め
  // };

  // ベンダー定義（メールアドレス・担当領域）
  const VENDORS = {
    'TATSUNO': { 
      name: '株式会社タツノ', 
      email: 'tatsuno_sample@example.com',
      keywords: ['検定', '計量機', 'エコステージ', 'ノズル', 'ホース', '油種', 'PUMP', 'ECO'] 
    },
    'ASAHI': { 
      name: '朝日エティック株式会社', 
      email: 'asahi_sample@example.com', 
      keywords: ['塗装', '照明', '投光器', 'サイン', '看板', 'キャノピー', 'PAINT', 'LED', 'FLOOD', 'SIGN', 'WASH-S'] 
    },
    'SHARP': { 
      name: 'シャープ', 
      email: 'sharp_sample@example.com', 
      keywords: ['POS', 'カメラ', '釣銭機', 'シール', 'カバー', 'タッチパネル', 'CAM'] 
    },
    'ALTIA': { 
      name: '株式会社アルティア', 
      email: 'altia_sample@example.com', 
      keywords: ['コンプレッサー', 'エアタワー', 'COMP', 'AIR'] 
    },
    'DAIICHI': { 
      name: '第一工業株式会社', 
      email: 'daiichi_sample@example.com', 
      keywords: ['漏洩', '漏えい', '液面', '通気', 'TANK', 'LVL', 'VENT'] 
    },
    'OTHERS': { 
      name: 'その他', 
      email: '', 
      keywords: [] 
    }
  };

  // メンテナンスサイクル定義
  let MAINTENANCE_CYCLES = {
    // ========================================
    // 法定検査（4.8ヶ月前アラート）
    // ========================================
    'INSPECTION_TANK': { 
      category: '法定検査', 
      years: 3, 
      label: 'タンク漏えい検査', 
      searchKey: 'タンク', 
      suffix: 'TANK-01',
      alertTiming: {
        prepare: 0,
        notice: 0.4    // 4.8ヶ月前
      }
    },
    
    'INSPECTION_KEROSENE': { 
      category: '法定検査', 
      years: 7, 
      label: '灯油計量機検定', 
      searchKey: '灯油検定', 
      suffix: 'PUMP-K-CHK',
      alertTiming: {
        prepare: 0,
        notice: 0.4    // 4.8ヶ月前
      }
    },

    // ========================================
    // 本体更新（前年度1月アラート）
    // ========================================
    'REPLACE_GAS_PUMP': { 
      category: '本体更新', 
      years: 7, 
      label: 'ガソリン計量機更新', 
      searchKey: 'ガソリン', 
      suffix: 'PUMP-G-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'REPLACE_KEROSENE_PUMP': { 
      category: '本体更新', 
      years: 14, 
      label: '灯油計量機更新', 
      searchKey: '灯油計量', 
      suffix: 'PUMP-K-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'REPLACE_ECOSTAGE': { 
      category: '本体更新', 
      years: 9, 
      label: 'エコステージL100R更新', 
      searchKey: 'エコ', 
      suffix: 'ECO-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'REPLACE_POS': { 
      category: '本体更新', 
      years: 10, 
      label: 'POS更新', 
      searchKey: 'POS', 
      suffix: 'POS-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'PAINTING': { 
      category: '美観', 
      years: 7, 
      label: '塗装', 
      searchKey: '塗装', 
      suffix: 'PAINT-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },

    // ========================================
    // 本体更新（6ヶ月前アラート）
    // ========================================
    'REPLACE_COMPRESSOR': { 
      category: '本体更新', 
      years: 12, 
      label: 'コンプレッサー更新', 
      searchKey: 'コンプレッサー', 
      suffix: 'COMP-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5    // 6ヶ月前
      }
    },
    
    'REPLACE_BREAKER': { 
      category: '本体更新', 
      years: 10, 
      label: '電子ブレーカー更新', 
      searchKey: 'ブレーカー', 
      suffix: 'BRK-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'REPLACE_WELL_PUMP': { 
      category: '本体更新', 
      years: 10, 
      label: '井戸ポンプ更新', 
      searchKey: '井戸', 
      suffix: 'WELL-P-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'REPLACE_WATER_PUMP': { 
      category: '本体更新', 
      years: 15, 
      label: '送水ポンプ更新', 
      searchKey: '送水', 
      suffix: 'WTR-P-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'REPLACE_VENT_PIPE': { 
      category: '本体更新', 
      years: 20, 
      label: '通気管設置経過', 
      searchKey: '通気', 
      suffix: 'VENT-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'REFORM_TOILET': { 
      category: '本体更新', 
      years: 15, 
      label: 'トイレリフォーム', 
      searchKey: 'リフォーム', 
      suffix: 'TOILET-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },

    // ========================================
    // 本体更新（季節性・3月1日アラート）
    // ========================================
    'REPLACE_AIRCON': { 
      category: '本体更新', 
      years: 6, 
      label: 'エアコン更新', 
      searchKey: 'エアコン', 
      suffix: 'AC-01', 
      seasonal: true, 
      alertMonth: 3, 
      alertDay: 1
    },
    
    'REPLACE_FIRE_EXTINGUISHER': { 
      category: '本体更新', 
      years: 10, 
      label: '消火器更新', 
      searchKey: '消火器', 
      suffix: 'FIRE-EXT-01',
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },

    // ========================================
    // 不定期（アラートなし）
    // ========================================
    'REPLACE_CAMERA': { 
      category: '不定期', 
      years: 99, 
      label: '監視カメラ更新', 
      searchKey: 'カメラ', 
      suffix: 'CAM-01'
    },
    
    'REPLACE_AIR_TOWER': { 
      category: '不定期', 
      years: 99, 
      label: 'エアタワー更新', 
      searchKey: 'エアタワー', 
      suffix: 'AIR-01'
    },
    
    'REPLACE_LEVEL_GAUGE': { 
      category: '不定期', 
      years: 99, 
      label: '液面計更新', 
      searchKey: '液面', 
      suffix: 'LVL-01'
    },

    // ========================================
    // 美観（6ヶ月前アラート）
    // ========================================
    'LED_CANOPY': { 
      category: '美観', 
      years: 10, 
      label: 'キャノピーLED更新', 
      searchKey: 'LED', 
      suffix: 'LED-C-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'FLOODLIGHT': { 
      category: '美観', 
      years: 10, 
      label: '投光器更新', 
      searchKey: '投光器', 
      suffix: 'FLOOD-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'FIELD_LIGHT': { 
      category: '美観', 
      years: 20, 
      label: 'フィールド照明灯', 
      searchKey: 'フィールド', 
      suffix: 'FIELD-L-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'SIGN_POLE_FACE': { 
      category: '美観', 
      years: 10, 
      label: 'サインポール面板更新', 
      searchKey: '面板', 
      suffix: 'SIGN-P-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'SIGN_POLE_LED': { 
      category: '美観', 
      years: 10, 
      label: 'サインポール蛍光灯/LED更新', 
      searchKey: '蛍光灯', 
      suffix: 'SIGN-L-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'SIGN_PRICE_LED': { 
      category: '美観', 
      years: 8, 
      label: 'サインポール価格LED更新', 
      searchKey: '価格LED', 
      suffix: 'SIGN-PR-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'SIGN_CANOPY': { 
      category: '美観', 
      years: 10, 
      label: 'キャノピーサイン更新', 
      searchKey: 'キャノピーサイン', 
      suffix: 'SIGN-CANOPY',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },
    
    'SIGN_CARWASH': { 
      category: '美観', 
      years: 10, 
      label: '洗車機価格看板更新', 
      searchKey: '洗車機価格', 
      suffix: 'WASH-S-01',
      alertTiming: {
        prepare: 0,
        notice: 0.5
      }
    },

    // ========================================
    // 部材更新・メンテ
    // ========================================
    'PARTS_PUMP_1Y': { 
      category: '部材更新', 
      years: 1, 
      label: '計量機消耗品(毎年)', 
      searchKey: 'ノズルカバー', 
      suffix: 'PARTS-PUMP-1Y', 
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'PARTS_CHANGE_3Y': { 
      category: '部材更新', 
      years: 3, 
      label: '釣銭機シール貼替', 
      searchKey: 'シール', 
      suffix: 'PARTS-SEAL-3Y', 
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'PARTS_PUMP_4Y': { 
      category: '部材更新', 
      years: 4, 
      label: 'ガソリン計量機部品(4年)', 
      searchKey: 'ガソリン計量機部品', 
      suffix: 'PARTS-PUMP-4Y', 
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'MAINT_WELL_5Y': { 
      category: '本体更新', 
      years: 5, 
      label: '井戸ポンプメンテ', 
      searchKey: '井戸メンテ', 
      suffix: 'MAINT-WELL-5Y',
      alertTiming: {
        prepare: 0,
        notice: 0.2    // 2.4ヶ月前
      }
    },
    
    'PARTS_CHANGE_6Y': { 
      category: '部材更新', 
      years: 6, 
      label: '釣銭機カバー/パネル', 
      searchKey: '釣銭機', 
      suffix: 'CHG-01', 
      seasonal: true,
      alertMonth: 1,
      alertDay: 1
    },
    
    'PARTS_KEROSENE_7Y': { 
      category: '本体更新',    // カテゴリ変更
      years: 7, 
      label: '灯油パネル更新', 
      searchKey: '灯油パネル', 
      suffix: 'PARTS-K-PANEL-7Y', 
      seasonal: false,
      alertTiming: {
        prepare: 0,
        notice: 0.4    // 灯油検定と同じ4.8ヶ月前
      }
    }
  };

  const STATUS = { NORMAL: '正常', NOTICE: '実施時期', PREPARE: '期限超過', DONE: '実施済' };
  const PROJECT_STATUS = { ESTIMATE_REQ: '見積依頼中', ESTIMATE_RCV: '見積受領', ORDERED: '発注済', SCHEDULED: '日程確定', COMPLETED: '完了', CANCELLED: '取り消し' };
  const CALENDAR_ID = 'primary';

  return { SPREADSHEET_ID, SHEET_NAMES, MAINTENANCE_CYCLES, STATUS, PROJECT_STATUS, CALENDAR_ID, ADMIN_MAIL, VENDORS };
}

function getSheet(sheetName) {
  const config = getConfig();
  let ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート "${sheetName}" が見つかりません。`);
  return sheet;
}