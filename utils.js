/**
 * utils.js v1.0
 * 共通ユーティリティ関数
 * - PropertiesService経由の設定取得
 * - 標準レスポンス形式
 * - 列インデックス定数
 * - 共通ユーティリティ関数
 * - ログレベル管理
 */

// ========================================
// PropertiesService関連
// ========================================

/**
 * スクリプトプロパティから値を取得
 * @param {string} key - プロパティキー
 * @param {*} defaultValue - デフォルト値（省略可能）
 * @returns {string|*} プロパティ値またはデフォルト値
 * @throws {Error} 必須プロパティが未設定の場合
 */
function getScriptProperty(key, defaultValue = null) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value && defaultValue === null) {
    throw new Error(`必須プロパティ "${key}" が設定されていません`);
  }
  return value || defaultValue;
}

/**
 * スクリプトプロパティを設定（管理用）
 * @param {string} key - プロパティキー
 * @param {string} value - 設定する値
 */
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

// ========================================
// 標準レスポンス形式
// ========================================

/**
 * 標準APIレスポンスを生成
 * @param {boolean} success - 成功フラグ
 * @param {*} data - レスポンスデータ
 * @param {string|null} error - エラーメッセージ
 * @returns {Object} 標準レスポンスオブジェクト
 */
function createResponse(success, data, error) {
  return {
    success: success,
    data: data !== undefined ? data : null,
    error: error !== undefined ? error : null
  };
}

/**
 * 成功レスポンスを生成
 * @param {*} data - レスポンスデータ
 * @returns {Object} 成功レスポンス
 */
function successResponse(data) {
  return createResponse(true, data, null);
}

/**
 * エラーレスポンスを生成
 * @param {string} message - エラーメッセージ
 * @returns {Object} エラーレスポンス
 */
function errorResponse(message) {
  return createResponse(false, null, message);
}

// ========================================
// 列インデックス定数
// ========================================

/**
 * 案件管理シートの列インデックス（1始まり）
 */
const SCHEDULE_COLUMNS = {
  ID: 1,
  LOC_CODE: 2,
  EQ_ID: 3,
  WORK_TYPE: 4,
  DATE: 5,
  STATUS: 6,
  EVENT_ID: 7,
  NOTES: 8
};

/**
 * 設備マスタシートの列インデックス（0始まり、ヘッダー名でも取得可能）
 */
const MASTER_COLUMNS = {
  LOC_CODE: 0,
  LOC_NAME: 1,
  EQ_ID: 2,
  EQ_NAME: 3,
  SPEC: 4,
  INSTALL_DATE: 5,
  PART_A_DATE: 6,
  PART_B_DATE: 7,
  NEXT_WORK_MEMO: 9
};

// ========================================
// 共通ユーティリティ関数
// ========================================

/**
 * 拠点マップを構築
 * @returns {Object} 拠点コード -> 拠点名のマップ
 */
function buildLocationMap() {
  const config = getConfig();
  const locSheet = getSheet(config.SHEET_NAMES.MASTER_LOCATION);
  const locData = locSheet.getDataRange().getValues();
  const locMap = {};
  locData.slice(1).forEach(function(r) {
    if (r[0]) locMap[r[0]] = r[1];
  });
  return locMap;
}

/**
 * 本体更新案件がある拠点を取得
 * @param {Array} scheduleData - 案件管理シートのデータ
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} ガソリン/灯油の本体更新がある拠点のSetを含むオブジェクト
 */
function getBodyReplacementLocations(scheduleData, config) {
  const gasLocations = new Set();
  const keroseneLocations = new Set();

  scheduleData.slice(1).forEach(function(row) {
    if (row[5] !== config.PROJECT_STATUS.COMPLETED &&
        row[5] !== config.PROJECT_STATUS.CANCELLED) {
      const locCode = row[1];
      const eqId = row[2];

      if (eqId && (eqId.includes('PUMP-G-01') || eqId === 'REPLACE_GAS_PUMP')) {
        gasLocations.add(locCode);
      }
      if (eqId && (eqId.includes('PUMP-K-01') || eqId === 'REPLACE_KEROSENE_PUMP')) {
        keroseneLocations.add(locCode);
      }
    }
  });

  return { gasLocations: gasLocations, keroseneLocations: keroseneLocations };
}

// ========================================
// ログレベル管理
// ========================================

/**
 * ログレベル定義
 */
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

/**
 * 現在のログレベルを取得
 * @returns {number} ログレベル
 */
function getLogLevel() {
  const level = PropertiesService.getScriptProperties().getProperty('LOG_LEVEL') || 'INFO';
  return LOG_LEVELS[level] !== undefined ? LOG_LEVELS[level] : LOG_LEVELS.INFO;
}

/**
 * DEBUGログを出力
 * @param {string} message - ログメッセージ
 */
function logDebug(message) {
  if (getLogLevel() <= LOG_LEVELS.DEBUG) {
    Logger.log('[DEBUG] ' + message);
  }
}

/**
 * INFOログを出力
 * @param {string} message - ログメッセージ
 */
function logInfo(message) {
  if (getLogLevel() <= LOG_LEVELS.INFO) {
    Logger.log('[INFO] ' + message);
  }
}

/**
 * WARNログを出力
 * @param {string} message - ログメッセージ
 */
function logWarn(message) {
  if (getLogLevel() <= LOG_LEVELS.WARN) {
    Logger.log('[WARN] ' + message);
  }
}

/**
 * ERRORログを出力（常に出力）
 * @param {string} message - ログメッセージ
 */
function logError(message) {
  Logger.log('[ERROR] ' + message);
}

// ========================================
// メール関連のヘルパー
// ========================================

/**
 * 管理者メールアドレスを取得
 * @returns {string} 管理者メールアドレス
 */
function getAdminMail() {
  return getConfig().ADMIN_MAIL;
}

/**
 * Gmail下書きを作成（共通処理）
 * @param {string} to - 宛先（空文字列可）
 * @param {string} subject - 件名
 * @param {string} body - 本文
 * @returns {Object} 作成結果
 */
function createGmailDraftWithFrom(to, subject, body) {
  GmailApp.createDraft(to, subject, body, {
    from: getConfig().ADMIN_MAIL
  });
  return successResponse({ message: 'Gmail下書きを作成しました' });
}
