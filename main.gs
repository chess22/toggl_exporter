/*
  Toggl time entries export to GoogleCalendar
  original author: Masato Kawaguchi
  modified by: chess
  Released under the BSD-3-Clause license
  version: 1.3.00
  https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE
*/

/**
 * TogglとGoogleカレンダーを同期するGoogle Apps Script
 * 
 * 機能:
 * 1. Togglからタイムエントリを取得
 * 2. Googleカレンダーにイベントを追加・更新・削除
 * 3. 重複するイベントの作成を防止
 * 4. エラーハンドリングと通知
 * 5. カスタムメニューによる手動操作の提供
 */

const CONFIG = {
  CACHE_KEY: 'toggl_exporter:lastmodify_datetime', // キャッシュキー
  TIME_OFFSET: 9 * 60 * 60, // JST (秒)
  TOGGL_API_HOSTNAME: 'https://api.track.toggl.com',
  GOOGLE_CALENDAR_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_CALENDAR_ID'), // GoogleカレンダーID
  NOTIFICATION_EMAIL: PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL'), // エラー通知メールアドレス
  TOGGL_BASIC_AUTH: PropertiesService.getScriptProperties().getProperty('TOGGL_BASIC_AUTH'), // Toggl API認証情報（Base64エンコード済み）
  DEBUG_MODE: false, // デバッグモード（true: ログ出力, false: ログ出力なし）
  MAX_CACHE_AGE: 12 * 60 * 60, // 12時間（秒単位）
  RETRY_COUNT: 5, // 再試行回数
  RETRY_DELAY: 2000, // 再試行間隔（ミリ秒）
};

/**
 * ログレベルの定義
 */
const LOG_LEVELS = {
  DEBUG: 1,
  INFO: 2,
  ERROR: 3,
};

const CURRENT_LOG_LEVEL = CONFIG.DEBUG_MODE ? LOG_LEVELS.DEBUG : LOG_LEVELS.INFO;

/**
 * ログ出力用関数
 * @param {number} level - ログレベル
 * @param {string} message - ログメッセージ
 */
function log(level, message) {
  if (level >= CURRENT_LOG_LEVEL) {
    switch(level) {
      case LOG_LEVELS.DEBUG:
        Logger.log(`DEBUG: ${message}`);
        break;
      case LOG_LEVELS.INFO:
        Logger.log(`INFO: ${message}`);
        break;
      case LOG_LEVELS.ERROR:
        Logger.log(`ERROR: ${message}`);
        break;
    }
  }
}

/**
 * エラー発生時に通知メールを送信する関数
 * @param {Error} e - 発生したエラー
 * @param {string} [record_id] - 関連するレコードID（オプション）
 */
function notifyError(e, record_id = null) {
  const email = CONFIG.NOTIFICATION_EMAIL;
  const subject = 'Google Apps Script エラー通知';
  let body = `
以下のエラーが発生しました:

エラーメッセージ:
${e.toString()}

スタックトレース:
${e.stack}
  `;
  
  if (record_id) {
    body += `\n関連するレコードID: ${record_id}`;
  }
  
  if (email) {
    MailApp.sendEmail(email, subject, body);
    log(LOG_LEVELS.INFO, `エラー通知メールを送信しました: ${email}`);
  } else {
    log(LOG_LEVELS.ERROR, "通知先メールアドレスが設定されていません。");
  }
}

/**
 * ロックを取得する関数
 * @returns {Lock} 取得したロックオブジェクト
 */
function getLock() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 最大30秒待機
    log(LOG_LEVELS.INFO, "Lock acquired");
    return lock;
  } catch (e) {
    log(LOG_LEVELS.ERROR, "Could not acquire lock: " + e);
    throw e;
  }
}

/**
 * 指定した関数を再試行するユーティリティ関数
 * @param {Function} fn - 実行する関数
 * @param {number} retries - 最大再試行回数
 * @param {number} delay - 再試行間の待機時間（ミリ秒）
 * @returns {*} 関数の戻り値
 */
function retry(fn, retries = CONFIG.RETRY_COUNT, delay = CONFIG.RETRY_DELAY) {
  for (let i = 0; i < retries; i++) {
    try {
      return fn();
    } catch (e) {
      if (i < retries - 1) {
        log(LOG_LEVELS.DEBUG, `リトライ中 (${i + 1}/${retries}) - エラー: ${e.message}`);
        Utilities.sleep(delay);
      } else {
        throw e;
      }
    }
  }
}

/**
 * テストメールを送信する関数
 */
function sendTestEmail() {
  const email = CONFIG.NOTIFICATION_EMAIL;
  const subject = 'Google Apps Script テストメール';
  const body = `
これはGoogle Apps Scriptからのテストメールです。

エラー通知機能が正常に動作していることを確認してください。
  `;
  
  if (email) {
    MailApp.sendEmail(email, subject, body);
    log(LOG_LEVELS.INFO, `テストメールを送信しました: ${email}`);
  } else {
    log(LOG_LEVELS.ERROR, "通知先メールアドレスが設定されていません。");
  }
}

/**
 * キャッシュをクリアする関数
 */
function clearScriptCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
  log(LOG_LEVELS.INFO, "Cache cleared");
  SpreadsheetApp.getUi().alert("キャッシュをクリアしました。");
}

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('キャッシュをクリア', 'clearScriptCache')
    .addItem('テストメールを送信', 'sendTestEmail')
    .addItem('重複イベントを削除', 'removeDuplicateEvents')
    .addItem('重複イベント作成テスト', 'testCreateDuplicateEvents')
    .addItem('重複イベント削除テスト', 'testRemoveDuplicateEvents')
    .addItem('重複イベント統合テスト', 'testDuplicateEventsWorkflow')
    .addToUi();
}

/**
 * 指定された日数の開始時刻（00:00:00）を返す関数
 * @param {number} days - 現在からの日数（正の数で未来、負の数で過去）
 * @returns {string} ISO8601形式の日時文字列（UTC）
 */
function beginningOfDay(days) {
  var now = new Date();
  now.setUTCHours(0, 0, 0, 0);
  now.setUTCDate(now.getUTCDate() + days);
  
  return Utilities.formatDate(now, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/**
 * キャッシュから最後の更新日時を取得する関数
 * @returns {number} UNIXタイムスタンプ（秒）または-1（キャッシュが存在しない/エラー）
 */
function getLastModifyDatetime() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get(CONFIG.CACHE_KEY);
  
  if (!cachedData) {
    log(LOG_LEVELS.DEBUG, "No cached data found");
    return -1;
  }
  
  log(LOG_LEVELS.DEBUG, "Cached data: " + cachedData);
  
  try {
    var parsedData = JSON.parse(cachedData);
    if (parsedData && typeof parsedData.timestamp === 'number') {
      var cacheAge = (Date.now() / 1000) - parsedData.timestamp;
      if (cacheAge > CONFIG.MAX_CACHE_AGE) {
        log(LOG_LEVELS.DEBUG, "Cache expired");
        cache.remove(CONFIG.CACHE_KEY);
        return -1;
      }
      log(LOG_LEVELS.DEBUG, "Parsed cached data: " + parsedData.timestamp);
      return parsedData.timestamp;
    } else {
      log(LOG_LEVELS.DEBUG, "Invalid cached data format");
      cache.remove(CONFIG.CACHE_KEY);
      return -1;
    }
  } catch (e) {
    log(LOG_LEVELS.ERROR, "Error parsing cached data: " + e.message);
    cache.remove(CONFIG.CACHE_KEY);
    return -1;
  }
}

/**
 * キャッシュに最後の更新日時を保存する関数
 * @param {number} unix_timestamp - UNIXタイムスタンプ（秒）
 */
function putLastModifyDatetime(unix_timestamp) {
  var cache = CacheService.getScriptCache();
  var data = JSON.stringify({ timestamp: unix_timestamp });
  cache.put(CONFIG.CACHE_KEY, data, CONFIG.MAX_CACHE_AGE); // 12時間の有効期限
  log(LOG_LEVELS.DEBUG, "Cache updated with timestamp: " + data);
}

/**
 * Toggl APIからタイムエントリを取得する関数
 * @param {number} start_timestamp - 開始時刻のUNIXタイムスタンプ（秒）
 * @param {number} days - 取得する日数
 * @returns {Array|null} タイムエントリの配列またはnull（エラー時）
 */
function getTimeEntries(start_timestamp, days) {
  return retry(() => {
    var startOfDayStr = beginningOfDay(0); // 現在の日の開始時刻
    var endOfDayStr = beginningOfDay(days); // 取得終了日
    
    var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries?start_date=${encodeURIComponent(startOfDayStr)}&end_date=${encodeURIComponent(endOfDayStr)}`;
    
    log(LOG_LEVELS.DEBUG, `Fetching time entries from: ${uri}`);
    
    var response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    log(LOG_LEVELS.DEBUG, `API Response Code: ${responseCode}`);
    log(LOG_LEVELS.DEBUG, `API Response: ${responseText}`);
    
    if (responseCode !== 200) {
      log(LOG_LEVELS.ERROR, `API Error: ${responseText}`);
      throw new Error(`Toggl API returned status code ${responseCode}`);
    }
    
    var parsed = JSON.parse(responseText);
    if (Array.isArray(parsed)) {
      return parsed;
    } else {
      log(LOG_LEVELS.ERROR, "API returned non-array response");
      return null;
    }
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Toggl APIからプロジェクトデータを取得する関数
 * @param {string} workspace_id - ワークスペースID
 * @param {string} project_id - プロジェクトID
 * @returns {Object} プロジェクトデータまたは空オブジェクト
 */
function getProjectData(workspace_id, project_id) {
  if (!workspace_id || !project_id) return {};
  
  return retry(() => {
    var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/workspaces/${workspace_id}/projects/${project_id}`;
    
    var response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    log(LOG_LEVELS.DEBUG, `Project API Response Code: ${responseCode}`);
    log(LOG_LEVELS.DEBUG, `Project API Response: ${responseText}`);
    
    if (responseCode !== 200) {
      log(LOG_LEVELS.ERROR, `Project API Error: ${responseText}`);
      return {};
    }
    
    return JSON.parse(responseText);
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Googleカレンダーから関連するすべてのイベントを取得し、マッピングする関数
 * @returns {Object} record_idをキーとしたイベントのマッピング
 */
function getAllRelevantEvents() {
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  var now = new Date();
  var pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate()); // 過去3ヶ月間
  
  var events = calendar.getEvents(pastDate, now);
  
  var eventMap = {};
  
  events.forEach(function(event) {
    var title = event.getTitle();
    var match = title.match(/ID:(\d+)$/);
    if (match && match[1]) {
      var record_id = match[1];
      // 最新のイベントのみを保持
      if (!eventMap[record_id] || new Date(event.getLastUpdated()) > new Date(eventMap[record_id].getLastUpdated())) {
        eventMap[record_id] = event;
      }
    }
  });
  
  return eventMap;
}

/**
 * Googleカレンダーにイベントを記録する関数
 * @param {string} title - イベントのタイトル（IDを含む）
 * @param {string} started_at - イベント開始時刻のISO8601文字列
 * @param {string} ended_at - イベント終了時刻のISO8601文字列
 */
function recordActivityLog(title, started_at, ended_at) {
  log(LOG_LEVELS.DEBUG, `recordActivityLog called with title: "${title}", started_at: "${started_at}", ended_at: "${ended_at}"`);
  
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  
  // カレンダーが取得できているか確認
  if (!calendar) {
    log(LOG_LEVELS.ERROR, `Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
  }
  
  // 日時文字列をDateオブジェクトに変換
  var startDate = new Date(started_at);
  var endDate = new Date(ended_at);
  
  // 日時の有効性をチェック
  if (isNaN(startDate.getTime())) {
    log(LOG_LEVELS.ERROR, `Invalid start date: "${started_at}"`);
    throw new RangeError(`Invalid start date: "${started_at}"`);
  }
  
  if (isNaN(endDate.getTime())) {
    log(LOG_LEVELS.ERROR, `Invalid end date: "${ended_at}"`);
    throw new RangeError(`Invalid end date: "${ended_at}"`);
  }
  
  // タイムゾーンを設定（必要に応じて）
  // calendar.setTimeZone('Asia/Tokyo'); // 必要に応じてコメントアウト
  
  try {
    // イベントを作成
    calendar.createEvent(title, startDate, endDate);
    log(LOG_LEVELS.INFO, `Created event: "${title}" from ${started_at} to ${ended_at}`);
  } catch (createError) {
    log(LOG_LEVELS.ERROR, `Error creating event: ${createError}`);
    throw createError;
  }
}

/**
 * Googleカレンダーに既にイベントが存在するかを確認し、必要に応じて更新する関数
 * @param {string} record_id - タイムエントリのID
 * @param {Object} newRecord - 更新後のタイムエントリデータ
 * @returns {boolean} イベントが存在し、更新された場合はtrue、存在しない場合はfalse
 */
function eventExistsAndUpdate(record_id, newRecord) {
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  
  // タイムエントリのIDを含むイベントを検索
  var searchQuery = `ID:${record_id}`;
  var now = new Date();
  var pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate()); // 過去3ヶ月間を検索対象
  
  var events = calendar.getEvents(pastDate, now, { search: searchQuery });
  
  log(LOG_LEVELS.DEBUG, `Searching for events with ID:${record_id}. Found ${events.length} events.`);
  
  if (events.length > 0) {
    // 正確に一致するイベントを見つける
    var matchingEvent = null;
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var titleMatch = event.getTitle().match(/ID:(\d+)$/);
      if (titleMatch && titleMatch[1] === String(record_id)) { // 型を統一
        matchingEvent = event;
        break;
      }
    }
    
    if (matchingEvent) {
      log(LOG_LEVELS.DEBUG, `Matching event found: "${matchingEvent.getTitle()}"`);
      var eventTitle = matchingEvent.getTitle();
      var newStart = newRecord.start;
      var newEnd = newRecord.stop;
      
      // タイムエントリの名称を基に新しいタイトルを生成
      var project_data = getProjectData(newRecord.wid, newRecord.pid);
      var project_name = project_data.name || '';
      var updatedTitle = [(newRecord.description || '名称なし'), project_name].filter(Boolean).join(" : ") + ` ID:${record_id}`;
      
      // イベントタイトルと時間の比較
      var eventTitleNeedsUpdate = matchingEvent.getTitle() !== updatedTitle;
      var eventTimeNeedsUpdate = (matchingEvent.getStartTime().toISOString() !== newStart) || (matchingEvent.getEndTime().toISOString() !== newEnd);
      
      log(LOG_LEVELS.DEBUG, `EventTitleNeedsUpdate: ${eventTitleNeedsUpdate}, EventTimeNeedsUpdate: ${eventTimeNeedsUpdate}`);
      
      if (eventTitleNeedsUpdate || eventTimeNeedsUpdate) {
        try {
          retry(() => {
            matchingEvent.setTitle(updatedTitle);
            matchingEvent.setTime(new Date(newStart), new Date(newEnd));
          }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY); // 再試行ロジック適用
          log(LOG_LEVELS.INFO, `Updated event for ID:${record_id}`);
        } catch (e) {
          log(LOG_LEVELS.ERROR, `Error updating event for ID: ${record_id} - ${e}`);
          notifyError(e, record_id);
          return false;
        }
      } else {
        log(LOG_LEVELS.DEBUG, `No update needed for event ID:${record_id}`);
      }
      
      return true; // イベントは存在し、必要な更新も行った
    }
  }
  
  return false; // イベントは存在しない
}

/**
 * Togglで削除されたタイムエントリに対応するGoogleカレンダーのイベントを削除する関数
 */
function deleteRemovedEntries() {
  return retry(() => {
    var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    if (!calendar) {
      log(LOG_LEVELS.ERROR, `Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
      throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    }
    
    // 過去1ヶ月間のイベントを対象に削除チェックを行う
    var now = new Date();
    var pastDate = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
    
    var events = calendar.getEvents(pastDate, now);
    
    events.forEach(function(event) {
      var title = event.getTitle();
      var match = title.match(/ID:(\d+)$/); // タイトルの末尾にIDがあることを確認
      if (match && match[1]) {
        var record_id = match[1];
        // Toggl APIでこのIDのタイムエントリが存在するか確認
        var exists = checkIfTogglEntryExists(record_id);
        if (!exists) {
          event.deleteEvent();
          log(LOG_LEVELS.INFO, `Deleted event for removed Toggl entry ID:${record_id}`);
        }
      }
    });
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Toggl APIで指定したIDのタイムエントリが存在するか確認する関数
 * @param {string} record_id - タイムエントリのID
 * @returns {boolean} 存在する場合はtrue、存在しない場合はfalse
 */
function checkIfTogglEntryExists(record_id) {
  return retry(() => {
    var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries/${record_id}`;
    
    var response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      return true;
    } else if (responseCode === 404) {
      return false;
    } else {
      log(LOG_LEVELS.ERROR, `Unexpected API response code when checking entry existence: ${responseCode}`);
      throw new Error(`Unexpected response code: ${responseCode}`);
    }
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Googleカレンダー内の重複イベントを削除する関数
 * 既に存在するイベントの中から最新の1つを残し、他を削除します。
 */
function removeDuplicateEvents() {
  return retry(() => {
    var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    var now = new Date();
    var pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate()); // 過去3ヶ月間
    
    var events = calendar.getEvents(pastDate, now);
    
    var eventMap = {};
    
    events.forEach(function(event) {
      var title = event.getTitle();
      var match = title.match(/ID:(\d+)$/);
      if (match && match[1]) {
        var record_id = match[1];
        if (eventMap[record_id]) {
          // 既に存在する場合、最新のイベント以外を削除
          event.deleteEvent();
          log(LOG_LEVELS.INFO, `Deleted duplicate event for ID:${record_id}`);
        } else {
          eventMap[record_id] = event;
        }
      }
    });
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * テスト用の重複イベント作成関数
 * カレンダーに同一IDの重複イベントを2つ作成します。
 */
function testCreateDuplicateEvents() {
  try {
    var now = new Date();
    var oneHourLater = new Date(now.getTime() + 60 * 60 * 1000); // 現在時刻 + 1時間
    var twoHoursLater = new Date(now.getTime() + 2 * 60 * 60 * 1000); // 現在時刻 + 2時間
    
    var title = "テストイベント ID:654321";
    
    // 最初のイベント作成
    recordActivityLog(title, oneHourLater.toISOString(), twoHoursLater.toISOString());
    
    // 2つ目の重複イベント作成（異なる時間で）
    var threeHoursLater = new Date(now.getTime() + 3 * 60 * 60 * 1000); // 現在時刻 + 3時間
    var fourHoursLater = new Date(now.getTime() + 4 * 60 * 60 * 1000); // 現在時刻 + 4時間
    recordActivityLog(title, threeHoursLater.toISOString(), fourHoursLater.toISOString());
    
    log(LOG_LEVELS.INFO, "重複イベントの作成テストが成功しました。");
    SpreadsheetApp.getUi().alert("重複イベントを2つカレンダーに作成しました。\nタイトル: " + title);
  } catch (e) {
    log(LOG_LEVELS.ERROR, "重複イベントの作成テストに失敗しました: " + e.message);
    notifyError(e, '654321');
    SpreadsheetApp.getUi().alert("重複イベントの作成に失敗しました。ログを確認してください。");
  }
}

/**
 * テスト用の重複イベント削除関数
 */
function testRemoveDuplicateEvents() {
  try {
    removeDuplicateEvents();
    log(LOG_LEVELS.INFO, "removeDuplicateEvents関数のテストが成功しました。");
    SpreadsheetApp.getUi().alert("重複イベントの削除テストが成功しました。");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "removeDuplicateEvents関数のテストに失敗しました: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("重複イベントの削除テストに失敗しました。ログを確認してください。");
  }
}

/**
 * 統合テスト関数
 * 重複イベントを作成し、その後削除する一連のテストを実行します。
 */
function testDuplicateEventsWorkflow() {
  try {
    // 重複イベントの作成
    testCreateDuplicateEvents();
    
    // 少し待機してから削除（必要に応じて）
    Utilities.sleep(2000); // 2秒待機
    
    // 重複イベントの削除
    testRemoveDuplicateEvents();
    
    SpreadsheetApp.getUi().alert("重複イベントの作成と削除のテストが完了しました。");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "統合テストに失敗しました: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("統合テストに失敗しました。ログを確認してください。");
  }
}

/**
 * メイン関数: Togglからデータを取得し、Googleカレンダーにイベントを記録・更新・削除
 */
function watch() {
  var lock;
  try {
    // LockServiceを使用して並行実行を防止
    lock = getLock();
    
    var check_datetime = getLastModifyDatetime();
    
    if (check_datetime == -1) {
      var initial_date = new Date();
      initial_date.setUTCHours(0,0,0,0);
      var initial_timestamp = Math.floor(initial_date.getTime() / 1000);
      log(LOG_LEVELS.INFO, "No cached data found, using initial timestamp: " + initial_timestamp);
      check_datetime = initial_timestamp;
    } else {
      log(LOG_LEVELS.INFO, "Data is being fetched starting from: " + new Date(check_datetime * 1000).toISOString());
    }
    
    var days_to_fetch = 3;
    var time_entries = getTimeEntries(check_datetime, days_to_fetch);
    
    if (time_entries && time_entries.length > 0) {
      log(LOG_LEVELS.INFO, "Number of time entries fetched: " + time_entries.length);
      var last_stop_datetime = check_datetime;
      
      // カレンダーから既存のイベントを取得
      var eventMap = getAllRelevantEvents();
      
      time_entries.forEach(function(record) {
        // データの検証
        if (!record.stop) {
          log(LOG_LEVELS.DEBUG, "Record with no stop time: " + JSON.stringify(record));
          return;
        }
        
        var stop_time = Math.floor(new Date(record.stop).getTime() / 1000);
        if (isNaN(stop_time)) {
          log(LOG_LEVELS.DEBUG, "Invalid stop time: " + record.stop);
          return;
        }
        
        var start_time = Math.floor(new Date(record.start).getTime() / 1000);
        if (isNaN(start_time)) {
          log(LOG_LEVELS.DEBUG, "Invalid start time: " + record.start);
          return;
        }
        
        var project_data = getProjectData(record.wid, record.pid);
        var project_name = project_data.name || '';
        var activity_log = [(record.description || '名称なし'), project_name].filter(Boolean).join(" : ") + ` ID:${record.id}`;
        
        var exists = eventExistsAndUpdate(record.id, record);
        
        if (!exists) {
          try {
            recordActivityLog(
              activity_log,
              record.start,
              record.stop
            );
            log(LOG_LEVELS.INFO, "Added event: " + activity_log);
          } catch (e) {
            log(LOG_LEVELS.ERROR, `Error adding event for ID: ${record.id} - ${e}`);
            notifyError(e, record.id);
          }
        } else {
          log(LOG_LEVELS.DEBUG, "Existing event processed for ID: " + record.id);
        }
        
        // 最後の停止時間を更新
        if (stop_time > last_stop_datetime) {
          last_stop_datetime = stop_time;
        }
      });
      
      log(LOG_LEVELS.INFO, "Updating cache with last stop datetime: " + last_stop_datetime);
      
      if (last_stop_datetime) {
        var new_timestamp = last_stop_datetime + 1;
        putLastModifyDatetime(new_timestamp);
        log(LOG_LEVELS.INFO, "Cache updated after fetching data with timestamp: " + new_timestamp);
      } else {
        log(LOG_LEVELS.INFO, "No valid stop time found, cache not updated.");
      }
    } else {
      log(LOG_LEVELS.INFO, "No time entries found");
    }
  } catch (e) {
    Logger.log(e);
    notifyError(e);
  } finally {
    if (lock) {
      lock.releaseLock();
      log(LOG_LEVELS.INFO, "Lock released");
    }
  }
}

/**
 * ==== Optional: Remove Duplicate Events Trigger ====
 * この関数は必要に応じてトリガーから実行してください。
 */
// function scheduleRemoveDuplicateEvents() {
//   ScriptApp.newTrigger('removeDuplicateEvents')
//     .timeBased()
//     .everyDays(1) // 毎日実行
//     .atHour(2) // 午前2時に実行
//     .create();
// }

/**
 * ==== Test Functions ====
 * テスト用の関数を追加しています。必要に応じて実行してください。
 */

/**
 * テスト用のイベント作成関数
 * カレンダーに1時間後から2時間後までのテストイベントを作成します。
 */
function testRecordActivityLog() {
  try {
    var now = new Date();
    var oneHourLater = new Date(now.getTime() + 60 * 60 * 1000); // 現在時刻 + 1時間
    var twoHoursLater = new Date(now.getTime() + 2 * 60 * 60 * 1000); // 現在時刻 + 2時間

    var title = "テストイベント ID:123456";
    var started_at = oneHourLater.toISOString();
    var ended_at = twoHoursLater.toISOString();

    recordActivityLog(title, started_at, ended_at);
    log(LOG_LEVELS.INFO, "recordActivityLog関数のテストが成功しました。");
    SpreadsheetApp.getUi().alert("テストイベントをカレンダーに作成しました。\n開始時刻: " + started_at + "\n終了時刻: " + ended_at);
  } catch (e) {
    log(LOG_LEVELS.ERROR, "recordActivityLog関数のテストに失敗しました: " + e.message);
    notifyError(e, '123456');
    SpreadsheetApp.getUi().alert("テストイベントの作成に失敗しました。ログを確認してください。");
  }
}

/**
 * テスト用の重複イベント削除関数
 */
function testRemoveDuplicateEvents() {
  try {
    removeDuplicateEvents();
    log(LOG_LEVELS.INFO, "removeDuplicateEvents関数のテストが成功しました。");
    SpreadsheetApp.getUi().alert("重複イベントの削除テストが成功しました。");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "removeDuplicateEvents関数のテストに失敗しました: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("重複イベントの削除テストに失敗しました。ログを確認してください。");
  }
}

/**
 * 統合テスト関数
 * 重複イベントを作成し、その後削除する一連のテストを実行します。
 */
function testDuplicateEventsWorkflow() {
  try {
    // 重複イベントの作成
    testCreateDuplicateEvents();
    
    // 少し待機してから削除（必要に応じて）
    Utilities.sleep(2000); // 2秒待機
    
    // 重複イベントの削除
    testRemoveDuplicateEvents();
    
    SpreadsheetApp.getUi().alert("重複イベントの作成と削除のテストが完了しました。");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "統合テストに失敗しました: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("統合テストに失敗しました。ログを確認してください。");
  }
}