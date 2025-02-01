/*
  Toggl time entries export to GoogleCalendar
  original author: Masato Kawaguchi
  modified by: chess
  Released under the BSD-3-Clause license
  version: 1.4.03
  https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE

  CHANGELOG:
    v1.4.03 (2025/02/01):
      - 初回実行時の分割取得・進捗保存機能を追加（処理途中の進捗を永続的に Script Properties と一時的な CacheService に保存）
      - 手動実行モードを3種類実装：タイムアウトモード（1分区切りで中断、ユーザー再実行）、完遂モード（自動再開、タイムアウト閾値5分30秒）、初回実行モード（キャッシュ無視）
      - 全実行モードでロック(getLock)を取得し、try…finally で必ずロック解放することで、手動実行と自動実行（watch）が排他的に動作するように改善
      - 進捗状況のログは、処理開始時、タイムアウト時、完了時にのみ出力
      - さらに、元々削除していた機能（削除チェック、テスト用関数、キャッシュクリア・テストメール送信機能）を復元
*/

/** CONFIG: 基本設定 **/
const CONFIG = {
  CACHE_KEY: 'toggl_exporter:lastmodify_datetime',
  TIME_OFFSET: 9 * 60 * 60, // JST (秒)
  TOGGL_API_HOSTNAME: 'https://api.track.toggl.com',
  GOOGLE_CALENDAR_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_CALENDAR_ID'),
  NOTIFICATION_EMAIL: PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL'),
  TOGGL_BASIC_AUTH: PropertiesService.getScriptProperties().getProperty('TOGGL_BASIC_AUTH'),
  DEBUG_MODE: false,  // DEBUG_MODE を true にすると、DEBUGレベルの詳細ログも出力されます（本番環境では false）
  MAX_CACHE_AGE: 12 * 60 * 60, // 12時間（秒）※本実装では有効期限チェックは省略し、永続保存として扱う
  RETRY_COUNT: 5,
  RETRY_DELAY: 2000, // ms
  
  // タイムアウト閾値（ミリ秒）
  AUTOMATIC_TIMEOUT_INTERVAL: 240000,         // 自動実行：4分
  MANUAL_COMPLETE_TIMEOUT_INTERVAL: 330000,     // 手動完遂モード：5分30秒
  MANUAL_TIMEOUT_MODE_INTERVAL: 60000           // 手動タイムアウトモード：1分
};

/** ログレベル **/
const LOG_LEVELS = {
  DEBUG: 1,
  INFO: 2,
  ERROR: 3,
};
const CURRENT_LOG_LEVEL = CONFIG.DEBUG_MODE ? LOG_LEVELS.DEBUG : LOG_LEVELS.INFO;

/**
 * ログ出力用関数
 */
function log(level, message) {
  if (level >= CURRENT_LOG_LEVEL) {
    Logger.log(message);
  }
}

/**
 * ロックを取得する関数（最大30秒待機）
 */
function getLock() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    log(LOG_LEVELS.INFO, "Lock acquired");
    return lock;
  } catch (e) {
    log(LOG_LEVELS.ERROR, "Could not acquire lock: " + e);
    throw e;
  }
}

/**
 * キャッシュ（最終更新日時）の取得：
 * まず CacheService を確認し、なければ Script Properties を参照する
 * キャッシュが無い場合や形式不正の場合は -1 を返す
 */
function getLastModifyDatetime() {
  const cache = CacheService.getScriptCache();
  let cachedData = cache.get(CONFIG.CACHE_KEY);
  
  if (cachedData) {
    try {
      const parsedData = JSON.parse(cachedData);
      if (parsedData && typeof parsedData.timestamp === 'number') {
        log(LOG_LEVELS.DEBUG, "CacheService hit: " + parsedData.timestamp);
        return parsedData.timestamp;
      }
    } catch (e) {
      log(LOG_LEVELS.ERROR, "Error parsing CacheService data: " + e.message);
    }
  }
  
  // Fallback: Script Properties
  const props = PropertiesService.getScriptProperties();
  cachedData = props.getProperty(CONFIG.CACHE_KEY);
  if (cachedData) {
    try {
      const parsedData = JSON.parse(cachedData);
      if (parsedData && typeof parsedData.timestamp === 'number') {
        log(LOG_LEVELS.DEBUG, "ScriptProperties hit: " + parsedData.timestamp);
        // 同じデータを CacheService にも更新
        cache.put(CONFIG.CACHE_KEY, cachedData, CONFIG.MAX_CACHE_AGE);
        return parsedData.timestamp;
      }
    } catch (e) {
      log(LOG_LEVELS.ERROR, "Error parsing ScriptProperties data: " + e.message);
    }
  }
  
  log(LOG_LEVELS.DEBUG, "No cached data found");
  return -1;
}

/**
 * キャッシュ（最終更新日時）を Script Properties と CacheService の両方に保存する関数
 */
function putLastModifyDatetime(timestamp) {
  const data = JSON.stringify({ timestamp: timestamp });
  const cache = CacheService.getScriptCache();
  cache.put(CONFIG.CACHE_KEY, data, CONFIG.MAX_CACHE_AGE);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.CACHE_KEY, data);
  log(LOG_LEVELS.DEBUG, "Cache updated with timestamp: " + data);
}

/**
 * retry ユーティリティ関数
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
 * Toggl API から指定期間のタイムエントリを取得する関数
 */
function getTimeEntriesRange(startIso, endIso) {
  return retry(() => {
    const uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries?start_date=${encodeURIComponent(startIso)}&end_date=${encodeURIComponent(endIso)}`;
    log(LOG_LEVELS.DEBUG, `Fetching time entries from: ${uri}`);
    
    const response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    log(LOG_LEVELS.DEBUG, `API Response Code: ${responseCode}`);
    log(LOG_LEVELS.DEBUG, `API Response: ${responseText}`);
    
    if (responseCode !== 200) {
      log(LOG_LEVELS.ERROR, `API Error: ${responseText}`);
      throw new Error(`Toggl API returned status code ${responseCode}`);
    }
    
    const parsed = JSON.parse(responseText);
    if (Array.isArray(parsed)) {
      return parsed;
    } else {
      log(LOG_LEVELS.ERROR, "API returned non-array response");
      return null;
    }
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Googleカレンダーにイベントを記録する関数
 */
function recordActivityLog(title, started_at, ended_at) {
  log(LOG_LEVELS.DEBUG, `recordActivityLog called with title: "${title}", started_at: "${started_at}", ended_at: "${ended_at}"`);
  
  const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  if (!calendar) {
    log(LOG_LEVELS.ERROR, `Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
  }
  
  const startDate = new Date(started_at);
  const endDate = new Date(ended_at);
  
  if (isNaN(startDate.getTime())) {
    log(LOG_LEVELS.ERROR, `Invalid start date: "${started_at}"`);
    throw new RangeError(`Invalid start date: "${started_at}"`);
  }
  if (isNaN(endDate.getTime())) {
    log(LOG_LEVELS.ERROR, `Invalid end date: "${ended_at}"`);
    throw new RangeError(`Invalid end date: "${ended_at}"`);
  }
  
  try {
    calendar.createEvent(title, startDate, endDate);
    log(LOG_LEVELS.INFO, `Created event: "${title}" from ${started_at} to ${ended_at}`);
  } catch (createError) {
    log(LOG_LEVELS.ERROR, `Error creating event: ${createError}`);
    throw createError;
  }
}

/**
 * 既存イベントの検索と更新（重複防止）
 * ※ Toggl のタイムエントリに固有のIDがある場合のみ対象とする
 */
function eventExistsAndUpdate(record_id, newRecord) {
  if (!record_id) {
    // IDが無い場合は従来の動作（新規登録）とする
    return false;
  }
  const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  const now = new Date();
  const pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
  const events = calendar.getEvents(pastDate, now, { search: `ID:${record_id}` });
  
  log(LOG_LEVELS.DEBUG, `Searching for events with ID:${record_id}. Found ${events.length} events.`);
  
  if (events.length > 0) {
    let matchingEvent = null;
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const titleMatch = event.getTitle().match(/ID:(\d+)$/);
      if (titleMatch && titleMatch[1] === String(record_id)) {
        matchingEvent = event;
        break;
      }
    }
    
    if (matchingEvent) {
      log(LOG_LEVELS.DEBUG, `Matching event found: "${matchingEvent.getTitle()}"`);
      const newStart = newRecord.start;
      const newEnd = newRecord.stop;
      
      const project_data = getProjectData(newRecord.wid, newRecord.pid);
      const project_name = project_data.name || '';
      const updatedTitle = [(newRecord.description || '名称なし'), project_name]
        .filter(Boolean).join(" : ") + ` ID:${record_id}`;
      
      const eventTitleNeedsUpdate = (matchingEvent.getTitle() !== updatedTitle);
      const eventTimeNeedsUpdate = (matchingEvent.getStartTime().toISOString() !== newStart) ||
                                     (matchingEvent.getEndTime().toISOString() !== newEnd);
      
      log(LOG_LEVELS.DEBUG, `EventTitleNeedsUpdate: ${eventTitleNeedsUpdate}, EventTimeNeedsUpdate: ${eventTimeNeedsUpdate}`);
      
      if (eventTitleNeedsUpdate || eventTimeNeedsUpdate) {
        try {
          retry(() => {
            matchingEvent.setTitle(updatedTitle);
            matchingEvent.setTime(new Date(newStart), new Date(newEnd));
          }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
          log(LOG_LEVELS.INFO, `Updated event for ID:${record_id}`);
        } catch (e) {
          log(LOG_LEVELS.ERROR, `Error updating event for ID:${record_id} - ${e}`);
          notifyError(e, record_id);
          return false;
        }
      } else {
        log(LOG_LEVELS.DEBUG, `No update needed for event ID:${record_id}`);
      }
      
      return true;
    }
  }
  
  return false;
}

/**
 * Toggl のプロジェクトデータ取得（IDベース）
 */
function getProjectData(workspace_id, project_id) {
  if (!workspace_id || !project_id) return {};
  return retry(() => {
    const uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/workspaces/${workspace_id}/projects/${project_id}`;
    
    const response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
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
 * 削除チェック：Toggl API で指定IDのタイムエントリが存在するか確認する関数
 */
function checkIfTogglEntryExists(record_id) {
  return retry(() => {
    const uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries/${record_id}`;
    const response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
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
 * Togglで削除されたエントリをカレンダーから削除する(過去1日版)
 * - 通常トリガー等で用い、低負荷運用
 */
function deleteRemovedEntriesShort() {
  return retry(() => {
    const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    if (!calendar) {
      throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    }
    const now = new Date();
    const oneDayMs = 1 * 24 * 60 * 60 * 1000;
    const pastDate = new Date(now.getTime() - oneDayMs);
    const events = calendar.getEvents(pastDate, now);
    events.forEach(function(event) {
      const title = event.getTitle();
      const match = title.match(/ID:(\d+)$/);
      if (match && match[1]) {
        const record_id = match[1];
        const exists = checkIfTogglEntryExists(record_id);
        if (!exists) {
          event.deleteEvent();
          log(LOG_LEVELS.INFO, `Deleted event (short range) for removed Toggl entry ID:${record_id}`);
        }
      }
    });
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Togglで削除されたエントリをカレンダーから削除する(過去1ヶ月版)
 * - 手動実行用、より古い削除を拾うため
 */
function deleteRemovedEntriesManual() {
  return retry(() => {
    const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    if (!calendar) {
      throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    }
    const now = new Date();
    const pastDate = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
    const events = calendar.getEvents(pastDate, now);
    events.forEach(function(event) {
      const title = event.getTitle();
      const match = title.match(/ID:(\d+)$/);
      if (match && match[1]) {
        const record_id = match[1];
        const exists = checkIfTogglEntryExists(record_id);
        if (!exists) {
          event.deleteEvent();
          log(LOG_LEVELS.INFO, `Deleted event (manual range) for removed Toggl entry ID:${record_id}`);
        }
      }
    });
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * 重複イベントを削除する (過去3ヶ月)
 * - 同じIDを持つ複数イベントがある場合、最新以外を削除
 */
function removeDuplicateEvents() {
  return retry(() => {
    const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    const now = new Date();
    const pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
    const events = calendar.getEvents(pastDate, now);
    const eventMap = {};
    events.forEach(function(event) {
      const title = event.getTitle();
      const match = title.match(/ID:(\d+)$/);
      if (match && match[1]) {
        const record_id = match[1];
        if (eventMap[record_id]) {
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
    log(LOG_LEVELS.INFO, `Test email sent to: ${email}`);
  } else {
    log(LOG_LEVELS.ERROR, "Notification email address is not configured.");
  }
}

/**
 * キャッシュをクリアする関数
 */
function clearScriptCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
  log(LOG_LEVELS.INFO, "Cache cleared");
  SpreadsheetApp.getUi().alert("キャッシュをクリアしました。");
}

/**
 * テスト用の重複イベント作成関数
 */
function testCreateDuplicateEvents() {
  try {
    var now = new Date();
    var oneHourLater = new Date(now.getTime() + 60 * 60 * 1000);
    var twoHoursLater = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    var title = "テストイベント ID:654321";
    
    recordActivityLog(title, oneHourLater.toISOString(), twoHoursLater.toISOString());
    
    var threeHoursLater = new Date(now.getTime() + 3 * 60 * 60 * 1000);
    var fourHoursLater  = new Date(now.getTime() + 4 * 60 * 60 * 1000);
    recordActivityLog(title, threeHoursLater.toISOString(), fourHoursLater.toISOString());
    
    log(LOG_LEVELS.INFO, "重複イベント作成テスト成功");
    SpreadsheetApp.getUi().alert("重複イベントを2件作成しました。\nタイトル: " + title);
  } catch (e) {
    log(LOG_LEVELS.ERROR, "重複イベント作成テスト失敗: " + e.message);
    notifyError(e, '654321');
    SpreadsheetApp.getUi().alert("重複イベント作成テスト失敗。ログを確認してください。");
  }
}

/**
 * テスト用の重複イベント削除関数
 */
function testRemoveDuplicateEvents() {
  try {
    removeDuplicateEvents();
    log(LOG_LEVELS.INFO, "重複イベント削除テスト成功");
    SpreadsheetApp.getUi().alert("重複イベント削除テスト成功");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "重複イベント削除テスト失敗: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("重複イベント削除テスト失敗。ログを確認してください。");
  }
}

/**
 * 統合テスト関数: 重複イベント作成→削除
 */
function testDuplicateEventsWorkflow() {
  try {
    testCreateDuplicateEvents();
    Utilities.sleep(2000);
    testRemoveDuplicateEvents();
    SpreadsheetApp.getUi().alert("重複イベント作成と削除の統合テスト完了");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "統合テスト失敗: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("統合テスト失敗。ログを確認してください。");
  }
}

/**
 * ----------------------------
 * タイムアウト対策＆実行モード分岐（進捗保存機構付き）
 * ----------------------------
 */

// 進捗管理用のキー（Script Properties を利用）
const PROGRESS_KEY = 'toggl_exporter:last_processed_index';

/**
 * バッチ処理のコア関数
 * @param {boolean} isManual   手動実行なら true、定期実行なら false
 * @param {boolean} autoResume 手動実行時に自動再開する場合は true、タイムアウトで中断する場合は false
 * @param {boolean} forceInitial true の場合、キャッシュを無視して初回実行（過去30日分取得）とする
 */
function processTimeEntriesBatch(isManual, autoResume, forceInitial) {
  forceInitial = forceInitial || false;
  
  // ロックを取得（同時実行防止）
  let lock = getLock();
  try {
    // モード別タイムアウト閾値の設定（ミリ秒）
    var MAX_EXECUTION_TIME;
    if (isManual) {
      MAX_EXECUTION_TIME = autoResume ? CONFIG.MANUAL_COMPLETE_TIMEOUT_INTERVAL : CONFIG.MANUAL_TIMEOUT_MODE_INTERVAL;
    } else {
      MAX_EXECUTION_TIME = CONFIG.AUTOMATIC_TIMEOUT_INTERVAL;
    }
    
    var startTime = new Date().getTime();
    var props = PropertiesService.getScriptProperties();
    var lastIndex = parseInt(props.getProperty(PROGRESS_KEY)) || 0;
    
    var lastModify = forceInitial ? -1 : getLastModifyDatetime();
    var now = new Date();
    var startDate;
    if (lastModify === -1) {
      startDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
      log(LOG_LEVELS.INFO, "初回実行: 過去30日分のデータを取得します");
    } else {
      startDate = new Date((lastModify - 24 * 60 * 60) * 1000);
      log(LOG_LEVELS.INFO, "継続実行: キャッシュのタイムスタンプに基づいてデータ取得を行います");
    }
    
    var startIso = startDate.toISOString();
    var endIso = now.toISOString();
    
    var timeEntries = getTimeEntriesRange(startIso, endIso);
    if (!timeEntries) {
      log(LOG_LEVELS.ERROR, "タイムエントリの取得に失敗しました");
      return;
    }
    log(LOG_LEVELS.INFO, "Number of time entries fetched: " + timeEntries.length);
    var totalCount = timeEntries.length;
    log(LOG_LEVELS.INFO, "Total records to process: " + totalCount);
    log(LOG_LEVELS.INFO, "Processing starts from index " + lastIndex + " at " + new Date().toISOString());
    
    for (var i = lastIndex; i < totalCount; i++) {
      var record = timeEntries[i];
      if (!record.stop) {
        log(LOG_LEVELS.DEBUG, "Record with no stop time: " + JSON.stringify(record));
        continue;
      }
      
      var stop_time = Math.floor(new Date(record.stop).getTime() / 1000);
      var start_time = Math.floor(new Date(record.start).getTime() / 1000);
      if (isNaN(stop_time) || isNaN(start_time)) {
        log(LOG_LEVELS.DEBUG, "Invalid time for record: " + JSON.stringify(record));
        continue;
      }
      
      try {
        if (!eventExistsAndUpdate(record.id, record)) {
          var project_data = getProjectData(record.wid, record.pid);
          var project_name = project_data.name || '';
          var activity_log = [(record.description || '名称なし'), project_name]
            .filter(Boolean).join(" : ") + " ID:" + record.id;
          recordActivityLog(activity_log, record.start, record.stop);
          log(LOG_LEVELS.INFO, "Added event: " + activity_log);
        } else {
          log(LOG_LEVELS.DEBUG, "Existing event processed for ID: " + record.id);
        }
      } catch (e) {
        log(LOG_LEVELS.ERROR, "Error processing record ID:" + record.id + " - " + e);
        notifyError(e, record.id);
      }
      
      if (stop_time > lastModify) {
        lastModify = stop_time;
      }
      
      var elapsed = new Date().getTime() - startTime;
      if (elapsed > MAX_EXECUTION_TIME) {
        props.setProperty(PROGRESS_KEY, i + 1);
        var processedCount = i + 1;
        var percentComplete = Math.floor((processedCount / totalCount) * 100);
        var remainingCount = totalCount - processedCount;
        log(LOG_LEVELS.INFO, "Timeout reached: Processed " + processedCount + " of " + totalCount +
            " (" + percentComplete + "%). Remaining: " + remainingCount +
            " records. Current record's stop date: " + record.stop);
        
        if (!isManual || (isManual && autoResume)) {
          log(LOG_LEVELS.INFO, (isManual ? "手動完遂" : "自動実行") + ": 閾値に達したため中断します。Next start index: " + (i + 1));
          ScriptApp.newTrigger('watch')
            .timeBased()
            .after(1000)
            .create();
        } else {
          log(LOG_LEVELS.INFO, "手動実行（タイムアウトモード）: 閾値に達しました。続きはユーザーが再実行してください。");
        }
        return;
      }
    }
    
    putLastModifyDatetime(lastModify + 1);
    props.deleteProperty(PROGRESS_KEY);
    log(LOG_LEVELS.INFO, "Processing complete: Processed all " + totalCount + " records at " + new Date().toISOString());
  } finally {
    if (lock) {
      lock.releaseLock();
      log(LOG_LEVELS.INFO, "Lock released");
    }
  }
}

/**
 * 自動実行用エントリポイント（watch） — トリガー経由で呼ばれる
 */
function watch() {
  processTimeEntriesBatch(false, true, false);
}

/**
 * 手動実行（タイムアウトモード）：1分で中断、進捗保存しユーザー再実行
 */
function manualProcessTimeEntriesTimeout() {
  processTimeEntriesBatch(true, false, false);
}

/**
 * 手動実行（完遂モード）：5分30秒で中断しても自動再開
 */
function manualProcessTimeEntriesComplete() {
  processTimeEntriesBatch(true, true, false);
}

/**
 * 手動実行（初回実行モード）：キャッシュ無視で常に初回実行
 */
function manualProcessTimeEntriesInitial() {
  PropertiesService.getScriptProperties().deleteProperty(PROGRESS_KEY);
  processTimeEntriesBatch(true, false, true);
}

/**
 * onOpen: カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('キャッシュをクリア', 'clearScriptCache')
    .addItem('テストメールを送信', 'sendTestEmail')
    .addItem('削除(短期間)', 'deleteRemovedEntriesShort')
    .addItem('削除(長期間)', 'deleteRemovedEntriesManual')
    .addItem('重複イベント作成テスト', 'testCreateDuplicateEvents')
    .addItem('重複イベント削除テスト', 'testRemoveDuplicateEvents')
    .addItem('重複イベント統合テスト', 'testDuplicateEventsWorkflow')
    .addItem('手動同期（タイムアウトモード）', 'manualProcessTimeEntriesTimeout')
    .addItem('手動同期（完遂モード）', 'manualProcessTimeEntriesComplete')
    .addItem('手動同期（初回実行モード）', 'manualProcessTimeEntriesInitial')
    .addToUi();
}