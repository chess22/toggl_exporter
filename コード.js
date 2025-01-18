/*
  Toggl time entries export to GoogleCalendar
  original author: Masato Kawaguchi
  modified by: chess
  Released under the BSD-3-Clause license
  version: 1.4.02
  https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE

  CHANGELOG:
    v1.4.02 (2025-01-18):
      - 初回実行時に過去30日分を取得
      - 2回目以降は (前回停止時刻 - 1日) から現在までを再取得
      - deleteRemovedEntriesShort() を過去1日に設定し、通常トリガーでの低負荷削除チェック
      - deleteRemovedEntriesManual() を過去1ヶ月に設定し、手動で古い削除にも対応可能
      - 「削除チェック」は Toggl から削除されたエントリをカレンダー側でも削除する機能
*/

/**
 * TogglとGoogleカレンダーを同期するGoogle Apps Script
 *
 * 機能:
 * 1. Togglからタイムエントリを取得（初回30日、2回目以降は前回停止時刻-1日）
 * 2. Googleカレンダーにイベントを作成・更新
 * 3. 重複イベントの作成を防止・削除
 * 4. Togglで削除されたエントリをカレンダーからも削除 (deleteRemovedEntriesXxx)
 * 5. エラーハンドリングと通知
 * 6. カスタムメニューによる手動操作の提供
 */

const CONFIG = {
  CACHE_KEY: 'toggl_exporter:lastmodify_datetime',
  TIME_OFFSET: 9 * 60 * 60, // JST (秒)
  TOGGL_API_HOSTNAME: 'https://api.track.toggl.com',
  GOOGLE_CALENDAR_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_CALENDAR_ID'),
  NOTIFICATION_EMAIL: PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL'),
  TOGGL_BASIC_AUTH: PropertiesService.getScriptProperties().getProperty('TOGGL_BASIC_AUTH'),
  DEBUG_MODE: false,
  MAX_CACHE_AGE: 12 * 60 * 60, // 12時間（秒）
  RETRY_COUNT: 5,
  RETRY_DELAY: 2000, // ms
};

/**
 * ログレベル
 */
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
 * 指定した関数を再試行するユーティリティ関数
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
  const cache = CacheService.getScriptCache();
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
    .addItem('削除(短期間)', 'deleteRemovedEntriesShort')    // 過去1日
    .addItem('削除(長期間)', 'deleteRemovedEntriesManual')   // 過去1ヶ月
    .addItem('重複イベント作成テスト', 'testCreateDuplicateEvents')
    .addItem('重複イベント削除テスト', 'testRemoveDuplicateEvents')
    .addItem('重複イベント統合テスト', 'testDuplicateEventsWorkflow')
    .addToUi();
}

/**
 * キャッシュから最後の更新日時を取得
 */
function getLastModifyDatetime() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(CONFIG.CACHE_KEY);
  
  if (!cachedData) {
    log(LOG_LEVELS.DEBUG, "No cached data found");
    return -1;
  }
  
  log(LOG_LEVELS.DEBUG, "Cached data: " + cachedData);
  
  try {
    const parsedData = JSON.parse(cachedData);
    if (parsedData && typeof parsedData.timestamp === 'number') {
      const cacheAge = (Date.now() / 1000) - parsedData.timestamp;
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
 */
function putLastModifyDatetime(unix_timestamp) {
  const cache = CacheService.getScriptCache();
  const data = JSON.stringify({ timestamp: unix_timestamp });
  cache.put(CONFIG.CACHE_KEY, data, CONFIG.MAX_CACHE_AGE);
  log(LOG_LEVELS.DEBUG, "Cache updated with timestamp: " + data);
}

/**
 * TogglのタイムエントリをISO期間指定で取得する関数
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
  const endDate   = new Date(ended_at);
  
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
 * 既存イベントを検索し必要に応じて更新(重複防止・変更反映)
 */
function eventExistsAndUpdate(record_id, newRecord) {
  const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  
  // 過去3ヶ月を検索するのは、更新が2か月前にあった場合にも対応するため
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
      const newEnd   = newRecord.stop;
      
      const project_data = getProjectData(newRecord.wid, newRecord.pid);
      const project_name = project_data.name || '';
      const updatedTitle = [(newRecord.description || '名称なし'), project_name]
        .filter(Boolean).join(" : ") + ` ID:${record_id}`;
      
      const eventTitleNeedsUpdate = (matchingEvent.getTitle() !== updatedTitle);
      const eventTimeNeedsUpdate  = (matchingEvent.getStartTime().toISOString() !== newStart)
                                 || (matchingEvent.getEndTime().toISOString() !== newEnd);
      
      log(LOG_LEVELS.DEBUG, `EventTitleNeedsUpdate: ${eventTitleNeedsUpdate}, EventTimeNeedsUpdate: ${eventTimeNeedsUpdate}`);
      
      if (eventTitleNeedsUpdate || eventTimeNeedsUpdate) {
        try {
          retry(() => {
            matchingEvent.setTitle(updatedTitle);
            matchingEvent.setTime(new Date(newStart), new Date(newEnd));
          }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
          log(LOG_LEVELS.INFO, `Updated event for ID:${record_id}`);
        } catch (e) {
          log(LOG_LEVELS.ERROR, `Error updating event for ID: ${record_id} - ${e}`);
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
 * Toggl APIからプロジェクトデータを取得 (IDベース)
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
 * Togglで削除されたエントリをカレンダーから削除する(過去1日版)
 * - 通常トリガー等で用いて、低負荷運用
 */
function deleteRemovedEntriesShort() {
  return retry(() => {
    const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    if (!calendar) {
      throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    }
    
    // 「過去1日」の範囲
    const now = new Date();
    const oneDayMs = 1 * 24 * 60 * 60 * 1000;
    const pastDate = new Date(now.getTime() - oneDayMs);

    const events = calendar.getEvents(pastDate, now);

    events.forEach(function(event) {
      const title = event.getTitle();
      const match = title.match(/ID:(\d+)$/);
      if (match && match[1]) {
        const record_id = match[1];
        // TogglでこのIDが存在するかチェック
        const exists = checkIfTogglEntryExists(record_id);
        if (!exists) {
          // Toggl側で削除された → カレンダーからも削除
          event.deleteEvent();
          log(LOG_LEVELS.INFO, `Deleted event (short range) for removed Toggl entry ID:${record_id}`);
        }
      }
    });
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * Togglで削除されたエントリをカレンダーから削除(過去1ヶ月版)
 * - 手動実行用、より古い削除を拾いたい場合に使う
 */
function deleteRemovedEntriesManual() {
  return retry(() => {
    const calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
    if (!calendar) {
      throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    }
    
    // 「過去1ヶ月」の範囲
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
 * Toggl APIで指定IDのタイムエントリが存在するか確認
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
      return true;   // 存在する
    } else if (responseCode === 404) {
      return false;  // 削除された
    } else {
      log(LOG_LEVELS.ERROR, `Unexpected API response code when checking entry existence: ${responseCode}`);
      throw new Error(`Unexpected response code: ${responseCode}`);
    }
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
 * メイン同期 (watch)
 * - 初回: 過去30日
 * - 2回目以降: (前回停止時刻 - 1日) 〜 現在
 */
function watch() {
  let lock;
  try {
    lock = getLock();
    let check_datetime = getLastModifyDatetime(); // 前回の最終停止時刻(UNIX秒)
    let startDate, endDate;
    
    if (check_datetime === -1) {
      // 初回: 過去30日を取得
      log(LOG_LEVELS.INFO, "No cached data found; fetching past 30 days");
      const now = new Date();
      endDate = now;
      startDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); // 30日前
    } else {
      // 2回目以降: (前回停止時刻 - 1日) 〜 現在
      log(LOG_LEVELS.INFO, "Cached data found");
      const overlapSeconds = 1 * 24 * 60 * 60; // 1日
      const now = new Date();
      endDate = now;
      let startUnix = check_datetime - overlapSeconds;
      if (startUnix < 0) startUnix = 0;
      startDate = new Date(startUnix * 1000);
      
      log(LOG_LEVELS.INFO, "Fetching data from: " + startDate.toISOString() + " to " + endDate.toISOString());
    }
    
    const startIso = startDate.toISOString();
    const endIso   = endDate.toISOString();
    
    const time_entries = getTimeEntriesRange(startIso, endIso);
    
    if (time_entries && time_entries.length > 0) {
      log(LOG_LEVELS.INFO, "Number of time entries fetched: " + time_entries.length);
      let last_stop_datetime = (check_datetime === -1) ? 0 : check_datetime;
      
      time_entries.forEach(record => {
        if (!record.stop) {
          log(LOG_LEVELS.DEBUG, "Record with no stop time: " + JSON.stringify(record));
          return;
        }
        
        let stop_time = Math.floor(new Date(record.stop).getTime() / 1000);
        if (isNaN(stop_time)) {
          log(LOG_LEVELS.DEBUG, "Invalid stop time: " + record.stop);
          return;
        }
        
        let start_time = Math.floor(new Date(record.start).getTime() / 1000);
        if (isNaN(start_time)) {
          log(LOG_LEVELS.DEBUG, "Invalid start time: " + record.start);
          return;
        }
        
        let project_data = getProjectData(record.wid, record.pid);
        let project_name = project_data.name || '';
        let activity_log = [(record.description || '名称なし'), project_name]
          .filter(Boolean).join(" : ") + ` ID:${record.id}`;
        
        let exists = eventExistsAndUpdate(record.id, record);
        
        if (!exists) {
          try {
            recordActivityLog(activity_log, record.start, record.stop);
            log(LOG_LEVELS.INFO, "Added event: " + activity_log);
          } catch (e) {
            log(LOG_LEVELS.ERROR, `Error adding event for ID: ${record.id} - ${e}`);
            notifyError(e, record.id);
          }
        } else {
          log(LOG_LEVELS.DEBUG, "Existing event processed for ID: " + record.id);
        }
        
        if (stop_time > last_stop_datetime) {
          last_stop_datetime = stop_time;
        }
      });
      
      log(LOG_LEVELS.INFO, "Updating cache with last stop datetime: " + last_stop_datetime);
      if (last_stop_datetime) {
        const new_timestamp = last_stop_datetime + 1;
        putLastModifyDatetime(new_timestamp);
        log(LOG_LEVELS.INFO, "Cache updated with new timestamp: " + new_timestamp);
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
    
    log(LOG_LEVELS.INFO, "重複イベントの作成テストが成功しました。");
    SpreadsheetApp.getUi().alert("重複イベントを2つ作成しました。\nタイトル: " + title);
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
 * 統合テスト関数: 重複イベントを作成→削除
 */
function testDuplicateEventsWorkflow() {
  try {
    testCreateDuplicateEvents();
    Utilities.sleep(2000);
    testRemoveDuplicateEvents();
    SpreadsheetApp.getUi().alert("重複イベントの作成と削除のテストが完了しました。");
  } catch (e) {
    log(LOG_LEVELS.ERROR, "統合テストに失敗しました: " + e.message);
    notifyError(e);
    SpreadsheetApp.getUi().alert("統合テストに失敗しました。ログを確認してください。");
  }
}