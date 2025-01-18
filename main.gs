/*
  Toggl time entries export to GoogleCalendar
  author: Masato Kawaguchi
  Released under the BSD-3-Clause license
  version: 1.1.01
  https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE

  required: moment.js
    project-key: 15hgNOjKHUG4UtyZl9clqBbl23sDvWMS8pfDJOyIapZk5RBqwL3i-rlCo

  Copyright 2024, Masato Kawaguchi

  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

  1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
  2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
  3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS” AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
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

// **設定セクション**
const CONFIG = {
  CACHE_KEY: 'toggl_exporter:lastmodify_datetime', // キャッシュキー
  TIME_OFFSET: 9 * 60 * 60, // JST (秒)
  TOGGL_API_HOSTNAME: 'https://api.track.toggl.com',
  GOOGLE_CALENDAR_ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_CALENDAR_ID'), // GoogleカレンダーID
  NOTIFICATION_EMAIL: PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL'), // エラー通知メールアドレス
};

const TOGGL_BASIC_AUTH = PropertiesService.getScriptProperties().getProperty('TOGGL_BASIC_AUTH'); // Toggl API認証情報（Base64エンコード済み）
const DEBUG_MODE = false; // デバッグモード（true: ログ出力, false: ログ出力なし）

/**
 * ログ出力用関数
 * @param {string} message - ログメッセージ
 */
function logDebug(message) {
  if (DEBUG_MODE) {
    Logger.log(message);
  }
}

/**
 * エラー発生時に通知メールを送信する関数
 * @param {Error} e - 発生したエラー
 */
function notifyError(e) {
  const email = CONFIG.NOTIFICATION_EMAIL;
  const subject = 'Google Apps Script エラー通知';
  const body = `
以下のエラーが発生しました:

エラーメッセージ:
${e.toString()}

スタックトレース:
${e.stack}
  `;
  
  if (email) {
    MailApp.sendEmail(email, subject, body);
    logDebug(`エラー通知メールを送信しました: ${email}`);
  } else {
    logDebug("通知先メールアドレスが設定されていません。");
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
    Logger.log(`テストメールを送信しました: ${email}`);
  } else {
    Logger.log("通知先メールアドレスが設定されていません。");
  }
}

/**
 * キャッシュをクリアする関数
 */
function clearScriptCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
  Logger.log("Cache cleared");
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
    .addToUi();
}

/**
 * 指定された日数の開始時刻（00:00:00）を返す関数
 * @param {number} days - 現在からの日数（正の数で未来、負の数で過去）
 * @returns {string} ISO8601形式の日時文字列（UTC）
 */
function beginningOfDay(days) {
  var now = new Date();
  if (isNaN(now.getTime())) {
    Logger.log("Invalid Date value!");
    return;
  }
  
  now.setUTCHours(0, 0, 0, 0);
  now.setUTCDate(now.getUTCDate() + days);
  
  return Utilities.formatDate(now, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

var MAX_CACHE_AGE = 12 * 60 * 60; // 12時間（秒単位）

/**
 * キャッシュから最後の更新日時を取得する関数
 * @returns {number} UNIXタイムスタンプ（秒）または-1（キャッシュが存在しない/エラー）
 */
function getLastModifyDatetime() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get(CONFIG.CACHE_KEY);
  
  if (!cachedData) {
    logDebug("No cached data found");
    return -1;
  }
  
  logDebug("Cached data: " + cachedData);
  
  try {
    var parsedData = JSON.parse(cachedData);
    if (parsedData && typeof parsedData.timestamp === 'number') {
      logDebug("Parsed cached data: " + parsedData.timestamp);
      return parsedData.timestamp;
    } else {
      logDebug("Invalid cached data format");
      cache.remove(CONFIG.CACHE_KEY);
      return -1;
    }
  } catch (e) {
    logDebug("Error parsing cached data: " + e.message);
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
  cache.put(CONFIG.CACHE_KEY, data, 6 * 60 * 60); // 6時間の有効期限
  logDebug("Cache updated with timestamp: " + data);
}

/**
 * Toggl APIからタイムエントリを取得する関数
 * @param {number} start_timestamp - 開始時刻のUNIXタイムスタンプ（秒）
 * @param {number} days - 取得する日数
 * @returns {Array|null} タイムエントリの配列またはnull（エラー時）
 */
function getTimeEntries(start_timestamp, days) {
  var startOfDay = new Date(start_timestamp * 1000);
  startOfDay.setUTCHours(0, 0, 0, 0);
  var endOfDay = new Date(startOfDay);
  endOfDay.setUTCDate(startOfDay.getUTCDate() + days);
  
  var startOfDayStr = Utilities.formatDate(startOfDay, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var endOfDayStr = Utilities.formatDate(endOfDay, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  
  var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries?start_date=${encodeURIComponent(startOfDayStr)}&end_date=${encodeURIComponent(endOfDayStr)}`;
  
  logDebug(`Fetching time entries from: ${uri}`);
  
  var response;
  try {
    response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
  } catch (fetchError) {
    logDebug(`Fetch error: ${fetchError}`);
    throw fetchError;
  }
  
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  logDebug(`API Response Code: ${responseCode}`);
  logDebug(`API Response: ${responseText}`);
  
  if (responseCode !== 200) {
    logDebug(`API Error: ${responseText}`);
    throw new Error(`Toggl API returned status code ${responseCode}`);
  }
  
  try {
    var parsed = JSON.parse(responseText);
    if (Array.isArray(parsed)) {
      return parsed;
    } else {
      logDebug("API returned non-array response");
      return null;
    }
  } catch (e) {
    logDebug(`JSON parse error: ${e}`);
    throw e;
  }
}

/**
 * Toggl APIからプロジェクトデータを取得する関数
 * @param {string} workspace_id - ワークスペースID
 * @param {string} project_id - プロジェクトID
 * @returns {Object} プロジェクトデータまたは空オブジェクト
 */
function getProjectData(workspace_id, project_id) {
  if (!workspace_id || !project_id) return {};
  
  var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/workspaces/${workspace_id}/projects/${project_id}`;
  
  var response;
  try {
    response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
  } catch (fetchError) {
    logDebug(`Fetch error for project data: ${fetchError}`);
    return {};
  }
  
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  logDebug(`Project API Response Code: ${responseCode}`);
  logDebug(`Project API Response: ${responseText}`);
  
  if (responseCode !== 200) {
    logDebug(`Project API Error: ${responseText}`);
    return {};
  }
  
  try {
    return JSON.parse(responseText);
  } catch (e) {
    logDebug(`Project JSON parse error: ${e}`);
    return {};
  }
}

/**
 * Googleカレンダーにイベントを記録する関数
 * @param {string} title - イベントのタイトル（IDを含む）
 * @param {string} started_at - イベント開始時刻のISO8601文字列
 * @param {string} ended_at - イベント終了時刻のISO8601文字列
 */
function recordActivityLog(title, started_at, ended_at) {
  logDebug(`recordActivityLog called with title: "${title}", started_at: "${started_at}", ended_at: "${ended_at}"`);
  
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  
  // カレンダーが取得できているか確認
  if (!calendar) {
    logDebug(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
    throw new Error(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
  }
  
  // 日時文字列をDateオブジェクトに変換
  var startDate = new Date(started_at);
  var endDate = new Date(ended_at);
  
  // 日時の有効性をチェック
  if (isNaN(startDate.getTime())) {
    logDebug(`Invalid start date: "${started_at}"`);
    throw new RangeError(`Invalid start date: "${started_at}"`);
  }
  
  if (isNaN(endDate.getTime())) {
    logDebug(`Invalid end date: "${ended_at}"`);
    throw new RangeError(`Invalid end date: "${ended_at}"`);
  }
  
  // タイムゾーンを設定（必要に応じて）
  calendar.setTimeZone('Asia/Tokyo');
  
  try {
    // イベントを作成
    calendar.createEvent(title, startDate, endDate);
    logDebug(`Created event: "${title}" from ${started_at} to ${ended_at}`);
  } catch (createError) {
    logDebug(`Error creating event: ${createError}`);
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
  
  logDebug(`Searching for events with ID:${record_id}. Found ${events.length} events.`);
  
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
      logDebug(`Matching event found: "${matchingEvent.getTitle()}"`);
      var eventTitle = matchingEvent.getTitle();
      var eventStart = matchingEvent.getStartTime().toISOString();
      var eventEnd = matchingEvent.getEndTime().toISOString();
      var newStart = newRecord.start;
      var newEnd = newRecord.stop;
      
      // タイムエントリの名称を基に新しいタイトルを生成
      var project_data = getProjectData(newRecord.wid, newRecord.pid);
      var project_name = project_data.name || '';
      var updatedTitle = [(newRecord.description || '名称なし'), project_name].filter(Boolean).join(" : ") + ` ID:${record_id}`;
      
      // イベントタイトルと時間の比較
      var eventTitleNeedsUpdate = matchingEvent.getTitle() !== updatedTitle;
      var eventTimeNeedsUpdate = (matchingEvent.getStartTime().toISOString() !== newStart) || (matchingEvent.getEndTime().toISOString() !== newEnd);
      
      logDebug(`EventTitleNeedsUpdate: ${eventTitleNeedsUpdate}, EventTimeNeedsUpdate: ${eventTimeNeedsUpdate}`);
      
      if (eventTitleNeedsUpdate || eventTimeNeedsUpdate) {
        matchingEvent.setTitle(updatedTitle);
        matchingEvent.setTime(new Date(newStart), new Date(newEnd));
        logDebug(`Updated event for ID:${record_id}`);
        return true;
      }
      
      logDebug(`No update needed for event ID:${record_id}`);
      return true; // イベントは存在し、必要な更新も行った
    }
  }
  
  return false; // イベントは存在しない
}

/**
 * Togglで削除されたタイムエントリに対応するGoogleカレンダーのイベントを削除する関数
 */
function deleteRemovedEntries() {
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  if (!calendar) {
    logDebug(`Invalid GOOGLE_CALENDAR_ID: "${CONFIG.GOOGLE_CALENDAR_ID}"`);
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
        logDebug(`Deleted event for removed Toggl entry ID:${record_id}`);
      }
    }
  });
}

/**
 * Toggl APIで指定したIDのタイムエントリが存在するか確認する関数
 * @param {string} record_id - タイムエントリのID
 * @returns {boolean} 存在する場合はtrue、存在しない場合はfalse
 */
function checkIfTogglEntryExists(record_id) {
  var uri = `${CONFIG.TOGGL_API_HOSTNAME}/api/v9/me/time_entries/${record_id}`;
  
  var response;
  try {
    response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
  } catch (fetchError) {
    logDebug(`Fetch error for checking entry existence: ${fetchError}`);
    return false;
  }
  
  var responseCode = response.getResponseCode();
  
  if (responseCode === 200) {
    return true;
  } else if (responseCode === 404) {
    return false;
  } else {
    logDebug(`Unexpected API response code when checking entry existence: ${responseCode}`);
    return false;
  }
}

/**
 * Googleカレンダー内の重複イベントを削除する関数
 * 既に存在するイベントの中から最新の1つを残し、他を削除します。
 */
function removeDuplicateEvents() {
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  var now = new Date();
  var pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate()); // 過去3ヶ月間
  
  var events = calendar.getEvents(pastDate, now);
  
  var eventMap = {};
  
  events.forEach(function(event) {
    var title = event.getTitle();
    var match = title.match(/ID:(\d+)$/); // タイトルの末尾にIDがあることを確認
    if (match && match[1]) {
      var record_id = match[1];
      if (eventMap[record_id]) {
        // 既に存在する場合、最新のイベント以外を削除
        event.deleteEvent();
        logDebug(`Deleted duplicate event for ID:${record_id}`);
      } else {
        eventMap[record_id] = event;
      }
    }
  });
}

/**
 * ロックを取得する関数
 * @returns {Lock} 取得したロックオブジェクト
 */
function getLock() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 最大30秒待つ
    logDebug("Lock acquired");
    return lock;
  } catch (e) {
    logDebug("Could not acquire lock: " + e);
    throw e;
  }
}

/**
 * メイン関数: Togglからデータを取得し、Googleカレンダーにイベントを記録・更新・削除
 */
function watch() {
  var lock;
  try {
    // ロックを取得して他の実行を防止
    lock = getLock();
    
    var check_datetime = getLastModifyDatetime();
    
    if (check_datetime == -1) {
      var initial_date = new Date();
      initial_date.setUTCHours(0,0,0,0);
      var initial_timestamp = Math.floor(initial_date.getTime() / 1000);
      logDebug("No cached data found, using initial timestamp: " + initial_timestamp);
      check_datetime = initial_timestamp;
    } else {
      logDebug("Data is being fetched starting from: " + new Date(check_datetime * 1000).toISOString());
    }
    
    var days_to_fetch = 3;
    var time_entries = getTimeEntries(check_datetime, days_to_fetch);
    
    if (time_entries) {
      logDebug("Number of time entries fetched: " + time_entries.length);
      var last_stop_datetime = check_datetime;
      
      time_entries.forEach(function(record) {
        if (!record.stop) {
          logDebug("Record with no stop time: " + JSON.stringify(record));
          return;
        }
        
        var stop_time = Math.floor(new Date(record.stop).getTime() / 1000);
        if (isNaN(stop_time)) {
          logDebug("Invalid stop time: " + record.stop);
          return;
        }
        
        var start_time = Math.floor(new Date(record.start).getTime() / 1000);
        if (isNaN(start_time)) {
          logDebug("Invalid start time: " + record.start);
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
            logDebug("Added event: " + activity_log);
          } catch (e) {
            logDebug(`Error adding event for ID: ${record.id} - ${e}`);
            // 必要に応じて追加のエラーハンドリングを実装
          }
        } else {
          logDebug("Existing event processed for ID: " + record.id);
        }
        
        if (stop_time > last_stop_datetime) {
          last_stop_datetime = stop_time;
        }
      });
      
      logDebug("Updating cache with last stop datetime: " + last_stop_datetime);
      
      if (last_stop_datetime) {
        var new_timestamp = last_stop_datetime + 1;
        putLastModifyDatetime(new_timestamp);
        logDebug("Cache updated after fetching data with timestamp: " + new_timestamp);
      } else {
        logDebug("No valid stop time found, cache not updated.");
      }
    } else {
      logDebug("No time entries found");
    }
  } catch (e) {
    Logger.log(e);
    notifyError(e);
  } finally {
    if (lock) {
      lock.releaseLock();
      logDebug("Lock released");
    }
  }
}
