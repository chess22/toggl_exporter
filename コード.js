/*
  Toggl time entries export to GoogleCalendar
  original author: Masato Kawaguchi
  modified by: chess
  Released under the BSD-3-Clause license
  version: 1.4.02-modified
  https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE

  CHANGELOG:
    v1.4.02-modified (2025/02/01):
      - タイムアウト対策として、進捗保存および途中再開の仕組みを追加
      - 手動実行モードを3種類実装：タイムアウトモード（1分区切り）、完遂モード（自動再開、5分30秒）、初回実行モード（キャッシュ無視）
      - 初期取得時、IDが無いタイムエントリは、開始時刻と名称（説明）から複合キーを生成し、重複チェック時に使用
      - 既存のカレンダーイベントにIDが付与されていない場合、重複チェック時にタイトルを更新して複合キーを付与するように変更
      - 進捗状況のログは、処理開始時、タイムアウト時、および処理完了時にのみ出力
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
  
  MAX_CACHE_AGE: 12 * 60 * 60, // 12時間（秒）
  RETRY_COUNT: 5,
  RETRY_DELAY: 2000, // ms
  
  // タイムアウト閾値（ミリ秒）
  AUTOMATIC_TIMEOUT_INTERVAL: 240000,         // 自動実行：4分
  MANUAL_COMPLETE_TIMEOUT_INTERVAL: 330000,     // 手動完遂モード：5分30秒
  MANUAL_TIMEOUT_MODE_INTERVAL: 60000           // 手動タイムアウトモード：1分
};

/** ログレベル **/
const LOG_LEVELS = { DEBUG: 1, INFO: 2, ERROR: 3 };
const CURRENT_LOG_LEVEL = CONFIG.DEBUG_MODE ? LOG_LEVELS.DEBUG : LOG_LEVELS.INFO;

/** ログ出力用関数 **/
function log(level, message) {
  if (level >= CURRENT_LOG_LEVEL) {
    Logger.log(message);
  }
}

/** キャッシュから最後の更新日時（UNIXタイムスタンプ）を取得。キャッシュが無ければ -1 **/
function getLastModifyDatetime() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get(CONFIG.CACHE_KEY);
  if (!cachedData) {
    log(LOG_LEVELS.DEBUG, "キャッシュが見つかりません");
    return -1;
  }
  try {
    var parsedData = JSON.parse(cachedData);
    log(LOG_LEVELS.DEBUG, "キャッシュヒット: " + JSON.stringify(parsedData));
    return parsedData.timestamp;
  } catch (e) {
    log(LOG_LEVELS.ERROR, "キャッシュ解析エラー: " + e.message);
    cache.remove(CONFIG.CACHE_KEY);
    return -1;
  }
}

/** 最終更新日時をキャッシュに保存 **/
function putLastModifyDatetime(timestamp) {
  var cache = CacheService.getScriptCache();
  var data = JSON.stringify({ timestamp: timestamp });
  cache.put(CONFIG.CACHE_KEY, data, CONFIG.MAX_CACHE_AGE);
  log(LOG_LEVELS.DEBUG, "キャッシュ更新: " + data);
}

/** retry ユーティリティ関数 **/
function retry(fn, retries, delay) {
  for (var i = 0; i < retries; i++) {
    try { 
      return fn(); 
    } catch (e) {
      if (i < retries - 1) {
        log(LOG_LEVELS.DEBUG, "リトライ中 (" + (i + 1) + "/" + retries + ") - エラー: " + e.message);
        Utilities.sleep(delay);
      } else { 
        throw e; 
      }
    }
  }
}

/** Toggl API から指定期間のタイムエントリを取得 **/
function getTimeEntriesRange(startIso, endIso) {
  return retry(function(){
    var uri = CONFIG.TOGGL_API_HOSTNAME + "/api/v9/me/time_entries?start_date=" 
              + encodeURIComponent(startIso) + "&end_date=" + encodeURIComponent(endIso);
    log(LOG_LEVELS.DEBUG, "Fetching time entries from: " + uri);
    var response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    log(LOG_LEVELS.DEBUG, "API Response Code: " + responseCode);
    log(LOG_LEVELS.DEBUG, "API Response: " + responseText);
    if (responseCode !== 200) {
      log(LOG_LEVELS.ERROR, "API Error: " + responseText);
      throw new Error("Toggl API returned status code " + responseCode);
    }
    var parsed = JSON.parse(responseText);
    if (Array.isArray(parsed)) { return parsed; }
    else { log(LOG_LEVELS.ERROR, "API returned non-array response"); return null; }
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/** Google カレンダーにイベントを作成 **/
function recordActivityLog(title, started_at, ended_at) {
  log(LOG_LEVELS.DEBUG, 'recordActivityLog called with title: "' + title + '", started_at: "' + started_at + '", ended_at: "' + ended_at + '"');
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  if (!calendar) { 
    log(LOG_LEVELS.ERROR, 'Invalid GOOGLE_CALENDAR_ID: "' + CONFIG.GOOGLE_CALENDAR_ID + '"'); 
    throw new Error('Invalid GOOGLE_CALENDAR_ID: "' + CONFIG.GOOGLE_CALENDAR_ID + '"'); 
  }
  var startDate = new Date(started_at), endDate = new Date(ended_at);
  if (isNaN(startDate.getTime())) { 
    log(LOG_LEVELS.ERROR, 'Invalid start date: "' + started_at + '"'); 
    throw new RangeError('Invalid start date: "' + started_at + '"'); 
  }
  if (isNaN(endDate.getTime())) { 
    log(LOG_LEVELS.ERROR, 'Invalid end date: "' + ended_at + '"'); 
    throw new RangeError('Invalid end date: "' + ended_at + '"'); 
  }
  try {
    calendar.createEvent(title, startDate, endDate);
    log(LOG_LEVELS.INFO, 'Created event: "' + title + '" from ' + started_at + " to " + ended_at);
  } catch (createError) { 
    log(LOG_LEVELS.ERROR, "Error creating event: " + createError);
    throw createError; 
  }
}

/** 複合キーを生成する関数（IDが無い場合用） **/
function getCompositeId(newRecord) {
  // 開始時刻と説明を組み合わせ、一意のキーを生成
  return "NO_ID:" + new Date(newRecord.start).toISOString() + "_" + (newRecord.description || "名称なし");
}

/** 既存イベントの検索と更新（重複防止）
    ・IDがある場合はそのIDを使用、無い場合は複合キーを生成し利用
    ・さらに、名称と開始・終了時刻が一致するかで重複を判定
    ・既存イベントにIDが付与されていない場合は、タイトルを更新して複合キーを付与
*/
function eventExistsAndUpdate(record_id, newRecord) {
  var idKey = record_id ? record_id : getCompositeId(newRecord);
  var calendar = CalendarApp.getCalendarById(CONFIG.GOOGLE_CALENDAR_ID);
  var now = new Date();
  var pastDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
  // まず、"ID:" + idKey を含むイベントを検索
  var events = calendar.getEvents(pastDate, now, { search: "ID:" + idKey });
  log(LOG_LEVELS.DEBUG, "Searching for events with key " + idKey + ". Found " + events.length + " events.");
  
  if (events.length > 0) {
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var title = event.getTitle();
      // 既存イベントのタイトルに "ID:" が含まれているかチェック
      if (title.indexOf("ID:") === -1) {
        // ID情報が付与されていない場合は、名称と時刻から新たなタイトルを生成し更新
        var project_data = getProjectData(newRecord.wid, newRecord.pid);
        var project_name = project_data.name || '';
        var newName = [(newRecord.description || '名称なし'), project_name].filter(Boolean).join(" : ");
        var updatedTitle = newName + " ID:" + idKey;
        try {
          retry(function(){
            event.setTitle(updatedTitle);
          }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
          log(LOG_LEVELS.INFO, "Updated existing event title to include composite key: " + updatedTitle);
        } catch(e) {
          log(LOG_LEVELS.ERROR, "Error updating event title with composite key: " + e);
        }
        return true;
      } else {
        // タイトルに「ID:」がある場合、"ID:"以前の部分を名称として抽出
        var titleParts = title.split(" ID:");
        var eventName = titleParts[0];
        var project_data = getProjectData(newRecord.wid, newRecord.pid);
        var project_name = project_data.name || '';
        var newName = [(newRecord.description || '名称なし'), project_name].filter(Boolean).join(" : ");
        var eventStart = event.getStartTime().toISOString();
        var eventEnd = event.getEndTime().toISOString();
        if (eventName === newName &&
            eventStart === newRecord.start &&
            eventEnd === newRecord.stop) {
          log(LOG_LEVELS.INFO, "Matching event found with identical name and time for key " + idKey);
          return true;
        }
      }
    }
  }
  return false;
}

/** Toggl のプロジェクトデータ取得（IDベース） **/
function getProjectData(workspace_id, project_id) {
  if (!workspace_id || !project_id) return {};
  return retry(function(){
    var uri = CONFIG.TOGGL_API_HOSTNAME + "/api/v9/workspaces/" + workspace_id + "/projects/" + project_id;
    var response = UrlFetchApp.fetch(uri, {
      method: 'GET',
      headers: { "Authorization": "Basic " + CONFIG.TOGGL_BASIC_AUTH },
      muteHttpExceptions: true
    });
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    log(LOG_LEVELS.DEBUG, "Project API Response Code: " + responseCode);
    log(LOG_LEVELS.DEBUG, "Project API Response: " + responseText);
    if (responseCode !== 200) {
      log(LOG_LEVELS.ERROR, "Project API Error: " + responseText);
      return {};
    }
    return JSON.parse(responseText);
  }, CONFIG.RETRY_COUNT, CONFIG.RETRY_DELAY);
}

/**
 * ----------------------------
 * タイムアウト対策＆実行モード分岐（進捗保存機構付き）
 * ----------------------------
 */

// 進捗管理用のキー
const PROGRESS_KEY = 'toggl_exporter:last_processed_index';

/**
 * バッチ処理のコア関数
 * @param {boolean} isManual   手動実行なら true、定期実行なら false
 * @param {boolean} autoResume 手動実行時に自動再開する場合は true、タイムアウトで中断する場合は false
 * @param {boolean} forceInitial true の場合、キャッシュを無視して初回実行（過去30日分取得）とする
 */
function processTimeEntriesBatch(isManual, autoResume, forceInitial) {
  forceInitial = forceInitial || false;
  
  var MAX_EXECUTION_TIME;
  if (isManual) {
    if (autoResume) {
      MAX_EXECUTION_TIME = CONFIG.MANUAL_COMPLETE_TIMEOUT_INTERVAL;  // 手動完遂モード：5分30秒
    } else {
      MAX_EXECUTION_TIME = CONFIG.MANUAL_TIMEOUT_MODE_INTERVAL;       // 手動タイムアウトモード：1分
    }
  } else {
    MAX_EXECUTION_TIME = CONFIG.AUTOMATIC_TIMEOUT_INTERVAL;             // 自動実行：4分
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
        var idKey = record.id ? record.id : getCompositeId(record);
        var activity_log = [(record.description || '名称なし'), project_name].filter(Boolean).join(" : ") + " ID:" + idKey;
        recordActivityLog(activity_log, record.start, record.stop);
        log(LOG_LEVELS.INFO, "Added event: " + activity_log);
      } else {
        log(LOG_LEVELS.DEBUG, "Existing event processed for ID (or key): " + (record.id ? record.id : getCompositeId(record)));
      }
    } catch (e) {
      log(LOG_LEVELS.ERROR, "Error processing record ID:" + (record.id || "NoID") + " - " + e);
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
      log(LOG_LEVELS.INFO, "Timeout reached: Processed " + processedCount + " of " + totalCount + " (" + percentComplete + "%). Remaining: " + remainingCount + " records. Current record's stop date: " + record.stop);
      
      if (!isManual || (isManual && autoResume)) {
        log(LOG_LEVELS.INFO, (isManual ? "手動完遂" : "自動実行") + ": 閾値に達したため中断します。Next start index: " + (i + 1));
        ScriptApp.newTrigger('automaticProcessTimeEntries')
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
}

/**
 * 自動実行用エントリポイント（トリガー経由）
 */
function automaticProcessTimeEntries() {
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
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('手動同期（タイムアウトモード）', 'manualProcessTimeEntriesTimeout')
    .addItem('手動同期（完遂モード）', 'manualProcessTimeEntriesComplete')
    .addItem('手動同期（初回実行モード）', 'manualProcessTimeEntriesInitial')
    .addToUi();
}