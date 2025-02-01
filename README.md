# TogglからGoogleカレンダーへのエクスポートツール  
# Toggl to Google Calendar Export Tool

このプロジェクトは、Masato Kawaguchi氏による [Toggl Exporter](https://github.com/mkawaguchi/toggl_exporter) をベースに、Toggl と Google カレンダーの同期ロジックを強化・拡張した Google Apps Script です。  
This project is a Google Apps Script based on Masato Kawaguchi's [Toggl Exporter](https://github.com/mkawaguchi/toggl_exporter) that enhances and extends the synchronization logic between Toggl and Google Calendar.

---

## 主な機能 / Main Features

1. **Toggl → Googleカレンダー 同期 / Toggl to Google Calendar Synchronization**  
   - **初回実行時 / Initial Run:**  
     過去30日分のタイムエントリを一括取得し、Google カレンダーに登録します。  
     Fetches 30 days of time entries and registers them in Google Calendar.
   - **2回目以降 / Subsequent Runs:**  
     前回停止時刻から1日前以降の差分を取得し、事後修正・削除に対応します。  
     Retrieves incremental changes from (previous stop time - 1 day) to the present.

2. **重複イベントの作成防止 / Duplicate Event Prevention**  
   - イベントタイトルの末尾に `ID:xxxx` を付与し、既存のイベントがあれば更新、無ければ新規作成します。  
     Appends `ID:xxxx` to event titles; if an event with the same ID exists, it is updated instead of creating a duplicate.
   - `removeDuplicateEvents()` により、重複したイベントを削除し、最新の1件のみを残します。  
     The `removeDuplicateEvents()` function deletes duplicate events, leaving only the latest one.

3. **削除チェック / Deletion Check**  
   - **短期間削除チェック (`deleteRemovedEntriesShort()`) / Short-range Delete Check:**  
     過去1日間のカレンダーイベントをチェックし、Togglから削除されたエントリに対応。  
     Checks calendar events for the past 1 day and deletes those whose corresponding Toggl entries have been removed.
   - **長期間削除チェック (`deleteRemovedEntriesManual()`) / Manual Delete Check:**  
     過去1ヶ月を対象に、手動実行で古い削除に対応します。  
     Manually checks the past 1 month to remove outdated events.

4. **通知機能・テスト機能 / Notification & Test Features**  
   - エラー発生時に指定メールアドレスへ通知を送信。  
     Sends error notification emails to a configured address.
   - `clearScriptCache()`, `sendTestEmail()`, `testCreateDuplicateEvents()`, `testRemoveDuplicateEvents()`, `testDuplicateEventsWorkflow()` などのテスト用関数により、各機能の動作確認が可能。  
     Test functions (e.g. `clearScriptCache()`, `sendTestEmail()`, etc.) are provided to verify the functionality.

---

## バージョン情報 / Version Info

- **最新バージョン / Latest Version:** 1.4.03 (2025-02-01)

### 変更点 / Changelog Highlights
- 初回実行時の分割取得・進捗保存機能を追加  
  (Added batch processing with progress saving when timeout occurs.)
- 手動実行モードを3種類実装：  
  - タイムアウトモード：1分区切りで中断し、ユーザー再実行で続行  
  - 完遂モード：タイムアウト閾値を5分30秒に設定し、自動再開する  
  - 初回実行モード：常に初回実行（保存された進捗を無視）  
- 全実行モードでロック (LockService) を利用し、排他制御を実現  
  (All modes acquire a lock to prevent concurrent execution.)
- 進捗状況のログ出力を、処理開始時、タイムアウト時、完了時に限定  
- 削除チェック機能およびテスト機能を復元  
  (Restored deletion checks and test functions.)

---

## セットアップ手順 / Setup

1. **リポジトリのクローン or ダウンロード / Clone or Download Repository**  
   ```bash
   git clone https://github.com/<your-username>/toggl_exporter.git
   cd toggl_exporter
```

2. **clasp / Google Apps Script の準備 / Setup clasp / Google Apps Script**

```
npm install -g @google/clasp
clasp login
clasp clone <YOUR_SCRIPT_ID>
```

  

3. **スクリプトプロパティの設定 / Set Script Properties**

• TOGGL_BASIC_AUTH: Toggl API トークン (Base64認証情報)

• GOOGLE_CALENDAR_ID: 同期先の Google カレンダー ID

• NOTIFICATION_EMAIL: エラー通知用メールアドレス

4. **push & デプロイ / Push & Deploy**

```
clasp push
```

• 定期実行のため、Google Apps Script の時間ベースのトリガーで watch() を設定してください。

• 削除チェック用の deleteRemovedEntriesShort() なども、必要に応じてトリガー設定可能です。

  

5. **動作確認 / Test**

• スプレッドシートのカスタムメニューから「テストメールを送信」で通知機能をチェック。

• testCreateDuplicateEvents() などで、イベント作成や重複チェックの動作を確認してください。

**使い方 / Usage**

• **定期実行 (watch)**

• **初回実行:** 過去30日分のエントリを取得

• **継続実行:** 前回停止時刻から1日前以降の差分を取得

• **削除チェック**

• deleteRemovedEntriesShort(): 過去1日間のみをチェック（低負荷）

• deleteRemovedEntriesManual(): 過去1ヶ月を対象に手動実行し、古い削除を拾う

• **重複イベント削除**

• removeDuplicateEvents(): 同一IDを持つ古いイベントを削除

• **手動実行モード**

• タイムアウトモード、完遂モード、初回実行モードから選択可能。

• カスタムメニューから実行できます。

**ライセンス / License**

  

This project is released under the BSD-3-Clause License.

• **Original Author:** Masato Kawaguchi

• **Modified by:** chess

See the [LICENSE](https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE) file for details.
