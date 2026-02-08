# CHANGELOG

## [1.4.05] - 2026-02-08

### Fixed / 修正
- **Basic認証の二重エンコードを修正**: スクリプトプロパティにBase64エンコード済みの値が格納されているため、`Utilities.base64Encode()` を除去して401エラーを解消
  (Removed `Utilities.base64Encode()` to fix double-encoding of already Base64-encoded TOGGL_BASIC_AUTH, resolving 401 errors)
- **空データ時のlastModify更新を抑止**: 取得件数0件のときに `timestamp=0` が保存される問題を修正
  (Avoid updating lastModify when there are zero entries to prevent saving `timestamp=0`)
- **再開位置のフォールバックを追加**: 前回レコードIDが見つからない場合、`lastModify` で再開位置を推定
  (Added resume fallback using `lastModify` when the last processed record ID is missing)
- **プロジェクト情報取得をメモ化**: 同一プロジェクトへのAPI呼び出しを削減
  (Memoized project lookups to reduce duplicate API calls)

---

## [1.4.04] - 2026-02-08

### Fixed / 修正
- **`notifyError` 関数の追加**: 5箇所で呼ばれていたが未定義だったため、エラー時にReferenceErrorが発生していた
  (Added missing `notifyError` function that was called in 5 places but never defined)
- **ワンタイムトリガー蓄積防止**: タイムアウト時に既存トリガーを削除してから再作成するよう修正
  (Prevent one-time trigger accumulation by cleaning up old triggers before creating new ones)
- **delete関数のretry範囲修正**: 関数全体をretryで囲んでいたため再試行時に既削除イベントでエラーになる問題を修正
  (Fixed retry scope in delete functions to prevent errors on already-deleted events)
- **Toggl API v9 非推奨フィールド対応**: `record.pid`/`record.wid` を `record.project_id`/`record.workspace_id` に変更
  (Replaced deprecated `pid`/`wid` fields with `project_id`/`workspace_id` for Toggl API v9)
- **Basic認証のBase64エンコード復元**: `Utilities.base64Encode()` が欠落しており認証が失敗する可能性があった
  (Restored `Utilities.base64Encode()` for proper Basic authentication)
- **進捗管理をレコードIDベースに変更**: 配列インデックスだとデータ変動時にずれるため、record.idで再開位置を特定
  (Changed progress tracking from array index to record ID for reliable resume after timeout)
- **タイムアウト中断時にもlastModifyキャッシュを保存**: 次回の取得開始位置を適切に更新
  (Save lastModify cache on timeout to correctly update the next fetch start position)
- **`removeDuplicateEvents` で最新イベントを残すように修正**: コメントと挙動の不一致を解消
  (Fixed `removeDuplicateEvents` to keep the latest event, matching the JSDoc description)
- **`eventExistsAndUpdate` の時間比較をミリ秒ベースに変更**: `toISOString()` の文字列比較ではタイムゾーン表記差（Z vs +09:00）で毎回不要な更新が発生していた
  (Changed time comparison from `toISOString()` string matching to `getTime()` millisecond comparison)
- **タイムアウト再開を `watchResume` 関数に分離**: 定期実行の `watch` トリガーが巻き添え削除される問題を防止
  (Separated timeout resume into `watchResume` function to protect user's periodic `watch` trigger)
- **DEBUGログのAPIレスポンス出力を先頭1000文字に制限**: 大量データ取得時のログ肥大化を防止
  (Limited API response debug logging to first 1000 characters)

### Changed / 変更
- **`var` を `const`/`let` に統一**: V8ランタイムのブロックスコープを活用
  (Unified variable declarations to `const`/`let` for V8 runtime block scoping)
- **READMEの `TOGGL_BASIC_AUTH` 説明を修正**: 「Base64認証情報」→「平文 `{API_TOKEN}:api_token` 形式」に統一
  (Updated README to clarify TOGGL_BASIC_AUTH should be stored in plaintext format)

---

## [1.4.03] - 2025-02-01

### Added / 追加
- 初回実行時の分割取得・進捗保存機能を追加
  (Added batch processing with progress saving for initial run)
- 手動実行モードを3種類実装：タイムアウトモード、完遂モード、初回実行モード
  (Implemented 3 manual execution modes: timeout, complete, and initial)
- 全実行モードでロック(LockService)による排他制御
  (All modes acquire a lock to prevent concurrent execution)
- 削除チェック・テスト用関数・キャッシュクリア機能を復元
  (Restored deletion checks, test functions, and cache clear functionality)

---

## [1.4.02] - 2025-01-18

### Changed / 変更
- **Overlap period** for subsequent Toggl fetch shortened to **1 day**  
  (従来の 3 日から 1 日に短縮)
- **Short-range delete check** (`deleteRemovedEntriesShort`) now covers only **past 1 day**  
  (短期間削除チェックは過去 1 日のみ対象)
- Manual delete (`deleteRemovedEntriesManual`) remains **past 1 month**  
  (手動実行の削除チェックは従来通り過去 1 ヶ月を維持)

### Fixed / 修正
- Minor code adjustments for consistent logging  
  (ログ出力の一貫性向上のための微調整)

---

## [1.4.01] - 2025-01-18

### Changed / 変更
- **Overlap period** was set to **3 days** (previously 1 day in test)  
  (オーバーラップ期間を 1 日→ 3 日に拡張)
- **Short-range delete** (`deleteRemovedEntriesShort`) also updated to past **3 days**  
  (短期間削除チェックを過去 3 日に)
- **Manual delete** for past **1 month**  
  (古い削除に対応するため手動で過去 1 ヶ月検索)
- Improved comments and clarification regarding deletion logic  
  (削除処理に関するコメントや説明を強化)

---
## [1.4.00] - 2025-01-18
### Added / 追加
- **Initial 30-day fetch**:
  - When no cache is found (first run), fetch 30 days of Toggl entries.
  - キャッシュが無い場合に30日分のTogglエントリを取得。

### Changed / 変更
- **1-day overlap**:
  - On subsequent runs, fetch from (previous stop time - 1 day) to now.
  - 2回目以降は「前回停止時刻 - 1日」から現在までを再取得し、過去1日以内の変更・削除を拾えるように。

### Fixed / 修正
- Ensured existing duplication prevention and event update logic remain compatible with the new fetch strategy.
- 既存の重複防止ロジックやイベント更新フローとの整合性を確認・維持。
