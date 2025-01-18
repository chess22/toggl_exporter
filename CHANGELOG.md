# CHANGELOG

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