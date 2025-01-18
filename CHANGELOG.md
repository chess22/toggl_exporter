# CHANGELOG

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