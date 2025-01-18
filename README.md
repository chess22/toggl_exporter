# TogglからGoogleカレンダーへのエクスポートツール

このプロジェクトは、Masato Kawaguchi氏による [Toggl Exporter](https://github.com/mkawaguchi/toggl_exporter) をベースに、TogglとGoogleカレンダーの同期ロジックを強化・拡張したGoogle Apps Scriptです。

---

## 主な機能 / Main Features

1. **Toggl → Googleカレンダー 同期**  
   - **初回実行時**: 過去30日分のタイムエントリを一括取得し、Googleカレンダーに登録します。  
   - **2回目以降**: (前回停止時刻 - 1日)から現在までの差分を取得し、事後修正・削除にある程度対応。

2. **重複イベントの作成を防止**  
   - イベントタイトルの末尾に `ID:xxxx` を付与。IDが一致する既存イベントがあれば更新、なければ新規作成。  
   - `removeDuplicateEvents` で重複イベントを削除し、最新の1つを残します。

3. **Togglで削除されたエントリをカレンダーからも削除**  
   - `deleteRemovedEntriesShort()`: 過去1日間のカレンダーイベントをチェック。Togglから削除されたIDをカレンダー側でも削除（低負荷）。  
   - `deleteRemovedEntriesManual()`: 過去1ヶ月を対象に手動で古い削除に対応。

4. **通知機能・テスト機能**  
   - エラーが起きると指定メールアドレスに通知。  
   - `testRecordActivityLog` や `testCreateDuplicateEvents` などで動作確認が可能。

---

## バージョン情報 / Version Info

- **最新バージョン**: 1.4.02 (2025-01-18)

### 変更点
- **オーバーラップ期間**を (前回停止時刻 - 1日) に短縮  
- **削除チェック（短期間）**を 過去1日 に変更  
- **手動削除**は 過去1ヶ月 を維持  
- 負荷軽減を図りつつ、必要に応じて古い削除にも対応可能

---

## セットアップ手順 / Setup

1. **リポジトリのクローン or ダウンロード**  
   ```bash
   git clone https://github.com/<your-username>/toggl_exporter.git
   cd toggl_exporter
   ```
2. **clasp / Google Apps Script の準備**  
   ```bash
   npm install -g @google/clasp
   clasp login
   clasp clone <YOUR_SCRIPT_ID>
   ```
3. **スクリプトプロパティ** を設定  
   - `TOGGL_BASIC_AUTH`: Toggl APIトークン (Base64認証情報)  
   - `GOOGLE_CALENDAR_ID`: 同期先GoogleカレンダーID  
   - `NOTIFICATION_EMAIL`: エラー通知用メールアドレス
4. **push & デプロイ**  
   ```bash
   clasp push
   ```
   - 時間ベースのトリガー等で `watch()` を定期実行。  
   - `deleteRemovedEntriesShort()` などを別トリガーで動かすことも可能。
5. **動作確認**  
   - カスタムメニューから「テストメールを送信」で通知機能をチェック。  
   - `testRecordActivityLog` などでイベント作成テスト。

---

## 使い方 / Usage

- **定期実行 (watch)**  
  - **初回**: 過去30日を取得。  
  - **2回目以降**: (前回停止時刻 - 1日) 〜 現在 の範囲を差分取得。
- **削除チェック**  
  - `deleteRemovedEntriesShort()`: 過去1日間だけをチェック（低負荷）。  
  - `deleteRemovedEntriesManual()`: 過去1ヶ月を手動実行し、古い削除を拾う。
- **重複イベント削除**  
  - `removeDuplicateEvents()`: 同一IDを持つ古いイベントを削除。

---

## ライセンス / License

- BSD-3-Clause License  
- Original Author: Masato Kawaguchi  
- Modified by: chess  

詳細は [LICENSE](https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE) をご参照ください。

