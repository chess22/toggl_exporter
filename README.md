# TogglからGoogleカレンダーへのエクスポートツール

このプロジェクトは、Masato Kawaguchi氏による[Toggl Exporter](https://github.com/mkawaguchi/toggl_exporter)をフォークしたもので、TogglとGoogleカレンダーの同期機能を強化するために、追加の機能や改善を行っています。

---

## 新機能と改善点
- 重複イベントの自動削除と作成防止機能。
- テスト機能（例: `testRecordActivityLog`の改良）。
- `clasp`を利用した自動化対応。
- エラーハンドリングと通知機能の強化。
- デプロイや設定手順の簡素化。
- **初回は過去30日分取得し、2回目以降は「前回停止時刻 - 1日」から差分取得**できるようにし、事後修正や削除を反映。

---

## 同期ロジックの概要

1. **初回実行**  
   - キャッシュが無い場合（初回）は、**過去30日分**のTogglデータを取得してGoogleカレンダーに登録します。  
   - 実行後、最後に取得したタイムエントリの停止時刻をキャッシュし、次回以降の差分取得に活かします。

2. **2回目以降の実行**  
   - 前回保存した「最終停止時刻」から **1日分のオーバーラップ**を取ります。  
   - 具体的には `(前回最終停止時刻 - 1日) 〜 現在` の区間でToggl APIを呼び出し、過去1日以内に修正・削除されたデータも検知します。

3. **重複イベントの防止**  
   - カレンダーに登録する際、タイトル末尾に `ID:xxxx`（TogglのエントリID）を付与。  
   - 既存イベントがあれば更新、なければ新規作成し、重複を防ぎます。

4. **削除・修正の検知**  
   - `deleteRemovedEntries` や `removeDuplicateEvents` 関数を活用し、Togglから削除されたエントリなどをカレンダーから除外可能。

---

## 設定手順

### 1. リポジトリのクローン
``````bash
git clone https://github.com/<your-username>/toggl_exporter.git
cd toggl_exporter
``````

### 2. claspのインストール・ログイン
``````bash
npm install -g @google/clasp
clasp login
``````

### 3. Google Apps Scriptとの連携
既存のGASプロジェクトをクローンする、または新規作成します。
``````bash
clasp clone <YOUR_SCRIPT_ID>
``````

### 4. スクリプトプロパティ
- `TOGGL_BASIC_AUTH`: Toggl APIのBase64トークン
- `GOOGLE_CALENDAR_ID`: 同期先GoogleカレンダーID
- `NOTIFICATION_EMAIL`: エラー通知用メール

### 5. デプロイとトリガー設定
``````bash
clasp push
``````
- 時間トリガーをセットして自動実行をオンにします。

---

## バージョン情報

**現在のバージョン:** 1.4.00

### 変更履歴

#### [1.4.00] - 2025-01-18
- **初回実行時に過去30日分を取得**  
- **2回目以降は (前回停止時刻 - 1日) から現在まで**を差分取得
- 重複防止・更新ロジックを維持し、過去1日以内の変更・削除を反映

---

## ライセンス

このプロジェクトはBSD-3-Clauseライセンス下で提供されています。  
- **元の作者**: Masato Kawaguchi  
- **変更者**: chess  

詳細は[LICENSE](https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE)をご参照ください。

---

## 主な機能

1. **初回30日 + 2回目以降1日オーバーラップ**でTogglデータをGoogleカレンダーへ同期。
2. **重複イベント防止**: タイトル末尾に `ID:xxxx` 形式で登録し、重複や更新を管理。
3. **メール通知付きのエラーハンドリング**。
4. **テスト用関数**: `testRecordActivityLog`, `testCreateDuplicateEvents` など。