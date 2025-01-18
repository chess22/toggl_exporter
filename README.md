# TogglからGoogleカレンダーへのエクスポートツール

このプロジェクトは、Masato Kawaguchi氏による[Toggl Exporter](https://github.com/mkawaguchi/toggl_exporter)をフォークしたもので、TogglとGoogleカレンダーの同期機能を強化するために、追加の機能や改善を行っています。

---

## 新機能と改善点
- 重複イベントの自動削除と作成防止機能を追加。
- テスト機能（例: `testRecordActivityLog`の改良）。
- `clasp`を利用した自動化対応。
- エラーハンドリングと通知機能の強化。
- デプロイや設定手順の簡素化。
- `moment.js`の依存を排除し、ネイティブな日時操作を採用。

---

## 設定手順

### 1. リポジトリのクローン

以下のコマンドでリポジトリをクローンし、作業ディレクトリを移動します。

``````
bash
git clone https://github.com/chess22/toggl_exporter.git
cd toggl_exporter
``````

---

### 2. 必要なツールをインストール

#### A. clasp のインストール
以下のコマンドで`clasp`をグローバルにインストールします。

``````
bash
npm install -g @google/clasp
``````

#### B. clasp へのログイン
Googleアカウントを使用してclaspにログインします。

``````
bash
clasp login
``````

---

### 3. Google Apps Scriptプロジェクトとのリンク

既存のGoogle Apps Scriptプロジェクトをクローンします。プロジェクトIDは事前に取得しておいてください。

``````
bash
clasp clone <YOUR_SCRIPT_ID>
``````

---

### 4. プロジェクトの設定

#### A. スクリプトプロパティの設定
以下のプロパティをGoogle Apps Scriptのプロパティサービスに設定します。

1. **`TOGGL_BASIC_AUTH`**: Toggl APIのBasic認証トークン（Base64エンコード済み）。
2. **`GOOGLE_CALENDAR_ID`**: GoogleカレンダーのID。
3. **`NOTIFICATION_EMAIL`**: エラー通知を受信するメールアドレス。

**設定手順**:
1. Google Apps Scriptのエディタを開きます。
2. 「ファイル」 > 「プロジェクトのプロパティ」 > 「スクリプトのプロパティ」タブを選択します。
3. 上記のプロパティをキーと値で追加します。

#### B. スプレッドシートを利用した設定（オプション）
スプレッドシートを用いて動的な設定を管理する場合、以下の手順を参考にしてください。

1. スプレッドシートを作成。
2. シート名を`Settings`に変更。
3. 1列目にキー（例: `TOGGL_BASIC_AUTH`）、2列目に値を入力。

以下のコードをスクリプトに追加し、スプレッドシートから設定を取得するよう変更できます。

``````
javascript
function getSetting(key) {
  const spreadsheet = SpreadsheetApp.openById('<SPREADSHEET_ID>');
  const sheet = spreadsheet.getSheetByName('Settings');
  if (!sheet) throw new Error('Settingsシートが見つかりません');
  const data = sheet.getDataRange().getValues();
  const setting = data.find(row => row[0] === key);
  if (!setting) throw new Error(`設定キーが見つかりません: ${key}`);
  return setting[1];
}
``````

---

### 5. スクリプトのデプロイ

最新の変更をGoogle Apps Scriptプロジェクトにプッシュします。

``````
bash
clasp push
``````

その後、Google Apps Scriptエディタで時間ベースのトリガーを設定し、自動実行を有効にします。

---

## バージョン情報

**現在のバージョン**: 1.4.00

### 変更履歴
#### [1.4.00] - 2025-01-17
- `testCreateDuplicateEvents`関数を追加し、重複イベント処理を自動化。
- テストイベントの日時を未来の時刻に設定するよう、`testRecordActivityLog`を改良。
- エラーハンドリングとデバッグ機能を強化。
- `moment.js`依存を削除し、標準のJavaScript機能を使用。
- スプレッドシートを利用した動的な設定管理のサポートを追加。

---

## ライセンス

このプロジェクトはBSD-3-Clauseライセンスの下で提供されています。

- **元の作者**: Masato Kawaguchi  
- **変更者**: chess  

詳細は[LICENSE](https://github.com/mkawaguchi/toggl_exporter/blob/master/LICENSE)をご参照ください。

---

## 主な機能
1. 重複するGoogleカレンダーイベントを自動的に削除。
2. デバッグ機能の強化により、問題の特定を容易化。
3. メール通知を用いた強力なエラーハンドリング。
4. テスト機能を簡素化するための補助関数を提供。