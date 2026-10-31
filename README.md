# LINE Loan Management Bot

LINEグループ内の友人同士の貸し借りを記録し，Googleスプレッドシートで一元管理するGoogle Apps Script製のBotです．

「貸したはずなのに返されていない」「返したはずなのに相手が把握していない」といった認識の齟齬を減らすことを目的に作成しました．LINE上の簡単なコマンドで貸付・返済を登録でき，Webダッシュボードから未返済状況や履歴を確認できます．

## 主な機能

- LINEから貸付を登録
- LINEから返済済みに更新
- LINE User IDを利用したメンバー管理
- 未返済一覧の表示
- 60日経過後の初回通知と，その後7日ごとの再通知
- Webダッシュボードで未返済総額・返済率・メンバー別状況を表示
- Webダッシュボードから返済済みに更新
- 複数ユーザーによる更新を考慮したScript Lock

## 使用技術

- JavaScript / Google Apps Script
- LINE Messaging API
- Google Sheets
- HTML / CSS / JavaScript

## 構成

```text
line-loan-management-bot/
├── Code.gs             # LINE Webhook，貸付・返済処理，自動通知
├── Dashboard.gs        # Webダッシュボードのサーバ側処理
├── Index.html          # WebダッシュボードUI
├── appsscript.json     # Apps Script設定
└── docs/
    └── architecture.md
```

## セットアップ

### 1. Googleスプレッドシートを作成

貸し借り情報を保存するGoogleスプレッドシートを作成し，URL中のSpreadsheet IDを控えます．

### 2. Apps Scriptプロジェクトを作成

Google Apps Scriptに以下のファイルを作成し，本リポジトリの内容を配置します．

- `Code.gs`
- `Dashboard.gs`
- `Index.html`
- `appsscript.json`

### 3. Script Propertiesを設定

Apps Scriptの「プロジェクトの設定」→「スクリプト プロパティ」に，環境ごとの値を設定します．

| Key | Required | Description |
|---|---:|---|
| `SPREADSHEET_ID` | Yes | 自分のGoogleスプレッドシートID |
| `CHANNEL_ACCESS_TOKEN` | Yes | LINE Messaging APIのChannel Access Token |
| `DASHBOARD_PIN` | No | Webダッシュボード用PIN |
| `DASHBOARD_TALK_ID` | No | 複数グループ利用時の対象Talk ID |
| `MIGRATION_ADMIN_USER_ID` | No | 既存データ移行を許可するLINE User ID |
| `OVERDUE_FIRST_DAYS` | No | 初回督促までの日数．未設定時60日 |
| `OVERDUE_REPEAT_DAYS` | No | 再通知間隔．未設定時7日 |
| `OVERDUE_TRIGGER_HOUR` | No | 自動確認を実行する時刻．未設定時9時台 |

**実際のIDやアクセストークンをGitHubへcommitしないでください．**

### 4. 初期化

Apps Script上で一度だけ`setupSheet()`を実行します．自動督促を利用する場合は，通知対象を`previewOverdueNotifications()`で確認した後，`setupOverdueNotificationTrigger()`を実行します．

### 5. Webアプリとしてデプロイ

Apps ScriptをWebアプリとしてデプロイし，発行されたURLをLINE DevelopersのWebhook URLとして設定します．同じURLをブラウザで開くとダッシュボードを利用できます．

## コマンド例

```text
貸した 田中 3000 昼ごはん
返した 田中 3000
名前登録 田中
未返済
名前一覧
```

## 工夫した点

- 名前の表記揺れだけに依存せず，LINE User IDを基準に人物を管理
- 返済状況をLINEとWebの両方から更新可能にし，同じGoogle Sheetsへ反映
- 長期未返済を定期的に検出し，通知履歴を保持して過剰通知を防止
- Spreadsheet IDやChannel Access Tokenなどをソースコードから分離し，再利用しやすい構成に変更

## 注意

本リポジトリには実際の利用者名，LINE User ID，Talk ID，貸付データ，Spreadsheet ID，API認証情報を含めないでください．
