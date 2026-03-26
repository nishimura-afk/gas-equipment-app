# 設備管理アプリ

ガソリンスタンドの全般的な設備（洗車機、計量機、コンプレッサー等）を管理するWebアプリ。

## 何ができるか

- 設備の一覧表示と状態管理
- 交換時期のアラート（閾値ベース）
- ベンダー情報の管理（メールアドレス含む）
- カレンダー連携（交換予定の登録）
- Google Drive 連携（関連資料の管理）
- メール下書きの自動作成
- プロジェクト進捗の管理

## 技術

- Google Apps Script（clasp で管理）
- スプレッドシートをデータベースとして使用
- Calendar API・Drive API・Gmail 連携
- Webアプリとしてデプロイ

## ファイル構成

- `0_Config.js` … 設定（ベンダー情報、アラート閾値、機密情報の外部化）
- `1_Initialize.gs.js` … 初期化処理
- `2_DataService.js` … データ読み書き
- `3_DataUpdateService.js` … データ更新処理
- `4_CalendarService.js` … カレンダー連携
- `5_DriveService.js` … Google Drive 連携
- `Code.js` … Webアプリのメイン処理
- `index.html` … トップページ
- `dashboard.html` … ダッシュボード画面
- `equipment-list.html` … 設備一覧画面

## 開発方法

```
clasp push    # GASにコードを反映
clasp pull    # GASからコードを取得
```
