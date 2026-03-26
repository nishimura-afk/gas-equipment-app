## 2026-03-26

- **やったこと**: ベーパー回収率(D70S/L100R)管理機能を設計・実装。efaxプロジェクトにZOHO→Gemini→スプレッドシートのパイプライン追加。GAS Webアプリに回収率管理画面+ダッシュボードアラート追加。既存データ5595件を見積管理DBの新「回収率」シートに移行。
- **決めたこと**: efaxプロジェクトに機能追加（別プロジェクトにしない）。PDFはGoogle Drive（04_設備管理/回収率PDF/）に月別保存。正常範囲0.05%〜0.2%。

### 残タスク
- [x] `.env` に `ZOHO_VR_FOLDERS` を設定（26店舗分完了）
- [x] `npx tsx src/setup-gmail-auth.ts` でOAuth再認証（drive.fileスコープ追加完了）
- [x] `clasp push` でGASコードを反映（2回実施、手動入力機能含む）
- [x] サンプルPDFでGemini抽出の精度検証（岐阜東・かつらぎ正常抽出確認）
- [x] 手動入力機能追加（月計表が届かない店舗用）
- [ ] Webアプリの動作確認（回収率タブ）
- [ ] 実運用テスト（来月の月計表到着時に自動取得を確認）

### 関連ファイル
- efax追加: `src/vapor-recovery-fetcher.ts`, `src/vapor-recovery-extractor.ts`, `src/vapor-recovery-writer.ts`
- GAS追加: `vapor-recovery.html`, `Code.js`(末尾), `dashboard.html`, `index.html`, `0_Config.js`
