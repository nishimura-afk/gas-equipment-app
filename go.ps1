# 0. Claspのログイン状態チェック
Write-Host "🔍 Claspのログイン状態を確認しています..."
# エラー出力を含めて取得するように 2>&1 を使用
$loginStatus = clasp login --status 2>&1 | Out-String

# "Logged in" が含まれていない場合はログインを試みる
if ($loginStatus -notmatch "Logged in") {
    Write-Host "⚠️ Claspにログインの必要があるようです..." -ForegroundColor Yellow
    
    # ログインを実行（すでにログイン済みの警告が出ても無視して進むようにする）
    clasp login
    
    # 認証ファイルの保存などを少し待つ
    Write-Host "⏳ 処理を待機しています..."
    Start-Sleep -Seconds 3
} else {
    Write-Host "✅ ログイン済みを確認しました。" -ForegroundColor Green
}

# 1. GASに送る
Write-Host "🚀 GASに送っています..."
clasp push

# clasp pushが失敗した場合（エラーコードが0以外）、ここで処理を止める安全策
# ここで本当にログインできていなければエラーになり、止まります
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ GASへのプッシュに失敗しました。Gitへの保存を中止します。" -ForegroundColor Red
    exit
}

# 2. GitHubに送る
Write-Host "📦 GitHubに保存しています..."
git add .
git commit -m "自動更新"
git push

Write-Host "✅ すべて完了しました！" -ForegroundColor Cyan