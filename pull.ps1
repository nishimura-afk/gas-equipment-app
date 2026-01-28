# pull.ps1 - GAS側の変更をローカル・GitHubに反映します
Write-Host "📥 GASから最新コードを取得しています..." -ForegroundColor Cyan

# GASからローカルにダウンロード
clasp pull

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ clasp pull に失敗しました" -ForegroundColor Red
    exit 1
}

Write-Host "✅ ローカルに反映完了" -ForegroundColor Green
Write-Host ""
Write-Host "📦 GitHubに保存しています..." -ForegroundColor Cyan

# Gitにコミット
git add .
git commit -m "GAS側の変更を反映"

if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️  コミットする変更がないか、コミットに失敗しました" -ForegroundColor Yellow
} else {
    # GitHubにプッシュ
    git push origin main
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ GitHub への push に失敗しました" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "✅ GitHubに保存完了" -ForegroundColor Green
}

Write-Host ""
Write-Host "✅ すべて完了しました！" -ForegroundColor Green
