# 1. GASに送る
Write-Host "🚀 GASに送っています..."
clasp push

# 2. GitHubに送る
Write-Host "📦 GitHubに保存しています..."
git add .
git commit -m "自動更新"
git push

Write-Host "✅ すべて完了しました！"