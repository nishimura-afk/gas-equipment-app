# 0. Claspのログイン状態チェック（改善版）
Write-Host "🔍 Claspのログイン状態を確認しています..." -ForegroundColor Cyan

# ログイン状態を確認（エラーを抑制）
$loginCheck = clasp login --status 2>&1
$loginStatus = $loginCheck | Out-String

# "Logged in" が含まれているか、またはエラーが認証関連でない場合はログイン済みとみなす
$isLoggedIn = ($loginStatus -match "Logged in") -or ($LASTEXITCODE -eq 0)

if (-not $isLoggedIn) {
    Write-Host "⚠️ Claspにログインの必要があります..." -ForegroundColor Yellow
    Write-Host "   ブラウザが開きますので、Googleアカウントでログインしてください。" -ForegroundColor Yellow
    
    # ログインを実行
    clasp login
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ ログインに失敗しました。" -ForegroundColor Red
        exit 1
    }
    
    # 認証ファイルの保存などを少し待つ
    Write-Host "⏳ 認証情報を保存中..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Write-Host "✅ ログイン完了" -ForegroundColor Green
} else {
    Write-Host "✅ ログイン済みを確認しました" -ForegroundColor Green
}

# 1. GASにプッシュ（認証エラー時は再ログインを試みる）
Write-Host ""
Write-Host "🚀 GASにコードをプッシュしています..." -ForegroundColor Cyan
$pushOutput = clasp push 2>&1 | Out-String
$pushSuccess = $LASTEXITCODE -eq 0

# 認証エラーの場合、再ログインを試みる
if (-not $pushSuccess -and $pushOutput -match "invalid_grant|invalid_rapt|reauth") {
    Write-Host "⚠️  認証エラーが検出されました。再ログインを試みます..." -ForegroundColor Yellow
    Write-Host "   ブラウザが開きますので、Googleアカウントでログインしてください。" -ForegroundColor Yellow
    
    clasp login
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ 再ログインに失敗しました。" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "⏳ 認証情報を保存中..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    Write-Host "🔄 再度プッシュを試みます..." -ForegroundColor Cyan
    clasp push
    $pushSuccess = $LASTEXITCODE -eq 0
}

# clasp pushが失敗した場合、処理を中止
if (-not $pushSuccess) {
    Write-Host "❌ GASへのプッシュに失敗しました。処理を中止します。" -ForegroundColor Red
    Write-Host "   エラー内容: $pushOutput" -ForegroundColor Red
    exit 1
}

Write-Host "✅ プッシュ完了" -ForegroundColor Green

# 2. GASをデプロイ
Write-Host ""
Write-Host "📦 GASをデプロイしています..." -ForegroundColor Cyan

# 既存のデプロイメントを確認
$deployments = clasp deployments 2>&1 | Out-String

# 既存のデプロイメントIDを取得（最初のデプロイメントを使用）
if ($deployments -match "- ([a-zA-Z0-9_-]+) @\d+") {
    $deploymentId = $matches[1]
    Write-Host "   既存のデプロイメントを更新します (ID: $deploymentId)" -ForegroundColor Gray
    clasp deploy -i $deploymentId
} else {
    Write-Host "   新規デプロイメントを作成します" -ForegroundColor Gray
    clasp deploy
}

# デプロイが失敗した場合でも続行（警告のみ）
if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️  デプロイに失敗しましたが、プッシュは完了しています。" -ForegroundColor Yellow
    Write-Host "   手動でデプロイする場合は: clasp deploy" -ForegroundColor Yellow
} else {
    Write-Host "✅ デプロイ完了" -ForegroundColor Green
}

# 3. GitHubに保存
Write-Host ""
Write-Host "💾 GitHubに保存しています..." -ForegroundColor Cyan
git add .
git commit -m "自動更新"

if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️  コミットする変更がないか、コミットに失敗しました" -ForegroundColor Yellow
} else {
    git push
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ GitHub への push に失敗しました" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "✅ GitHubに保存完了" -ForegroundColor Green
}

Write-Host ""
Write-Host "✅ すべて完了しました！" -ForegroundColor Green