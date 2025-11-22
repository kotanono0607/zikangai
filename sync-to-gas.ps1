# ========================================
# GAS同期PowerShellスクリプト
# Claude Code → Git → ローカル → GAS
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GAS同期処理を開始します" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# C:\GASディレクトリに移動
Set-Location C:\GAS
if (-not $?) {
    Write-Host "[エラー] C:\GAS ディレクトリが見つかりません" -ForegroundColor Red
    Read-Host "Enterキーを押して終了"
    exit 1
}

Write-Host "[1/3] 現在のディレクトリ: $(Get-Location)" -ForegroundColor Green
Write-Host ""

# Git pull - 最新を取得
Write-Host "[2/3] Gitから最新を取得中..." -ForegroundColor Yellow
git pull
if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] git pull に失敗しました" -ForegroundColor Red
    Read-Host "Enterキーを押して終了"
    exit 1
}
Write-Host ""

# clasp push - GASにプッシュ
Write-Host "[3/3] GASにプッシュ中..." -ForegroundColor Yellow
clasp push
if ($LASTEXITCODE -ne 0) {
    Write-Host "[エラー] clasp push に失敗しました" -ForegroundColor Red
    Read-Host "Enterキーを押して終了"
    exit 1
}
Write-Host ""

Write-Host "========================================" -ForegroundColor Green
Write-Host " 同期完了！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "GASエディタをリロードして確認してください。" -ForegroundColor Cyan
Write-Host ""

Read-Host "Enterキーを押して終了"
