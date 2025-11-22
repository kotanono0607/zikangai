@echo off
REM ========================================
REM GAS同期バッチファイル
REM Claude Code → Git → ローカル → GAS
REM ========================================

echo.
echo ========================================
echo  GAS同期処理を開始します
echo ========================================
echo.

REM C:\GASディレクトリに移動
cd /d C:\GAS
if errorlevel 1 (
    echo [エラー] C:\GAS ディレクトリが見つかりません
    pause
    exit /b 1
)

echo [1/3] 現在のディレクトリ: %CD%
echo.

REM Git pull - 最新を取得
echo [2/3] Gitから最新を取得中...
git pull
if errorlevel 1 (
    echo [エラー] git pull に失敗しました
    pause
    exit /b 1
)
echo.

REM clasp push - GASにプッシュ
echo [3/3] GASにプッシュ中...
clasp push
if errorlevel 1 (
    echo [エラー] clasp push に失敗しました
    pause
    exit /b 1
)
echo.

echo ========================================
echo  同期完了！
echo ========================================
echo.
echo GASエディタをリロードして確認してください。
echo.

pause
