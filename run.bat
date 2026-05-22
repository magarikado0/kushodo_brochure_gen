@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo 京大書道部パンフレット生成ツール
echo.

if not exist "input\作品情報フォーム.xlsx" (
  echo [ERROR] input\作品情報フォーム.xlsx が見つかりません。
  echo input フォルダに作品情報フォーム.xlsxを置いてください。
  echo.
  pause
  exit /b 1
)

if not exist "templates\冬樟展パンフ.docx" (
  echo [ERROR] templates\冬樟展パンフ.docx が見つかりません。
  echo templates フォルダに前回パンフレットのdocxを置いてください。
  echo.
  pause
  exit /b 1
)

where uv >nul 2>nul
if errorlevel 1 (
  echo uv が見つからないため、インストールします。
  echo 初回のみ少し時間がかかります。
  powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
  if errorlevel 1 (
    echo.
    echo [ERROR] uv のインストールに失敗しました。
    pause
    exit /b 1
  )
  set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%"
)

echo.
echo パンフレットを生成します。
uv run python scripts\generate.py

if errorlevel 1 (
  echo.
  echo [ERROR] 生成に失敗しました。
  pause
  exit /b 1
)

echo.
echo 完了しました。output フォルダを確認してください。
pause
