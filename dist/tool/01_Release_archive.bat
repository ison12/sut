@echo off

setlocal

set curdir=%~dp0
rem ログファイル
set log=logs\01_Release_archive.log

rem ログフォルダ作成
mkdir "%curdir%\logs" 2>&1

echo  > "%curdir%\%log%"

rem ZIPファイルを予め削除する
del "%curdir%\..\Sut.zip" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "ZIPファイルの削除に失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem ZIP圧縮
cd "%curdir%\..\"
"%curdir%\zip\zip.exe" -r "Sut.zip" "Sut" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "ZIP圧縮に失敗" >> "%curdir%\%log%"
    echo ZIP圧縮に失敗
    pause
    exit /b %errorlevel%
)

echo リリースファイルの準備が完了しました。
pause

exit /b 0
