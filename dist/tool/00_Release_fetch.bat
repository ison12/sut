@echo off

setlocal

set curdir=%~dp0
rem ログファイル
set log=logs\00_Release_fetch.log

rem ログフォルダ作成
mkdir "%curdir%\logs"

echo  > "%curdir%\%log%"

rem vbaリリースバッチ
call "%curdir%\10_Release_vba.bat"
if %errorlevel% neq 0 (
    echo "vbaリリースバッチに失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem cppリリースバッチ
call "%curdir%\20_Release_cpp.bat"
if %errorlevel% neq 0 (
    echo "cppリリースバッチに失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem ZIP圧縮
cd "%curdir%\..\"
"%curdir%\zip\zip.exe" -r "Sut.zip" "Sut" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "ZIP圧縮に失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

echo Sut.xlamファイルを開いて条件付きコンパイルのDEBUG_MODEを消去してください。
pause

exit /b 0
