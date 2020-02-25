@echo off

setlocal

set curdir=%~dp0
rem ログファイル
set log=logs\00_Release_collect.log

rem ログフォルダ作成
mkdir "%curdir%\logs" 2>&1

echo  > "%curdir%\%log%"

rem vbaリリースバッチ
call "%curdir%\10_Release_vba.bat"
if %errorlevel% neq 0 (
    echo "vbaリリースバッチに失敗" >> "%curdir%\%log%" 2>&1
    echo vbaリリースバッチに失敗
    pause
    exit /b %errorlevel%
)

rem cppリリースバッチ
call "%curdir%\20_Release_cpp.bat"
if %errorlevel% neq 0 (
    echo "cppリリースバッチに失敗" >> "%curdir%\%log%" 2>&1
    echo cppリリースバッチに失敗
    pause
    exit /b %errorlevel%
)

echo Sut.xlamファイルを開いて条件付きコンパイルのDEBUG_MODEを消去してください。
pause

exit /b 0
