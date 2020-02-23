@echo off

setlocal

set curdir=%~dp0

rem CPPソースフォルダ
set cpp=%curdir%\..\..\source\cpp\
rem 出力先フォルダ
set des=%curdir%\..\Sut\
rem ログファイル
set log=logs\20_Release_cpp.log

rem ログファイルを初期化
echo  > "%curdir%\%log%"

rem SutInstallerなどをコピーする
copy /b /v /y "%cpp%\Sut\Release\*" "%des%" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "Sutのコピーに失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

exit /b 0
