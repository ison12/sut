@echo off

setlocal

set curdir=%~dp0

rem VBAソースフォルダ
set vba=%curdir%\..\..\source\vba\
rem 出力先フォルダ
set des=%curdir%\..\Sut\
rem ログファイル
set log=logs\20_Release_cpp.log

rem ログファイルを初期化
echo  > "%curdir%\%log%"

exit /b 0
