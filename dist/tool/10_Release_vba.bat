@echo off

setlocal

set curdir=%~dp0

rem VBAソースフォルダ
set vba=%curdir%\..\..\source\vba\
rem 出力先フォルダ
set des=%curdir%\..\Sut\
rem ログファイル
set log=logs\10_Release_vba.log

rem ログファイルを初期化
echo  > "%curdir%\%log%"

rem Sutをコピーする
copy /b /v /y "%vba%\src\Sut.xlam" "%des%" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "Sutのコピーに失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem リソースフォルダを完全に同期する
robocopy /mir "%vba%\src\resource" "%des%\resource" /xd config >> "%curdir%\%log%" 2>&1
if %errorlevel% geq 8 (
    echo "リソースフォルダの同期に失敗" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

exit /b 0
