@echo off

setlocal

set curdir=%~dp0
rem ���O�t�@�C��
set log=logs\01_Release_archive.log

rem ���O�t�H���_�쐬
mkdir "%curdir%\logs"

echo  > "%curdir%\%log%"

rem ZIP�t�@�C����\�ߍ폜����
del "%curdir%\..\Sut.zip" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "ZIP�t�@�C���̍폜�Ɏ��s" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem ZIP���k
cd "%curdir%\..\"
"%curdir%\zip\zip.exe" -r "Sut.zip" "Sut" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "ZIP���k�Ɏ��s" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

exit /b 0
