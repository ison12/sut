@echo off

setlocal

set curdir=%~dp0
rem ���O�t�@�C��
set log=logs\00_Release_collect.log

rem ���O�t�H���_�쐬
mkdir "%curdir%\logs" 2>&1

echo  > "%curdir%\%log%"

rem vba�����[�X�o�b�`
call "%curdir%\10_Release_vba.bat"
if %errorlevel% neq 0 (
    echo "vba�����[�X�o�b�`�Ɏ��s" >> "%curdir%\%log%" 2>&1
    echo vba�����[�X�o�b�`�Ɏ��s
    pause
    exit /b %errorlevel%
)

rem cpp�����[�X�o�b�`
call "%curdir%\20_Release_cpp.bat"
if %errorlevel% neq 0 (
    echo "cpp�����[�X�o�b�`�Ɏ��s" >> "%curdir%\%log%" 2>&1
    echo cpp�����[�X�o�b�`�Ɏ��s
    pause
    exit /b %errorlevel%
)

echo Sut.xlam�t�@�C�����J���ď����t���R���p�C����DEBUG_MODE���������Ă��������B
pause

exit /b 0
