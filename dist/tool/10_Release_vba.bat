@echo off

setlocal

set curdir=%~dp0

rem VBA�\�[�X�t�H���_
set vba=%curdir%\..\..\source\vba\
rem �o�͐�t�H���_
set des=%curdir%\..\Sut\
rem ���O�t�@�C��
set log=logs\10_Release_vba.log

rem ���O�t�@�C����������
echo  > "%curdir%\%log%"

rem Sut���R�s�[����
copy /b /v /y "%vba%\src\Sut.xlam" "%des%" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "Sut�̃R�s�[�Ɏ��s" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

rem ���\�[�X�t�H���_�����S�ɓ�������
robocopy /mir "%vba%\src\resource" "%des%\resource" /xd config >> "%curdir%\%log%" 2>&1
if %errorlevel% geq 8 (
    echo "���\�[�X�t�H���_�̓����Ɏ��s" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

exit /b 0
