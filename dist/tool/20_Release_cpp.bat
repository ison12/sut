@echo off

setlocal

set curdir=%~dp0

rem CPP�\�[�X�t�H���_
set cpp=%curdir%\..\..\source\cpp\
rem �o�͐�t�H���_
set des=%curdir%\..\Sut\
rem ���O�t�@�C��
set log=logs\20_Release_cpp.log

rem ���O�t�@�C����������
echo  > "%curdir%\%log%"

rem SutInstaller�Ȃǂ��R�s�[����
copy /b /v /y "%cpp%\Sut\Release\*" "%des%" >> "%curdir%\%log%" 2>&1
if %errorlevel% neq 0 (
    echo "Sut�̃R�s�[�Ɏ��s" >> "%curdir%\%log%"
    exit /b %errorlevel%
)

exit /b 0
