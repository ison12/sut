@echo off

setlocal

set curdir=%~dp0

rem VBA�\�[�X�t�H���_
set vba=%curdir%\..\..\source\vba\
rem �o�͐�t�H���_
set des=%curdir%\..\Sut\
rem ���O�t�@�C��
set log=logs\20_Release_cpp.log

rem ���O�t�@�C����������
echo  > "%curdir%\%log%"

exit /b 0
