@echo off

set curpath=%~dp0

cscript //nologo "%curpath%\vbac.wsf" decombine /vbaproj /binary:"%curpath%\..\src\" /source:"%curpath%\..\src_export"

IF NOT %ERRORLEVEL% == 0 (
	exit /b %ERRORLEVEL%
)

exit /b 0
