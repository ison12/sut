@echo off

rem Args check
if "%~1" == "" (
	exit /b 1
)

if "%~2" == "" (
	exit /b 2
)

if "%~3" == "" (
	exit /b 3
)

rem Variable declare
set curPath=%~dp0

set applicationDirPath=%~1
set downloadDirPath=%~2
set downloadApplicationDirPath=%~3

set nowDate=%date:~0,4%%date:~5,2%%date:~8,2%
set nowTime=%time:~0,2%%time:~3,2%%time:~6,2%

rem Copy module files
copy /b /v /y "%downloadApplicationDirPath%\*.xlam" "%applicationDirPath%"
copy /b /v /y "%downloadApplicationDirPath%\*.exe" "%applicationDirPath%"
copy /b /v /y "%downloadApplicationDirPath%\*.exe.config" "%applicationDirPath%"
copy /b /v /y "%downloadApplicationDirPath%\*.prop" "%applicationDirPath%"

rem Copy Update.bat file
copy /b /v /y "%downloadApplicationDirPath%\SutUpdate.bat" "%applicationDirPath%"

rem Copy dll files (If not exists dll then no check)
copy /b /v /y "%downloadApplicationDirPath%\*.dll" "%applicationDirPath%"

rem Copy manual
rem rd /s /q "%applicationDirPath%\Manual"
rem move /y "%downloadApplicationDirPath%\Manual" "%applicationDirPath%\Manual"
rem Backup convert scripts
rem robocopy /s /e "%applicationDirPath%\ConvertScripts" "%applicationDirPath%\ConvertScripts_bk_%nowDate%%nowTime%"
rem Copy convert scripts (Newer or Changed)
rem robocopy /s /e /xo "%downloadApplicationDirPath%\ConvertScripts" "%applicationDirPath%\ConvertScripts"

rem Success
exit /b 0

@echo on