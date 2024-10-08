setlocal

if "%~1"=="" (
	echo First parameter must be module name
	goto :err
)
if /I "%~2"=="DLL" (
	echo DLL >nul
) else if /I "%~2"=="OCX" (
	echo OCX >nul
) else (
	echo Second parameter must be 'dll' or 'ocx'
	goto :err
)
if /I "%~3"=="NOVERSION" (
	set OBJECTFILE=%~1.%~2
) else (
	set OBJECTFILE=%~1%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.%~2
)

echo Unregistering "%OBJECTFILE%"
regsvr32 -s -u "%OBJECTFILE%"
if errorlevel 1 goto :err

exit /B

:err
exit /B 1

