@echo off

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
echo.
echo Builds a VB6 exe project
echo.
echo Usage:
echo.
echo makeExe projectName [path] [/CONSOLE] [/NOV6CC]
echo.
echo   projectName      project name (excluding version)
echo   path             path to folder containing project file
echo   /CONSOLE         link the exe to the Console library
echo   /NOV6CC          don't use Version 6 Common Controls
echo.
exit /B

:doIt

set PROJECTNAME=%~1

set FOLDER=%~2
if not "%FOLDER:~0,1%"=="." (
	set FOLDER=%~2\
	pushd %FOLDER%
	shift
) else (
	set FOLDER=
)

if /I "%~2" == "/CONSOLE" (
	set LINKTOCONSOLE=yes
) else if /I "%~2" == "/NOV6CC" (
	set NOV6CC=NOV6CC
) else if not "%~2"=="" (
	echo Invalid parameter %~2
	goto :err
)

if /I "%~3" == "/CONSOLE" (
	set LINKTOCONSOLE=yes
) else if /I "%~3" == "/NOV6CC" (
	set NOV6CC=NOV6CC
) else if not "%~3"=="" (
	echo Invalid parameter %~3
	goto :err
)

set FILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.exe

echo =================================
echo Building %FILENAME%

echo Setting version = %VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%
setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-REVISION% -mode:N
if errorlevel 1 goto :err

vb6.exe /m %PROJECTNAME%.vbp
if errorlevel 1 goto :err

if defined LINKTOCONSOLE (
	echo Linking CONSOLE
	link /EDIT /SUBSYSTEM:CONSOLE %BIN-PATH%\%FILENAME%
	if errorlevel 1 goto :err
)

if defined NOV6CC (
	call generateAssemblyManifest.bat %PROJECTNAME% .exe /NOV6CC
) else (
	call generateAssemblyManifest.bat %PROJECTNAME% .exe
)
if errorlevel 1 goto :err

if defined FOLDER popd %FOLDER%
exit /B 0

:err
pause
if defined FOLDER popd %FOLDER%
exit /B 1


