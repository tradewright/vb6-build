@echo off
setlocal
setlocal enabledelayedexpansion

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
::===0=========1=========2=========3=========4=========5=========6=========7=========8
echo.
echo Builds a VB6 exe project
echo.
echo Usage:
echo.
echo makeExe projectName [path] [/M:{N^|E^|F}] [/CONSOLE] [/NOV6CC]
echo.
echo   projectName      Project name (excluding version)
echo   path             Path to folder containing project file
echo   /M               Manifest requirement:
echo                        N =^> no manifest (default)
echo                        E =^> embed manifest in object file
echo                        F =^> freestanding manifest file
echo   /CONSOLE         Link the exe to the Console library
echo   /NOV6CC          Don't use Version 6 Common Controls
echo.
echo   If /CONSOLE is used, then on entry ^%%BIN-PATH^%% must be 
echo   set to the folder where the generated manifest will be stored.
exit /B
::===0=========1=========2=========3=========4=========5=========6=========7=========8

:doIt
set MANIFEST=NONE
set PROJECTNAME=
set FOLDER=
set LINKTOCONSOLE=
set NOV6CC=

:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if /I "%ARG%" == "/M:N" (
	set MANIFEST=NONE
) else if /I "%ARG%" == "/M:E" (
	set MANIFEST=EMBED
) else if /I "%ARG%" == "/M:F" (
	set MANIFEST=NOEMBED
) else if /I "%ARG:~0,3%" == "/M:" (
	set MANIFEST=
) else if /I "%ARG%" == "/CONSOLE" (
	set LINKTOCONSOLE=yes
) else if /I "%ARG%" == "/NOV6CC" (
	set NOV6CC=NOV6CC
) else if "%ARG:~0,1%"=="/" (
	echo Invalid parameter '%ARG%'
	set ERROR=1
) else if not defined PROJECTNAME (
	set PROJECTNAME=%ARG%
) else if not defined FOLDER (
	set FOLDER=%ARG%
	pushd !FOLDER!>nul
	if errorlevel 1 (
		echo Invalid folder parameter '!FOLDER!'
		set ERROR=1
	)
) else (
	echo Invalid parameter '%ARG%'
	set ERROR=1
)

shift
goto :parse
	
:parsingComplete

if not defined PROJECTNAME (
	echo Projectname parameter must be supplied
	set ERROR=1
)
if defined LINKTOCONSOLE (
	if not defined BIN-PATH (
		echo ^%%BIN-PATH^%% is not defined
		set ERROR=1
	)
)
if not defined MANIFEST (
	echo /M:{N^|E^|F} setting missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

set FILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.exe

echo =================================
if defined FOLDER (
	echo Building %FOLDER%\%FILENAME%.exe
) else (
	echo Building %FILENAME%.exe
)

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

if "%MANIFEST%"=="NONE" (
	echo don't generate a manifest>nul
) else if "%MANIFEST%"=="EMBED" (
	if defined NOV6CC (
		call generateAssemblyManifest.bat %PROJECTNAME% .exe /NOV6CC
	) else (
		call generateAssemblyManifest.bat %PROJECTNAME% .exe
	)
) else (
	if defined NOV6CC (
		call generateAssemblyManifest.bat %PROJECTNAME% .exe /NOV6CC /NOEMBED
	) else (
		call generateAssemblyManifest.bat %PROJECTNAME% .exe /NOEMBED
	)
)
if errorlevel 1 goto :err

if defined FOLDER popd %FOLDER%
exit /B 0

:err
if defined FOLDER popd %FOLDER%
exit /B 1


