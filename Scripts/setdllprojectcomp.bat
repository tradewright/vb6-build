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
echo Sets the binary compatibility for a VB6 dll or ocx project
echo.
echo Usage:
echo.
echo setDllProjectComp projectName [path] [/T:{DLL^|OCX}] [/B:{P^|B^|N}]
echo                     
echo.
echo   projectName             Project name (excluding version)
echo.
echo   path                    Path to folder containing project file
echo.
echo   /T                      Project type: DLL (default) or OCX
echo.
echo   /B                      Binary compatibility: 
echo                               B =^> binary compatibility (default)
echo                               P =^> project compatibility
echo                               N =^> no compatibility
echo.
exit /B
::===0=========1=========2=========3=========4=========5=========6=========7=========8

:doIt
set BINARY_COMPAT=B
set EXTENSION=dll
set MANIFEST=NONE
set PROJECTNAME=
set FOLDER=

:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if /I "%ARG%" == "/T:DLL" (
	set EXTENSION=dll
) else if /I "%ARG%" == "/T:OCX" (
	set EXTENSION=ocx
) else if /I "%ARG:~0,3%" == "/T:" (
	set EXTENSION=
) else if /I "%ARG%" == "/B:P" (
	set BINARY_COMPAT=P
) else if /I "%ARG%" == "/B:B" (
	set BINARY_COMPAT=B
) else if /I "%ARG%" == "/B:N" (
	set BINARY_COMPAT=N
) else if /I "%ARG:~0,3%" == "/B:" (
	set BINARY_COMPAT=
) else if "%ARG:~0,1%"=="/" (
	echo Invalid parameter '%ARG%'
	set ERROR=1
) else if not defined PROJECTNAME (
	set PROJECTNAME=%ARG%
) else if not defined FOLDER (
	set FOLDER=%ARG%
	pushd !FOLDER!
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
	echo projectName parameter missing
	set ERROR=1
)
if not defined EXTENSION (
	echo /T:{DLL^|OCX} setting missing or invalid
	set ERROR=1
)
if not defined BINARY_COMPAT (
	echo /B:{P^|B^|N} setting missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

set FILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%

echo Setting binary compatibility mode = %BINARY_COMPAT%; version = %VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%
echo ... for file: %PROJECTNAME%.vbp 
setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-MAJOR% %VB6-BUILD-MINOR% %VB6-BUILD-REVISION% -mode:%BINARY_COMPAT%
if errorlevel 1 goto :err

if defined FOLDER popd %FOLDER%
exit /B 0

:err
if defined FOLDER popd %FOLDER%
exit /B 1

