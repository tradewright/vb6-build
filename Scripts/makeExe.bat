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
echo makeExe projectName [path] [/M:{N^|E^|F}] [/CONSOLE] [/NOV6CC] [/INLINE]
echo                     [/DEP:depFilename]
echo.
echo   projectName             Project name (excluding version).
echo.
echo   path                    Path to folder containing project file.
echo.
echo   /M                      Manifest requirement:
echo                               N =^> no manifest (default)
echo                               E =^> embed manifest in object file
echo                               F =^> freestanding manifest file
echo.
echo   /CONSOLE                Link the exe to the Console library.
echo.
echo   /NOV6CC                 Don't use Version 6 Common Controls.
echo.
echo   /DEP:depFilename        Specifies a file containing external dependencies
echo                           for this project. See below for further details.
echo.
echo   /INLINE                 Specifies that ^<dependentAsssembly^> XML elements 
echo                           are not to be be included in the manifest for 
echo                           external references. Rather a ^<file^> element 
echo                           containing the COM class information is to be 
echo                           generated for each external reference. Ignored 
echo                           if a projectFileName.man file exists.
echo.
echo   On entry, ^%%BIN-PATH^%% must be set to the folder where the compiled object
echo   file will be stored.
echo.
echo   depFilename
echo     Contains details of one dependent assembly per line. Each line is formatted
echo     as a complete ^<assemblyIdentity^> XML element. Blank lines and lines that
echo     start with // are ignored.
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
) else if /I "%ARG%" == "/INLINE" (
	set INLINE=/INLINE
) else if "%ARG:~0,5%"=="/DEP:" (
	set DEP=%ARG%
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
	echo Building %FOLDER%\%FILENAME%
) else (
	echo Building %FILENAME%
)

echo Setting version = %VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%
setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-MAJOR% %VB6-BUILD-MINOR% %VB6-BUILD-REVISION% -mode:N
if errorlevel 1 goto :err

:: Delay for a short time to prevent VB6 being run too frequently which
:: can cause problems (now commented out as not sure it's needed)
::PING localhost -n 2 >NUL

vb6.exe /m %PROJECTNAME%.vbp /outdir %BIN-PATH%
if errorlevel 1 goto :err

if defined LINKTOCONSOLE (
	echo Linking CONSOLE
	link /EDIT /SUBSYSTEM:CONSOLE %BIN-PATH%\%FILENAME%
	if errorlevel 1 goto :err
)

if "%MANIFEST%"=="NONE" (
	echo don't generate a manifest>nul
) else ( 
	:: NB: the following line sets a space in SWITCHES
	set SWITCHES= 
	if defined NOV6CC set SWITCHES=!SWITCHES! /NOV6CC
	if defined INLINE set SWITCHES=!SWITCHES! %INLINE%
	if defined DEP set SWITCHES=!SWITCHES! %DEP%
	if "%MANIFEST%"=="NOEMBED" set SWITCHES=!SWITCHES! /NOEMBED
	call generateAssemblyManifest.bat %PROJECTNAME% exe !SWITCHES!
	if errorlevel 1 goto :err
)

setprojectcomp.exe %PROJECTNAME%.vbp 1 0 0 -mode:N

if defined FOLDER popd %FOLDER%
exit /B 0

:err
if defined FOLDER popd %FOLDER%
exit /B 1


