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
echo Builds a VB6 dll or ocx project
echo.
echo Usage:
echo.
echo makedll projectName [path] [/O:objFileName] [/T:{DLL^|OCX}] [/B:{P^|PP^|B^|N}]
echo                     [/M:{N^|E^|F}] [/NOVERSION] [/C] [/INLINE] [/DEP:depFilename]
echo.
echo   projectName             Project name (excluding version)
echo.
echo   path                    Path to folder containing project file
echo.
echo   /O:objFilename          Objectfile name (no extension)
echo.
echo   /T                      Project type: DLL (default) or OCX
echo.
echo   /B                      Binary compatibility: 
echo                               B  =^> binary compatibility (default)
echo                               P  =^> project compatibility
echo                               PP =^> project compatibility (and don't set to 
echo                                      binary compatibility after building)
echo                               N  =^> no compatibility
echo.
echo   /M                      Manifest requirement:
echo                               N =^> no manifest (default)
echo                               E =^> embed manifest in object file
echo                               F =^> freestanding manifest file
echo.
echo   /NOVERSION              Indicates that the major and minor version numbers
echo                           are not to be included as a suffix in the object
echo                           filename.
echo.
echo   /C                      Indicates the compatibility location is the
echo                           project's Compat subfolder rather than the Bin 
echo                           folder
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
set BINARY_COMPAT=B
set EXTENSION=dll
set MANIFEST=NONE
set PROJECTNAME=
set FOLDER=
set NOVERSION=NO

:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if /I "%ARG:~0,3%"=="/O:" (
	set OBJECTFILENAME=%ARG:~3%
) else if /I "%ARG%" == "/T:DLL" (
	set EXTENSION=dll
) else if /I "%ARG%" == "/T:OCX" (
	set EXTENSION=ocx
) else if /I "%ARG:~0,3%" == "/T:" (
	set EXTENSION=
) else if /I "%ARG%" == "/B:P" (
	set BINARY_COMPAT=P
) else if /I "%ARG%" == "/B:PP" (
	set BINARY_COMPAT=PP
) else if /I "%ARG%" == "/B:B" (
	set BINARY_COMPAT=B
) else if /I "%ARG%" == "/B:N" (
	set BINARY_COMPAT=N
) else if /I "%ARG:~0,3%" == "/B:" (
	set BINARY_COMPAT=
) else if /I "%ARG%" == "/M:N" (
	set MANIFEST=NONE
) else if /I "%ARG%" == "/M:E" (
	set MANIFEST=EMBED
) else if /I "%ARG%" == "/M:F" (
	set MANIFEST=NOEMBED
) else if /I "%ARG:~0,3%" == "/M:" (
	set MANIFEST=
) else if /I "%ARG:~0,10%" == "/NOVERSION" (
	set NOVERSION=YES
) else if /I "%ARG%" == "/C" (
	set COMPAT=yes
) else if /I "%ARG%" == "/INLINE" (
	set INLINE=/INLINE
) else if "%ARG:~0,5%"=="/DEP:" (
	set DEP="%ARG%"
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

if not defined BIN-PATH (
	echo ^%%BIN-PATH^%% is not defined
	set ERROR=1
)
if not defined PROJECTNAME (
	echo projectName parameter missing
	set ERROR=1
)
if not defined EXTENSION (
	echo /T:{DLL^|OCX} setting missing or invalid
	set ERROR=1
)
if not defined BINARY_COMPAT (
	echo /B:{P^|PP^|B^|N} setting missing or invalid
	set ERROR=1
)
if not defined MANIFEST (
	echo /M:{N^|E^|F} setting missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

if not defined OBJECTFILENAME (
        set OBJECTFILENAME=%PROJECTNAME%
)
if not "%NOVERSION%"=="YES" (
	set OBJECTFILENAME=%OBJECTFILENAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%
)

echo =================================
if defined FOLDER (
	echo Building %FOLDER%\%PROJECTNAME%.vbp =^> %OBJECTFILENAME%.%EXTENSION%
) else (
	echo Building %PROJECTNAME%.vbp =^> %OBJECTFILENAME%.exe
)

if not exist Prev (
	echo Making Prev directory
	mkdir Prev 
)

if exist %BIN-PATH%\%OBJECTFILENAME%.%EXTENSION% (
	echo Copying previous binary
	copy %BIN-PATH%\%OBJECTFILENAME%.%EXTENSION% Prev\* 
)

echo Setting binary compatibility mode = %BINARY_COMPAT:~0,1%; version = %VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%
echo ... for file: %PROJECTNAME%.vbp

setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-MAJOR% %VB6-BUILD-MINOR% %VB6-BUILD-REVISION% -mode:%BINARY_COMPAT:~0,1% -ObjectFileName:%OBJECTFILENAME%.%EXTENSION%
if errorlevel 1 goto :err

echo Compiling
vb6.exe /m %PROJECTNAME%.vbp
if errorlevel 1 goto :err

if exist %BIN-PATH%\%OBJECTFILENAME%.lib (
	echo Deleting .lib file
	del %BIN-PATH%\%OBJECTFILENAME%.lib 
)
if exist %BIN-PATH%\%OBJECTFILENAME%.exp (
	echo Deleting .exp file
	del %BIN-PATH%\%OBJECTFILENAME%.exp 
)

if "%BINARY_COMPAT%"=="PP" (
	echo Leaving binary compatibility mode = P
	setprojectcomp.exe %PROJECTNAME%.vbp 1 0 0 -mode:P -n
	if errorlevel 1 goto :err
) ELSE (
	echo Setting binary compatibility mode = B
	setprojectcomp.exe %PROJECTNAME%.vbp 1 0 0 -mode:B -n
	if errorlevel 1 goto :err
)

if defined COMPAT (
	if not exist Compat (
		echo Making Compat directory
		mkdir Compat
		if errorlevel 1 goto :err
	)
	if not "%BINARY_COMPAT%" == "B" (
		echo Copying binary to Compat
		copy %BIN-PATH%\%OBJECTFILENAME%.%EXTENSION% COMPAT\* 
		if errorlevel 1 goto :err
	)
)

if "%MANIFEST%"=="NONE" (
	echo don't generate a manifest>nul
) else (
	:: NB: the following line sets a space in SWITCHES
	set SWITCHES= 
	if defined INLINE set SWITCHES=!SWITCHES! %INLINE%
	if defined DEP set SWITCHES=!SWITCHES! %DEP%
	if "%MANIFEST%"=="NOEMBED" set SWITCHES=!SWITCHES! /NOEMBED
	call generateAssemblyManifest.bat %PROJECTNAME% %EXTENSION% !SWITCHES!
	if errorlevel 1 goto :err
)

if defined FOLDER popd %FOLDER%
exit /B 0

:err
if defined FOLDER popd %FOLDER%
exit /B 1

