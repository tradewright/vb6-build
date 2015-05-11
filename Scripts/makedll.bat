@echo off

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
echo.
echo Builds a VB6 dll or ocx project
echo.
echo Usage:
echo.
echo makedll projectName [path] type binaryCompatibility [/COMPAT]
echo.
echo   projectName      Project name (excluding version)
echo   path             Path to folder containing project file
echo   type             Project type ('.dll' or '.ocx')
echo   binaryCompat     Binary compatibility: 
echo                    'P' project compatibility
echo                    'B' binary compatibility
echo                    'N' no compatibility
echo   /COMPAT          Indicates the compatibility location is 
echo                    the Compat subfolder rather than the Bin 
echo                    folder
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

if /I "%~2" == ".DLL" (
	set EXTENSION=.dll
) else if /I "%~2" == ".OCX" (
	set EXTENSION=.ocx
) else (
	echo Invalid project type '%~2'
	goto :err
)

if /I "%~3" == "P" (
	set BINARY_COMPAT=P
) else if /I "%~3" == "B" (
	set BINARY_COMPAT=B
) else if /I "%~3" == "N" (
	set BINARY_COMPAT=N
) else if not "%~3" == "" (
	echo Invalid binaryCompat '%~3'
	goto :err
)

set COMPAT=no
if /I "%~4" == "/COMPAT" (
	set COMPAT=yes
) else if not "%~4" == "" (
	echo Invalid parameter '%~4'
	goto :err
)
	
set FILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%

echo =================================
if defined FOLDER (
	echo Building %FOLDER%%FILENAME%%EXTENSION%
) else (
	echo Building %FILENAME%%EXTENSION%
)

if not exist Prev (
	echo Making Prev directory
	mkdir Prev 
)

if exist %BIN-PATH%\%FILENAME%%EXTENSION% (
	echo Copying previous binary
	copy %BIN-PATH%\%FILENAME%%EXTENSION% Prev\* 
)

echo Setting binary compatibility mode = %BINARY_COMPAT%; version = %VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.%VB6-BUILD-REVISION%
echo ... for file: %PROJECTNAME%.vbp 
setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-REVISION% -mode:%BINARY_COMPAT%
if errorlevel 1 goto :err

echo Compiling
vb6.exe /m %PROJECTNAME%.vbp
if errorlevel 1 goto :err

if exist %BIN-PATH%\%FILENAME%.lib (
	echo Deleting .lib file
	del %BIN-PATH%\%FILENAME%.lib 
)
if exist %BIN-PATH%\%FILENAME%.exp (
	echo Deleting .exp file
	del %BIN-PATH%\%FILENAME%.exp 
)

echo Setting binary compatibility mode = B
setprojectcomp.exe %PROJECTNAME%.vbp %VB6-BUILD-REVISION% -mode:B
if errorlevel 1 goto :err

if "%COMPAT%" == "yes" (
	if not exist Compat (
		echo Making Compat directory
		mkdir Compat
		if errorlevel 1 goto :err
	)
	if not "%BINARY_COMPAT%" == "B" (
		echo Copying binary to Compat
		copy %BIN-PATH%\%FILENAME%%EXTENSION% COMPAT\* 
		if errorlevel 1 goto :err
	)
)

call generateAssemblyManifest.bat %PROJECTNAME% %EXTENSION%
if errorlevel 1 goto :err

if defined FOLDER popd %FOLDER%
exit /B 0

:err
pause
if defined FOLDER popd %FOLDER%
exit /B 1

