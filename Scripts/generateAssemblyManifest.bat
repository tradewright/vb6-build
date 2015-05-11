@echo off
setlocal 

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
echo.
echo Builds the manifest for a VB6 exe, dll or ocx project
echo.
echo Usage:
echo.
echo generateAssemblyManifest projectName projectType [/NOEMBED] [/NOV6CC]
echo.
echo   projectName		project name (excluding version)
echo   projectType              project type ('.dll' or '.ocx' or '.exe')
echo   /NOEMBED                 don't embed manifest as resource
echo   /NOV6CC                  don't use Version 6 Common Controls
echo.
exit /B

:doIt
echo Generating manifest for %1

path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%

if /I "%~2" == ".DLL" (
	set EXTENSION=.dll
) else if /I "%~2" == ".OCX" (
	set EXTENSION=.ocx
) else if /I "%~2" == ".EXE" (
	set EXTENSION=.exe
) else (
	echo Invalid project type %~2
	goto :err
)

if /I "%~3" == "/NOEMBED" (
	set NOEMBED=NOEMBED
) else if /I "%~3" == "/NOV6CC" (
	set NOV6CC=NOV6CC
) else if not "%~3" == "" (
	echo Invalid parameter %3
	goto :err
)

if /I "%~4" == "/NOEMBED" (
	set NOEMBED=NOEMBED
) else if /I "%~4" == "/NOV6CC" (
	set NOV6CC=NOV6CC
) else if not "%~4" == "" (
	echo Invalid parameter %4
	goto :err
)

set OBJECTFILENAME=%1%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%%EXTENSION%
if "%EXTENSION%"==".exe" (
        set MANIFESTFILENAME=%1%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.exe.manifest
) else (
        set MANIFESTFILENAME=%1%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.manifest
)

if defined NOV6CC (
	GenerateManifest /Proj:%1.vbp /Out:%BIN-PATH%\%MANIFESTFILENAME%
) else (
	GenerateManifest /Proj:%1.vbp /Out:%BIN-PATH%\%MANIFESTFILENAME% /V6CC
)
if errorlevel 1 goto :err

:: ensure mt.exe can find object files when hashing
pushd %BIN-PATH%

echo Updating manifest hash
mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -hashupdate -nologo
if errorlevel 1 (
	popd		
	goto :err
)

popd

if not defined NOEMBED (
	if "%EXTENSION%"==".ocx" (
		echo Can't embed manifest as resource for .ocx
	) else (
		echo Embedding manifest as a resource
		mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -outputresource:%BIN-PATH%\%OBJECTFILENAME%;#1 -nologo
		if errorlevel 1 goto :err
		echo Deleting manifest file
		del %BIN-PATH%\%MANIFESTFILENAME%
	)
)

exit /B 0

:err
exit /B 1


