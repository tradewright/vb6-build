@echo off
setlocal 

if "%1"=="" goto :showUsage
if "%1"=="/?" goto :showUsage
if "%1"=="-?" goto :showUsage
if /I "%1"=="/HELP" goto :showUsage
goto :doIt

:showUsage
::===0=========1=========2=========3=========4=========5=========6=========7=========8
echo.
echo Builds the manifest for an existing dll or ocx object file
echo.
echo Usage:
echo.
echo generateAssemblyManifest objectName objectType
echo.
echo   objectName               object filename (excluding extension)
echo   objectType               object type ('dll' or 'ocx')
echo.
echo   On entry, ^%%BIN-PATH^%% must be set to the folder where the 
echo   generated manifest will be stored.
exit /B
::===0=========1=========2=========3=========4=========5=========6=========7=========8

:doIt
:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if not defined OBJECTNAME (
	set OBJECTNAME=%ARG%
) else if not defined EXTENSION (
	if /I "%ARG%"=="DLL" (
		set EXTENSION=dll
	) else if /I "%ARG%"=="OCX" (
		set EXTENSION=ocx
	) else (
	echo Invalid parameter '%ARG%'
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
if not defined OBJECTNAME (
	echo objectName parameter missing
	set ERROR=1
)
if not defined EXTENSION (
	echo objectType parameter missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

echo Generating manifest for %BIN-PATH%\%OBJECTNAME%.%EXTENSION% to %BIN-PATH%\%OBJECTNAME%.manifest

path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%

GenerateManifest /Bin:%BIN-PATH%\%OBJECTNAME%.%EXTENSION% /Out:%BIN-PATH%\%OBJECTNAME%.manifest
if errorlevel 1 goto :err

:: ensure mt.exe can find object files when hashing
pushd %BIN-PATH%

echo Updating manifest hash
mt.exe -manifest %BIN-PATH%\%OBJECTNAME%.manifest -hashupdate -nologo
if errorlevel 1 (
	popd
	goto :err
)
popd

exit /B 0

:err
exit /B 1


