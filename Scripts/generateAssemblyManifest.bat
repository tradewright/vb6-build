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
echo Builds the manifest for a VB6 exe, dll or ocx project
echo.
echo Usage:
echo.
echo generateAssemblyManifest projectName projectType [/NOEMBED] [/NOV6CC] [/INLINE]
echo                          [/DEP:depFilename]
echo.
echo   projectName             Project name (excluding version).
echo.
echo   projectType             Project type ('dll' or 'ocx' or 'exe').
echo.
echo   /NOEMBED                Don't embed manifest as resource.
echo.
echo   /NOV6CC                 Don't use Version 6 Common Controls.
echo.
echo   /DEP:depFilename        Specifies a file containing external dependencies
echo                           for this project. See below for further details.
echo.
echo   /INLINE                 Specifies that ^<dependentAsssembly^> XML elements are
echo                           not to be be included for external references. Rather
echo                           a ^<file^> element containing the COM class information
echo                           is to be generated for each external reference. 
echo                           Ignored if a projectFileName.man file exists.
echo.
echo   On entry, ^%%BIN-PATH^%% must be set to the folder where the 
echo   generated manifest will be stored.
echo.
echo   depFilename
echo     Contains details of one dependent assembly per line. Each line is formatted
echo     as a complete ^<assemblyIdentity^> XML element. Blank lines and lines that
echo     start with // are ignored.
exit /B
::===0=========1=========2=========3=========4=========5=========6=========7=========8

:doIt

:: don't inherit variables from caller
set PROJECTNAME=
set EXTENSION=
set NOEMBED=
set NOV6CC=

:parse

if "%~1" == "" goto :parsingComplete

set ARG=%~1
if /I "%ARG%" == "/NOEMBED" (
	set NOEMBED=NOEMBED
) else if /I "%ARG%" == "/NOV6CC" (
	set NOV6CC=/NOV6CC
) else if /I "%ARG%" == "/INLINE" (
	set INLINE=/INLINE
) else if "%ARG:~0,5%"=="/DEP:" (
	set DEP=%ARG%
) else if "%ARG:~0,1%"=="/" (
	echo Invalid parameter '%ARG%'
	set ERROR=1
) else if not defined PROJECTNAME (
	set PROJECTNAME=%ARG%
) else if not defined EXTENSION (
	if /I "%ARG%" == "DLL" (
		set NOV6CC=/NOV6CC
		set EXTENSION=dll
	) else if /I "%ARG%" == "OCX" (
		set NOV6CC=/NOV6CC
		set EXTENSION=ocx
	) else if /I "%ARG%" == "EXE" (
		set EXTENSION=exe
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
if not defined PROJECTNAME (
	echo projectName parameter missing
	set ERROR=1
)
if not defined EXTENSION (
	echo projectType parameter missing or invalid
	set ERROR=1
)
if defined ERROR goto :err

echo Generating manifest for %PROJECTNAME%

path C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin;%PATH%

set OBJECTFILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.%EXTENSION%
if "%EXTENSION%"=="exe" (
        set MANIFESTFILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.exe.manifest
) else (
        set MANIFESTFILENAME=%PROJECTNAME%%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%.manifest
)

:: NB: the following line sets a space in SWITCHES
set SWITCHES= 
if not defined NOV6CC set SWITCHES=/V6CC
if defined INLINE set SWITCHES=%SWITCHES% /INLINE
if defined DEP set SWITCHES=%SWITCHES% %DEP%

GenerateManifest /Proj:%PROJECTNAME%.vbp /Out:%BIN-PATH%\%MANIFESTFILENAME% %SWITCHES%
if errorlevel 1 goto :err

:: ensure mt.exe can find object files when hashing
:: pushd %BIN-PATH%
:: 
:: echo Updating manifest hash
:: mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -hashupdate -nologo
:: if errorlevel 1 (
:: 	popd		
:: 	goto :err
:: )

popd

if not defined NOEMBED (
	if "%EXTENSION%"=="ocx" (
		echo Can't embed manifest as resource for ocx
	) else (
		echo Embedding manifest as a resource
:retry
		mt.exe -manifest %BIN-PATH%\%MANIFESTFILENAME% -outputresource:%BIN-PATH%\%OBJECTFILENAME%;#1 -nologo
		if errorlevel 1 (
			echo mt.exe returned error %errorlevel% - retrying
			goto :retry
		)
		echo Deleting manifest file
		del %BIN-PATH%\%MANIFESTFILENAME%
	)
)

exit /B 0

:err
exit /B 1


