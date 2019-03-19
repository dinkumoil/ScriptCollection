:: *****************************************************************************
::
:: This is a build script to compile the file SSMSLoginTools.cs using MSBuild
:: and the C# compiler provided with every .NET installation
::
:: *****************************************************************************


@echo off & setlocal

:: =============================================================================
:: Configuration section:
::  - Name of project file
::  - Version number of target framework
::    (can be empty, the newest version
::     available is used then

set "ProjectFileName=SSMSLoginTool.proj"
set "TFV="
:: =============================================================================


:: Set working directory to the script's directory
pushd "%~dp0"

:: Initialize variables
set "Base.NetPath=%SystemRoot%\Microsoft.NET\Framework"
set "MSBuildPath="

:: Scan all .NET directories in alphabetic order for MSBuild.exe. The directory
:: scanned last contains the newest version. For every directory where MSBuild.exe
:: has been found extract the .NET version number and set it as target version
:: number for compilation, but only if no default value has been set in the
:: configuration section.
for /f "delims=" %%a in ('dir /b /a:d /o:ne "%Base.NetPath%\v?.*" 2^>NUL') do (
  for /f "delims=" %%b in ('dir /b /a:-d "%Base.NetPath%\%%a\msbuild.exe" 2^>NUL') do (
    set "MSBuildPath=%Base.NetPath%\%%a\%%b"

    if "%TFV%" equ "" (
      for /f "tokens=1,2 delims=v." %%c in ("%%a") do (
        set "TFV=%%c.%%d"
      )
    )
  )
)

:: If MSBuild.exe has been found build project using the project file
if "%MSBuildPath%" neq "" (
  "%MSBuildPath%" "%ProjectFileName%" /t:Rebuild /tv:%TFV% /p:TargetFrameworkVersion=v%TFV%
) else (
  echo MSBuild.exe wurde nicht gefunden.
)

:Quit
popd
