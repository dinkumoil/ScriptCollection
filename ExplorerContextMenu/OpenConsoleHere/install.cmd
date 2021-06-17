@echo off & setlocal

::------------------------------------------------------------------------------
:: Basic configuration
::------------------------------------------------------------------------------
set "DestFolder="
set "FilesToCopy="
set "RegFileToExecute=OpenConsole.reg"

::Path for temporary elevation script
set "ElevateScript=%Temp%\Elevate.vbs"


::------------------------------------------------------------------------------
:: Startup sequence
::------------------------------------------------------------------------------
::Set working directory to script's path
pushd "%~dp0"

::Check for admin permissions and restart elevated if required
call :CheckForAdminPermissions || (
  call :RestartElevated "%~f0"
  goto :Terminate
)


::------------------------------------------------------------------------------
:: Perform installation
::------------------------------------------------------------------------------
if defined DestFolder (
  md "%DestFolder%" 2>NUL

  for %%a in (%FilesToCopy%) do (
    copy "%%~a" "%DestFolder%" > NUL
  )
)

start "" regedit /s "%RegFileToExecute%"

::Clean up
del "%ElevateScript%" 1>NUL 2>&1


::------------------------------------------------------------------------------
:: Script termination
::------------------------------------------------------------------------------
:Terminate
popd
exit /b 0



::==============================================================================
:: Subroutines
::==============================================================================

:CheckForAdminPermissions
  net session 1>NUL 2>&1
  if ERRORLEVEL 1 exit /b 1
exit /b 0


:RestartElevated
  ::Get system's ANSI and OEM code page and set console's code page to ANSI code page.
  ::This is required if this script is stored in a path that contains characters
  ::with different code points in those code pages.
  for /f "tokens=2 delims==" %%a in ('wmic OS get CodeSet /format:list') do set /a "ACP=%%~a"
  for /f "tokens=2 delims=.:" %%a in ('chcp') do set /a "OEMCP=%%a"
  if "%ACP%" neq "" if "%ACP%" neq "0" chcp %ACP% > NUL

  > "%ElevateScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.strApplication = "cmd.exe"
  >>"%ElevateScript%" echo.strArguments   = "/c ""%~1"""
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 1

  ::Restore OEM code page
  if "%OEMCP%" neq "" if "%OEMCP%" neq "0" chcp %OEMCP% > NUL

  cscript /nologo "%ElevateScript%"
exit /b 0
