@echo off & setlocal

:: -----------------------------------------------------------------------------
:: Initialization
:: -----------------------------------------------------------------------------

::Set working directory to script's path
pushd "%~dp0"


:: -----------------------------------------------------------------------------
:: Basic configuration WMI Provider
:: -----------------------------------------------------------------------------

set "Implevel=/implevel:impersonate"
set "AuthLevel=/authlevel:PktPrivacy"
set "Computer=/node:."
set "NameSpace=/namespace:\\root\Subscription"
set "FailFast=/failfast:on"
set "WMIConfig=%FailFast% %Implevel% %AuthLevel% %Computer% %NameSpace%"


:: *****************************************************************************
:: Project specific configuration, adapt this to your needs
:: *****************************************************************************

:: Name prefix for event filter and event consumer
set "ProjectName=LAN Adapter Status Change"

:: Event filter parameter
set "EventQuery=SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_NetworkAdapter' AND TargetInstance.Name='<Your LAN adapter name like shown in device manager>'"
set "EventNamespace=root\cimv2"

:: Event consumer parameter
set "ScriptToExecute=%ProjectName: =%EventHandler.vbs"
set "ScriptParams="%%TargetInstance.NetConnectionStatus%%" "<Your WLAN PNP DeviceId>""
set "ShowWindow=0"

:: *****************************************************************************
:: *****************************************************************************


:: -----------------------------------------------------------------------------
:: Common configuration
:: -----------------------------------------------------------------------------

:: Event filter parameter
set "EventFilterName=%ProjectName% Event Filter"

:: Event consumer parameter
set "EventConsumerName=%ProjectName% Commandline Event Consumer"
set "CommandLineTemplate="%SystemRoot:\=\\%\\System32\\cscript.exe" /nologo "%CD:\=\\%\\%ScriptToExecute%" %ScriptParams%"
set "ExecutablePath=%SystemRoot:\=\\%\\System32\\cscript.exe"
set "WorkingDirectory=%CD:\=\\%"

set "ElevateScript=%Temp%\ElevateEventConsumerInstaller.vbs"


:: -----------------------------------------------------------------------------
:: Parse command line
:: -----------------------------------------------------------------------------

set "Task="

if /i "%~1" equ "/i" set "Task=Install"
if /i "%~1" equ "/install" set "Task=Install"

if /i "%~1" equ "/u" set "Task=Uninstall"
if /i "%~1" equ "/uninstall" set "Task=Uninstall"

if not defined Task (
  set "ERRORLEVEL=1"
  goto :ErrorExit
)


::------------------------------------------------------------------------------
:: Startup sequence
::------------------------------------------------------------------------------

::Check for admin permissions and restart elevated if required
call :CheckForAdminPermissions || (
  call :RestartElevated "%~f0" "%~1"
  goto :Terminate
)


:: -----------------------------------------------------------------------------
:: Dispatcher
:: -----------------------------------------------------------------------------

if /i "%Task%" equ "Install" goto :Install
if /i "%Task%" equ "Uninstall" goto :Uninstall

set "ERRORLEVEL=1"
goto :ErrorExit


:: -----------------------------------------------------------------------------
:: Install permanent event consumer
:: -----------------------------------------------------------------------------

:Install
:: Create event filter
wmic %WMIConfig% path __EventFilter create Name="%EventFilterName%", EventNamespace="%EventNamespace%", QueryLanguage="WQL", Query="%EventQuery%"
echo(
if ERRORLEVEL 1 goto :ErrorExit

:: Create commandline event consumer
wmic %WMIConfig% path CommandLineEventConsumer create Name="%EventConsumerName%", CommandLineTemplate='"%CommandLineTemplate%"', ExecutablePath="%ExecutablePath%", WorkingDirectory="%WorkingDirectory%", ShowWindowCommand="%ShowWindow%"
echo(
if ERRORLEVEL 1 goto :ErrorExit

:: Create filter-consumer binding
wmic %WMIConfig% path __FilterToConsumerBinding create Filter="__EventFilter.Name=\"%EventFilterName%\"", Consumer="CommandLineEventConsumer.Name=\"%EventConsumerName%\""
echo(
if ERRORLEVEL 1 goto :ErrorExit

goto :Quit


:: -----------------------------------------------------------------------------
:: Uninstall permanent event consumer
:: -----------------------------------------------------------------------------

:Uninstall
:: Delete filter-consumer binding
wmic %WMIConfig% path __FilterToConsumerBinding where Consumer="CommandLineEventConsumer.Name='%EventConsumerName%'" delete
echo(

:: Delete event filter
wmic %WMIConfig% path __EventFilter where Name="%EventFilterName%" delete
echo(

:: Delete event consumer
wmic %WMIConfig% path CommandLineEventConsumer where Name="%EventConsumerName%" delete
echo(

goto :Quit


::------------------------------------------------------------------------------
:: Error handling
::------------------------------------------------------------------------------

:ErrorExit
if "%ERRORLEVEL%" equ "1" (
  echo Install permanent WMI command line event consumer to process background task.
  echo(
  echo %~n0 {/i ^| /u}
  echo(
  echo   /i  Install WMI event consumer
  echo   /u  Uninstall WMI event consumer
  echo(
) else (
  echo(
  echo Installation aborted because an error occured!
  echo(
)


::------------------------------------------------------------------------------
:: Script termination
::------------------------------------------------------------------------------

:Quit
::Clean up
del "%ElevateScript%" 1>NUL 2>&1

pause

:Terminate
popd

exit /b %ERRORLEVEL%



::==============================================================================
:: Subroutines
::==============================================================================

:CheckForAdminPermissions
  net session 1>NUL 2>&1
  if ERRORLEVEL 1 exit /b 1
exit /b 0


:RestartElevated
  ::Get system's ANSI and OEM code page and set console's code page to ANSI code page
  ::This is required if this script is stored in a path that contains characters
  ::with different code points in those code pages.
  for /f "tokens=2 delims==" %%a in ('wmic OS get CodeSet /format:list') do set /a "ACP=%%~a"
  for /f "tokens=2 delims=.:" %%a in ('chcp') do set /a "OEMCP=%%a"
  if "%ACP%" neq "" if "%ACP%" neq "0" chcp %ACP% > NUL

  > "%ElevateScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.strApplication = "cmd.exe"
  >>"%ElevateScript%" echo.strArguments   = "/c """"%~1"" ""%~2"""""
  >>"%ElevateScript%" echo.
  >>"%ElevateScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 1

  ::Restore OEM code page
  if "%OEMCP%" neq "" if "%OEMCP%" neq "0" chcp %OEMCP% > NUL

  cscript /nologo "%ElevateScript%"
exit /b 0
