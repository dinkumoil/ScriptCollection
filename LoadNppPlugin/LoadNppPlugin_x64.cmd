@echo off & setlocal

::------------------------------------------------------------------------------
:: Retrieve arguments from file name
::------------------------------------------------------------------------------
set "ScriptFileName=%~n0"
set "PluginArchitecture=%ScriptFileName:~-3%"

if /i "%PluginArchitecture%" neq "x64" (
  if /i "%PluginArchitecture%" neq "x86" (
    set "PluginArchitecture=x86"
  )
)


::------------------------------------------------------------------------------
:: Retrieve arguments from command line
::------------------------------------------------------------------------------
set "Proxy="

:ParseArgsLoop
  if /i "%~1" equ "x64" (
    set "PluginArchitecture=x64"

  ) else if /i "%~1" equ "x86" (
    set "PluginArchitecture=x86"

  ) else if /i "%~1" equ "/p" (
    set "Proxy=/p:"%~2""
    shift

  ) else if /i "%~1" equ "-p" (
    set "Proxy=/p:"%~2""
    shift

  ) else if /i "%~1" equ "/?" (
    call :ShowHelp
    exit /b 0

  ) else if /i "%~1" equ "-?" (
    call :ShowHelp
    exit /b 0
  )

  shift
if "%~1" neq "" goto :ParseArgsLoop


::------------------------------------------------------------------------------
:: Execute download script
::------------------------------------------------------------------------------
start "" /b /wait /d "%CD%" cscript.exe /nologo "%CD%\script\LoadNppPlugin.vbs" %PluginArchitecture% %Proxy%
pause

exit /b 0



::==============================================================================
:: Show help message
::==============================================================================

:ShowHelp
  echo Select a Notepad++ plugin from a list and download it.
  echo(
  echo %~nx0 [x86^|x64] [/p:^<proxy URL^>]
  echo(
  echo x86, x64   Specifies the architecture of your Notepad++ installation.
  echo            Use x86 for a 32 bit installation and x64 for 64 bit.
  echo            The default is x86. It is also possible to append one of
  echo            those IDs to the filename of this script.
  echo(
  echo /p         Allows to specify a proxy. ^<proxy URL^> has to be the proxy
  echo            address using the following schema:
  echo(
  echo               [protocol://][user:password@]proxyhost[:port]
  echo(
  echo            Example: http://foo:bar@192.172.10.1:3128
  echo(
  echo            Since this argument is handed over to CURL (which actually
  echo            performs the download) refer to the CURL documentation for
  echo            further information.
exit /b 0
