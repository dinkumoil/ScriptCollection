@echo off & setlocal

pushd "%~dp0"


::------------------------------------------------------------------------------
:: Retrieve arguments from file name
::------------------------------------------------------------------------------
set "ScriptFileName=%~n0"
set "NppArchitecture=/a:%ScriptFileName:~-3%"

if /i "%NppArchitecture%" neq "/a:x64" (
  if /i "%NppArchitecture%" neq "/a:x86" (
    set "NppArchitecture=/a:x86"
  )
)


::------------------------------------------------------------------------------
:: Retrieve arguments from command line
::------------------------------------------------------------------------------
set "NppBasePath="
set "Task="
set "Proxy="

:ParseArgsLoop
  if /i "%~1" equ "/a" (
    set "NppArchitecture=/a:%~2"
    shift

  ) else if /i "%~1" equ "-a" (
    set "NppArchitecture=/a:%~2"
    shift

  ) else if /i "%~1" equ "/n" (
    set NppBasePath=/n:"%~2"
    shift

  ) else if /i "%~1" equ "-n" (
    set NppBasePath=/n:"%~2"
    shift

  ) else if /i "%~1" equ "/t" (
    set Task=/t:"%~2"
    shift

  ) else if /i "%~1" equ "-t" (
    set Task=/t:"%~2"
    shift

  ) else if /i "%~1" equ "/p" (
    set Proxy=/p:"%~2"
    shift

  ) else if /i "%~1" equ "-p" (
    set Proxy=/p:"%~2"
    shift

  ) else if /i "%~1" equ "/?" (
    call :ShowHelp
    exit /b 0

  ) else if /i "%~1" equ "-?" (
    call :ShowHelp
    exit /b 0

  ) else dir /b /a:-d "%~1\notepad++.exe" 1>NUL 2>&1 && (
    set NppBasePath=/n:"%~1"
  )

  shift
if "%~1" neq "" goto :ParseArgsLoop


::------------------------------------------------------------------------------
:: Execute Plugins Admin script
::------------------------------------------------------------------------------
:ExecPluginManagement
  mode con: cols=230
  start "" /b /wait /d "%CD%" cscript.exe /nologo "%CD%\script\NppPluginManagement.vbs" %NppArchitecture% %Task% %NppBasePath% %Proxy%
  echo( & pause

  if ERRORLEVEL 1 goto :Terminate
goto :ExecPluginManagement

:Terminate
exit /b 0



::==============================================================================
:: Show help message
::==============================================================================

:ShowHelp
  echo(
  echo Manage Notepad++ plugins: Download, install, update and list
  echo(
  echo %~n0 [/a ^<x86^|x64^>] [/t ^<Task^>] [/n ^<N++ base path^>] [/p ^<proxy URL^>]
  echo(
  echo   /a       Specifies the architecture of your Notepad++ installation.
  echo            Use x86 for a 32 bit installation and x64 for 64 bit.
  echo            The default is x86. It is also possible to append one of
  echo            those IDs to the filename of this script.
  echo(
  echo   /t       Allows to specify the task to perform:
  echo              download, install, update, updateall, list
  echo            If omitted a menu to select the task is displayed.
  echo(
  echo   /n       Allows to specify the base path of a Notepad++ installation.
  echo            If omitted the standard path for the provided architecture
  echo            is used.
  echo(
  echo   /p       Allows to specify a proxy. ^<proxy URL^> has to be the proxy
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
