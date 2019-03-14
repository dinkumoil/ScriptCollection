@echo off & setlocal

::Initialisierung
set /a "StartLine=0"
set /a "StopLine=0"
set /a "ShowLineNumbers=0"
set "InFile="


::Kommandozeilenargumente einlesen
:ParseArgsLoop
  if /i "%~1" equ "/b" (
    set /a "StartLine=%~2" & shift
  ) else if /i "%~1" equ "/e" (
    set /a "StopLine=%~2" & shift
  ) else if /i "%~1" equ "/n" (
    set /a "ShowLineNumbers=1"
  ) else (
    set "InFile=%~1"
  )

  shift
if "%~1" neq "" goto :ParseArgsLoop


::Check arguments
if not exist "%InFile%" (
  echo Input file not found.
  exit /b 1
)

if %StartLine% lss 1 (
  set /a "StartLine=1"
)

if %StopLine% lss 1 (
  for /f "tokens=1 delims=:" %%a in ('findstr /n "^" "%InFile%"') do set /a "StopLine=%%a"
)

if %StopLine% lss %StartLine% (
  echo The terminating line can not be less than the starting line.
  exit /b 2
)


::Value of starting line has to be corrected
set /a StLn=StartLine-1

if %StLn% geq 1 (
  set "LinesToSkip=skip=%StLn%"
) else (
  set "LinesToSkip="
)

::Do the CUT magic
for /f "%LinesToSkip% tokens=1,2* delims==" %%a in ('^(for /l %%a in ^(1,1,%StopLine%^) do @^(set "TmpVar=" ^& set /p "TmpVar=" ^& set /p "=%%a=" ^<NUL ^& set TmpVar 2^>NUL^) ^|^| echo TmpVar^=^) ^< "%InFile%"') do (
  if "%ShowLineNumbers%" equ "1" (
    echo(%%a: %%c
  ) else (
    echo(%%c
  )
)
