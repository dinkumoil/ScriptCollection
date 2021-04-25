@echo off & setlocal


::------------------------------------------------------------------------------
::Basic configuration
::------------------------------------------------------------------------------
set "SupportedHashFunctions=MD2, MD4, MD5, SHA1, SHA256, SHA384, SHA512"


::------------------------------------------------------------------------------
::Set default values
::------------------------------------------------------------------------------
set "HashFunction=SHA1"
set "FileToHash="

set /a ExitCode=0


::------------------------------------------------------------------------------
::Main script
::------------------------------------------------------------------------------
pushd "%~dp0"


::..............................................................................
::Retrieve command line arguments
::..............................................................................
:ParseCommandLineLoop
  if /i "%~1" equ "/?" (
    call :ShowHelp
    goto :Terminate

  ) else if /i "%~1" equ "-?" (
    call :ShowHelp
    goto :Terminate

  ) else if /i "%~1" equ "/h" (
    set "HashFunction=%~2"
    shift

  ) else if /i "%~1" equ "-h" (
    set "HashFunction=%~2"
    shift

  ) else (
    set "FileToHash=%~1"
  )

  shift
if "%~1" neq "" goto :ParseCommandLineLoop


::..............................................................................
::Check parameters
::..............................................................................
if not defined FileToHash (
  1>&2 echo File to hash not provided
  set /a ExitCode=1
  goto :Terminate
)

if not exist "%FileToHash%" (
  1>&2 echo File to hash not found
  set /a ExitCode=2
  goto :Terminate
)

call :GetHashFunction || (
  1>&2 echo Unknown hash function "%HashFunction%"
  1>&2 echo Supported hash functions: %SupportedHashFunctions%
  set /a ExitCode=3
  goto :Terminate
)


::..............................................................................
::Calculate hash value of provided file
::..............................................................................
for /f "skip=1 delims=" %%a in ('certutil -hashfile "%FileToHash%" "%HashFunction%" ^| findstr /ivb /c:"CertUtil"') do (
  for %%b in (%%~a) do set /p "=%%~b" < NUL
  echo(
)


::..............................................................................
::Terminate script
::..............................................................................
:Terminate
popd
exit /b %ExitCode%



::==============================================================================
::Decode command line hash function parameter
::==============================================================================
:GetHashFunction
  for %%a in (%SupportedHashFunctions%) do (
    if /i "%HashFunction%" equ "%%~a" (
      set "HashFunction=%%~a"
      exit /b 0
    )
  )
exit /b 1


::==============================================================================
::Show help message
::==============================================================================
:ShowHelp
  echo Calculate hash value of input file.
  echo(
  echo Usage: %~n0 [Drive:][Path]^<FileName^> [HashFunction]
  echo(
  echo   FileName       Name of file to calculate hash value of.
  echo(
  echo   HashFunction   Hash function to use. Supported values are:
  echo                    %SupportedHashFunctions%
  echo(
exit /b 0
