@echo off & setlocal

set "DestFolder=C:\Program Files\Delete ADS"
set "FilesToCopy="streams.exe";"DeleteADS.ico""
set "RegFileToExecute=DeleteADS.reg"

pushd "%~dp0"

set "VBScript=%TEMP%\Elevate.vbs"
call :WriteVBScript "%CD%" "%~nx0"

if /i "%~1" neq "/elevated" (
  start "" wscript /nologo "%VBScript%"
  exit /b 0
)

md "%DestFolder%" 2>NUL

for %%a in (%FilesToCopy%) do (
  copy "%%~a" "%DestFolder%" > NUL
)

start "" regedit /s "%RegFileToExecute%"

del "%VBScript%" > NUL

popd
exit /b


:WriteVBScript
  chcp 1252 > NUL
  > "%VBScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%VBScript%" echo.Set objFSO   = CreateObject("Scripting.FileSystemObject")
  >>"%VBScript%" echo.
  >>"%VBScript%" echo.strApplication = "cmd.exe"
  >>"%VBScript%" echo.strArguments   = "/c """"" ^& objFSO.BuildPath("%~1", "%~2") ^& """ /elevated"""
  >>"%VBScript%" echo.
  >>"%VBScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 0
  chcp 850 > NUL
exit /b 0
