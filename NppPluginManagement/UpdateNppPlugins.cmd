@echo off & setlocal

set "BasePath1=C:\Program Files (x86)\Notepad++"

pushd "%~dp0"

for %%a in ("%BasePath1%") do (
  for /f "delims=" %%b in ('dir /s /b /a:-d "%%~a\notepad++.exe" 2^>NUL ^| findstr /irev /c:"64.*\\notepad++.exe"') do (
    call :ProcessDir x86 "%%~dpb"
  )
)

popd
exit /b 0


:ProcessDir
  set "NppPath=%~2"
  set "NppPath=%NppPath:~0,-1%"

  call ".\NppPluginManagement_%~1.cmd" /a %~1 /t update /n "%NppPath%"
exit /b 0
