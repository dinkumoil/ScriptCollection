@echo off

setlocal disabledelayedexpansion
pushd "%~dp0"


set "ResultFile=%TEMP%\FselSelectedObject.txt"

::FileSelector aufrufen
::Wenn keine Datei ausgewaehlt wurde, Ende
::NICHT if defined selectedFile (...) verwenden, da man sonst
::ENABLEDELAYEDEXPANSION benoetigt, was Probleme mit in selectedFile
::enthaltenen Ausrufezeichen macht.
rem call FSel E:\ /m *.* /fdn
rem if not defined selectedObject goto :ScriptExit

call FSel E:\ /m *.* /fdn /r "%ResultFile%"
if not exist "%ResultFile%" goto :ScriptExit
<"%ResultFile%" set /p "selectedObject="


::Um den Datei-/Verzeichnisnamen als Parameter fuer ein Programm zu
::verwenden, einfach so uebernehmen.
set "selectedObjectParam=%selectedObject%"

::Zum Ausgeben des Datei-/Verzeichnisnamens kritische Zeichen escapen.
::Die Reihenfolge ist wichtig!
set "selectedObject=%selectedObject:^=^^%"
set "selectedObject=%selectedObject:&=^&%"


::Datei-/Verzeichnisnamen nur ausgeben
cls
echo.
echo Ausgew„hltes Objekt: %selectedObject%

::Handelt es sich um eine Datei?
>NUL 2>&1 dir /b "%selectedObjectParam%\*.*"

::Wenn es eine Datei ist, Namen als Parameter verwenden
if %errorlevel% neq 0 (
  "%selectedObjectParam%"
) else (
  echo. & pause
)


:ScriptEnd
del "%ResultFile%" 2>NUL

:ScriptExit
popd
endlocal
exit /b
