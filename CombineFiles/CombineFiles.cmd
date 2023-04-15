@echo off & setlocal enabledelayedexpansion

set "InFile1=.\File1.txt"
set "InFile2=.\File2.txt"
set "OutFile=.\Merged.txt"

(for /f "tokens=1* delims=:" %%a in ('findstr /n "^" "%InFile1%"') do (
   set "line="
   set /p "line="
   echo.%%b
   echo.!line!
)) <"%InFile2%" >"%OutFile%"
