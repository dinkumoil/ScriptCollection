@echo off & setlocal

for /f "tokens=1-3 delims=:" %%a in ("%~2") do (
  start "" "%~1\notepad++.exe" -n%%c "%%a:%%b"
)
