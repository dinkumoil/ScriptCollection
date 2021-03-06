@echo off & setlocal

set "VBScript=%TEMP%\GetUptime.vbs"

> "%VBScript%" echo.Set objOS         = GetObject("winmgmts:\\.\root\cimv2:Win32_OperatingSystem=@")
>>"%VBScript%" echo.Set objDateTime   = CreateObject("WbemScripting.SWbemDateTime")
>>"%VBScript%" echo.
>>"%VBScript%" echo.objDateTime.Value = objOS.LastBootupTime
>>"%VBScript%" echo.datBootupTime     = objDateTime.GetVarDate(True)
>>"%VBScript%" echo.intUptimeSeconds  = DateDiff("s", datBootupTime, Now)
>>"%VBScript%" echo.
>>"%VBScript%" echo.intUptimeDays     = Fix(intUptimeSeconds / 86400)
>>"%VBScript%" echo.intUptimeHours    = Fix(intUptimeSeconds / 3600) Mod 24
>>"%VBScript%" echo.intUptimeMinutes  = Fix(intUptimeSeconds / 60) Mod 60
>>"%VBScript%" echo.intUptimeSeconds  = intUptimeSeconds Mod 60
>>"%VBScript%" echo.
>>"%VBScript%" echo.WScript.Echo intUptimeDays    ^& ";" ^&_
>>"%VBScript%" echo.             intUptimeHours   ^& ";" ^&_
>>"%VBScript%" echo.             intUptimeMinutes ^& ";" ^&_
>>"%VBScript%" echo.             intUptimeSeconds ^& ";" ^&_
>>"%VBScript%" echo.             datBootupTime

for /f "tokens=5 delims=;" %%a in ('cscript /nologo "%VBScript%"') do (
  echo Last bootup time: %%a
)

del "%VBScript%" > NUL
