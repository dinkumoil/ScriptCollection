::##############################################################################
:: Main script
::##############################################################################

@echo off & setlocal


::******************************************************************************
:: Basic configuration
::******************************************************************************

::Path for temporary VBScript
set "VBScript=%Temp%\NppMigrationElevate.vbs"


::******************************************************************************
:: Startup sequence
::******************************************************************************

::Parse command line and jump to other code section according to return value
call :ParseCommandLine %*

if ERRORLEVEL 2 goto :DoJob
if ERRORLEVEL 1 goto :ErrorExit


::Display intro screen if script has been started normally
call :DisplayIntroScreen || goto :UserExit


::******************************************************************************
:: Search installation directory of local or hybrid installation of Notepad++
::******************************************************************************

::Set default Notepad++ installation path for local and hybrid installations
set "Npp32BitDir=%ProgramFiles(x86)%\Notepad++"
set "NppNativeDir=%ProgramFiles%\Notepad++"


::Search installation directory of Notepad++
set "NppDir="

for %%a in ("%Npp32BitDir%" "%NppNativeDir%") do (
  if exist "%%~a\notepad++.exe" set "NppDir=%%~a"
)

::Nothing found, maybe we have a portable installation
if "%NppDir%" equ "" goto :PromptForPortableInstallation1


::******************************************************************************
:: Set old and new plugins path according to installation type "Local",
:: "Localv7.6", "Localv7.6.1" or "Hybrid"
::******************************************************************************

REM Check for local installation
if not exist "%NppDir%\allowAppDataPlugins.xml" (
  dir /b /a:-d "%NppDir%\plugins\*.dll" 1>NUL 2>NUL && (
    REM Set plugins path for local installations of Notepad++ prior v7.6
    set "PluginsSrc=%NppDir%\plugins"
    set "PluginsDst=%NppDir%\plugins"
    set "InstType=Local"
  ) || (
    dir /s /b /a:-d "%LocalAppData%\Notepad++\plugins\*.dll" 1>NUL 2>NUL && (
      REM Set plugins path for local installation of Notepad++ v7.6
      set "PluginsSrc=%LocalAppData%\Notepad++\plugins"
      set "PluginsDst=%NppDir%\plugins"
      set "InstType=Localv7.6"
    ) || (
      dir /s /b /a:-d "%ProgramData%\Notepad++\plugins\*.dll" 1>NUL 2>NUL && (
        REM Set plugins path for local installation of Notepad++ v7.6.1 and v7.6.2
        set "PluginsSrc=%ProgramData%\Notepad++\plugins"
        set "PluginsDst=%NppDir%\plugins"
        set "InstType=Localv7.6.1"
      ) || (
        REM Show error message and quit when plugins directory could not been found
        echo.
        echo There seems to be a local installation of Notepad++ on this machine but its
        echo plugins directory could not been found.
        echo.
        goto :ErrorExit
      )
    )
  )

REM Assume hybrid installation
) else (
  REM Set plugins path for hybrid installations
  set "PluginsSrc=%AppData%\Notepad++\plugins"
  set "PluginsDst=%NppDir%\plugins"
  set "InstType=Hybrid"
)


::******************************************************************************
:: Show menu where user can choose to migrate the local or hybrid installation
:: found above or to migrate an additional portable installation
::******************************************************************************

:SelectMigrationType
echo.
echo.
echo There has been found a Notepad++ installation at
echo.
echo   "%NppDir%"
echo.
echo Please select an option:
echo.
echo   a) Migrate the plugins of this installation
echo   b) Migrate the plugins of an additional portable installation
echo.
echo Press A+ENTER to migrate the installation found and B+ENTER to migrate a
echo portable installation. Press only ENTER to exit.
echo.

set "KbInput="
set /p "KbInput=Your choice: "
echo.

::In case user pressed only ENTER, terminate script
if "%KbInput%" equ "" (
  goto :UserExit
)

::In case user choosed to migrate the local or hybrid installation found above,
::restart script with elevated user rights
if /i "%KbInput%" equ "A" (
  call :RestartElevated "%PluginsSrc%" "%PluginsDst%" "%InstType%"
  goto :RestartExit
)

::In case user choosed to migrate an additional local installation, jump to input
::query for its path
if /i "%KbInput%" equ "B" (
  goto :PromptForPortableInstallation2
)

::In all other cases show menu again
goto :SelectMigrationType


::******************************************************************************
:: Show input query to retrieve path of a portable installation to migrate
::******************************************************************************

:PromptForPortableInstallation1
echo.
echo.
echo There has been no local or hybrid installation of Notepad++ found on this
echo machine. If you run a portable installation of Notepad++ please enter the
echo complete path to its directory (without quotes). If you want to exit press
echo only ENTER.
echo.

goto :InputPortableInstallation

:PromptForPortableInstallation2
echo.
echo.
echo Please enter the complete path to the directory of your portable Notepad++
echo installation (without quotes). If you want to exit press only ENTER.
echo.

:InputPortableInstallation
set "InstType=Portable"

set "NppDir="
set /p "NppDir=Enter path: "
echo.

::In case user pressed only ENTER, terminate script
if "%NppDir%" equ "" (
  goto :UserExit
)

::In case user provided a non-existent directory or the directory doesn't contain
::a portable installation, terminate script
if not exist "%NppDir%\doLocalConf.xml" (
  echo This directory doesn't look like the home of a portable installation,
  echo aborting plugin migration.
  echo.
  goto :ErrorExit
)


::******************************************************************************
:: Set old and new plugins path according to installation type "Portable"
::******************************************************************************

set "PluginsSrc=%NppDir%\plugins"
set "PluginsDst=%NppDir%\plugins"


::******************************************************************************
:: Do some error checkings
::******************************************************************************

:DoJob
::Check if source and destination directories exist
dir /b /a:d "%PluginsSrc%" 1>NUL 2>NUL || (
  echo Directory "%PluginsSrc%" not found, aborting plugin migration.
  echo.
  goto :ErrorExit
)

dir /b /a:d "%PluginsDst%" 1>NUL 2>NUL || (
  echo Directory "%PluginsDst%" not found, aborting plugin migration.
  echo.
  goto :ErrorExit
)

echo.
echo.


::******************************************************************************
:: Move plugins and companion files to new plugins location
::******************************************************************************

::In case of a portable, hybrid or prior v7.6 local installation we have to create
::the new plugin directory structure
if /i "%InstType:~0,9%" neq "Localv7.6" (
  for %%a in ("%PluginsSrc%\*.dll") do (
    echo.
    echo ===============================================================================
    echo Processing plugin %%~na
    echo ===============================================================================

    REM Process directory under "plugins" related to current plugin
    REM If it already exists move it one level down in directory hierarchy
    dir /b /a:d "%PluginsDst%\%%~na" 1>NUL 2>NUL && (
      echo.
      echo ---- Move directory ----
      echo "%PluginsDst%\%%~na"
      echo to
      echo "%PluginsDst%\%%~na\%%~na"

      move "%PluginsDst%\%%~na" "%PluginsDst%\%%~na_MigTemp" 1>NUL
      md "%PluginsDst%\%%~na\%%~na" 1>NUL
      xcopy /eikqy "%PluginsDst%\%%~na_MigTemp\*.*" "%PluginsDst%\%%~na\%%~na" 1>NUL
      rd /s /q "%PluginsDst%\%%~na_MigTemp" 1>NUL
    ) || (
      dir /b /a:d "%PluginsSrc%\%%~na" 1>NUL 2>NUL && (
        echo.
        echo ---- Move directory ----
        echo "%PluginsSrc%\%%~na"
        echo to
        echo "%PluginsDst%\%%~na\%%~na"

        md "%PluginsDst%\%%~na\%%~na" 1>NUL
        xcopy /eikqy "%PluginsSrc%\%%~na\*.*" "%PluginsDst%\%%~na\%%~na" 1>NUL
        rd /s /q "%PluginsSrc%\%%~na" 1>NUL
      )
    )

    REM Process directory under "plugins\doc" related to current plugin
    dir /b /a:d "%PluginsSrc%\doc\%%~na" 1>NUL 2>NUL && (
      echo.
      echo ---- Move directory ----
      echo "%PluginsSrc%\doc\%%~na"
      echo to
      echo "%PluginsDst%\%%~na\doc\%%~na"

      md "%PluginsDst%\%%~na\doc\%%~na" 1>NUL 2>NUL
      xcopy /eikqy "%PluginsSrc%\doc\%%~na\*.*" "%PluginsDst%\%%~na\doc\%%~na" 1>NUL
      rd /s /q "%PluginsSrc%\doc\%%~na" 1>NUL
    )

    REM Process files under "plugins\doc" related to current plugin
    dir /b /a:-d "%PluginsSrc%\doc\%%~na*.*" 1>NUL 2>NUL && (
      echo.
      echo ---- Move files ----
      echo "%PluginsSrc%\doc\%%~na*.*"
      echo to
      echo "%PluginsDst%\%%~na\doc"

      md "%PluginsDst%\%%~na\doc" 1>NUL 2>NUL
      move "%PluginsSrc%\doc\%%~na*.*" "%PluginsDst%\%%~na\doc" 1>NUL
    )

    REM Process plugin DLL file
    echo.
    echo ---- Move file ----
    echo "%%~a"
    echo to
    echo "%PluginsDst%\%%~na"

    md "%PluginsDst%\%%~na" 1>NUL 2>NUL
    move "%%~a" "%PluginsDst%\%%~na" 1>NUL

    echo.
    echo.
  )

REM In case of a v7.6 or above local installation we only have to move the
REM already existing plugin directory structure
) else (
  for /d %%a in ("%PluginsSrc%\*.*") do (
    echo.
    echo ===============================================================================
    echo Processing plugin %%~nxa
    echo ===============================================================================

    echo.
    echo ---- Move directory ----
    echo "%PluginsSrc%\%%~nxa"
    echo to
    echo "%PluginsDst%\%%~nxa"

    md "%PluginsDst%\%%~nxa" 1>NUL
    xcopy /eikqy "%PluginsSrc%\%%~nxa\*.*" "%PluginsDst%\%%~nxa" 1>NUL
    rd /s /q "%PluginsSrc%\%%~nxa" 1>NUL

    echo.
    echo.
  )
)


::******************************************************************************
:: Final message
::******************************************************************************

echo Plugin migration finished
echo.
set /a ExitCode=0
goto :CleanUp


::******************************************************************************
:: Exit points for different cases
::******************************************************************************

:UserExit
echo Aborted by user 1>&2
echo.
set /a ExitCode=1
goto :CleanUp


:ErrorExit
echo Something went wrong 1>&2
echo.
set /a ExitCode=3
goto :CleanUp


:RestartExit
set /a ExitCode=2
goto :Terminate


::******************************************************************************
:: Cleanup and termination
::******************************************************************************

:CleanUp
del "%VBScript%" 1>NUL 2>NUL
goto :Quit


:Quit
pause

:Terminate
exit /b %ExitCode%




::##############################################################################
:: Subroutines
::##############################################################################

:DisplayIntroScreen
  cls

  echo *******************************************************************************
  echo.
  echo   This script moves the DLL and companion files of all Notepad++ plugins
  echo   from their current loction to the new plugin folder of Notepad++ v7.6.3
  echo.
  echo   Please note
  echo   -----------
  echo   This script uses a generic algorithm to do this task. There may be files
  echo   and/or directories at the current location which you have to move manually
  echo   to the new location. There may also be files and/or directories which get
  echo   moved by the script but which would have to stay at their current location.
  echo.
  echo   This script is distributed in the hope that it will be useful, but WITHOUT
  echo   ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  echo   FOR A PARTICULAR PURPOSE. The author is not responsible for any damage caused
  echo   by running this script. It is highly recommended to back up the directories
  echo   processed by this script and any critical data you may have stored on your
  echo.  hard disk.
  echo.
  echo *******************************************************************************
  echo.
  echo.

  set "KbInput="
  set /p "KbInput=Press ENTER to continue and E+ENTER to exit: "
  echo.

  if /i "%KbInput%" equ "E" exit /b 1
exit /b 0



:ParseCommandLine
  set "PluginsSrc=%~1"
  set "PluginsDst=%~2"
  set "InstType=%~3"

  if "%PluginsSrc%" equ "" exit /b 0
  if "%PluginsDst%" equ "" exit /b 1
  if "%InstType%" equ "" exit /b 1
exit /b 2



:RestartElevated
  chcp 1252 > NUL
  > "%VBScript%" echo.Set objShell = CreateObject("Shell.Application")
  >>"%VBScript%" echo.Set objFSO   = CreateObject("Scripting.FileSystemObject")
  >>"%VBScript%" echo.
  >>"%VBScript%" echo.strApplication = "cmd.exe"
  >>"%VBScript%" echo.strArguments   = "/c """"" ^& objFSO.BuildPath("%~dp0", "%~nx0") ^& """ ""%~1"" ""%~2"" ""%~3"""""
  >>"%VBScript%" echo.
  >>"%VBScript%" echo.objShell.ShellExecute strApplication, strArguments, "", "runas", 1
  chcp 850 > NUL

  cscript /nologo "%VBScript%"
exit /b 0
