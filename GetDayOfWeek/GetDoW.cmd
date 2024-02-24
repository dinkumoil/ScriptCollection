::*****************************************************************************
::*                                                                           *
::*                         GetDoW (Get Day of Week)                          *
::*                                                                           *
::*                         Autor: Andreas Heim, 2009                         *
::*                                                                           *
::*****************************************************************************


@ECHO OFF

SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION


CALL :CheckParam "%~1" "%~2" "%~3" "%~4"

IF ERRORLEVEL 4 GOTO ErrDate
IF ERRORLEVEL 3 GOTO ErrParUnk
IF ERRORLEVEL 2 GOTO ErrParMiss
IF ERRORLEVEL 1 GOTO Syntax

SET dd=%p1%
SET mm=%p2%
SET yy=%p3%
SET outc=%p5%


IF %mm% LEQ 2 (SET /A mm+=12 & SET /A yy-=1)
SET /A dow=((dd+(mm+1)*26/10+yy%%100+(yy%%100)/4+yy/400-2*(yy/100))%%7+6)%%7
IF %dow% EQU 0 SET /A dow=7


IF /I %outc%==gs (
   SET weekdays=Mo Di Mi Do Fr Sa So
) ELSE IF /I %outc%==gl (
   SET weekdays=Montag Dienstag Mittwoch Donnerstag Freitag Samstag Sonntag
) ELSE IF /I %outc%==es (
   SET weekdays=Mon Tue Wed Thu Fri Sat Sun
) ELSE IF /I %outc%==el (
   SET weekdays=Monday Tuesday Wednesday Thursday Friday Saturday Sunday
)

FOR /F "tokens=%dow%" %%i IN ("%weekdays%") DO SET dow=%%i


ECHO %dow%


ENDLOCAL

EXIT /B 0



:ErrParMiss
ECHO.
ECHO Parameter fehlt^^!
ECHO Rufen Sie getdow mit dem Parameter /? auf, um die Hilfe anzuzeigen.
ECHO.
EXIT /B 1

:ErrParUnk
ECHO.
ECHO Paramter unbekannt oder falsches Format^^!
ECHO Rufen Sie getdow mit dem Parameter /? auf, um die Hilfe anzuzeigen.
ECHO.
EXIT /B 1


:ErrDate
ECHO.
ECHO UngÅltiges Datum^^!
ECHO.
EXIT /B 1

:Syntax
ECHO.
ECHO Gibt den DoW (Tag der Woche) als numerischen Wert, englische oder deutsche
ECHO AbkÅrzung oder mit vollstÑndigem englischem oder deutschem Namen aus.
ECHO Die Berechnung erfolgt nach dem Gregorianischen Kalender, gÅltig ab
ECHO Freitag, dem 15.10.1582.
ECHO.
ECHO Syntax: GetDoW Datum^|/i [/m:maske] [/n^|/gs^|/gl^|/es^|/el]
ECHO.
ECHO.
ECHO Parameter: 1. Datum
ECHO               Tag und Monat 2-stellig mit fÅhrender 0, Jahr 4-stellig
ECHO               Pflichtparameter (au·er Parameter /i ist angegeben)
ECHO               Das Trennzeichen mu· das gleiche wie in der Maske sein.
ECHO.
ECHO            2. /m:maske
ECHO               Maske, die definiert, wo Tag, Monat, Jahr im Åbergebenen
ECHO               Datum stehen (z.B. YYYY-MM-DD). ZulÑssige Trennzeichen
ECHO               sind -./
ECHO               Default DD.MM.YYYY
ECHO.
ECHO            3. Ausgabesteuerung des Ergebnisses
ECHO                 /n  --^> numerische Ausgabe (Montag=1...Sonntag=7)
ECHO                 /gs --^> Wochentag wird mit deutscher AbkÅrzung
ECHO                         ausgegeben
ECHO                 /gl --^> Wochentag wird mit vollstÑndigem deutschem
ECHO                         Namen ausgegeben
ECHO                 /es --^> Wochentag wird mit englischer AbkÅrzung
ECHO                         ausgegeben
ECHO                 /el --^> Wochentag wird mit vollstÑndigem englischem
ECHO                         Namen ausgegeben
ECHO               Default /n
ECHO.
ECHO            4. /i
ECHO               Falls angegeben und kein Datum Åbergeben wurde, wird das
ECHO               Datum Åber STDIN eingelesen.
ECHO               Dadurch ist z.B. folgendes mîglich:
ECHO                 ECHO %%DATE%% ^| getdow /i /gl (deutsches Windows)
ECHO                   oder
ECHO                 ECHO %%DATE%% ^| getdow /i /el /m:YYYY-MM-DD (US-Windows)
ECHO                   oder
ECHO                 ECHO. ^| SET /P=Heute ist ^& ECHO %%DATE%% ^| getdow /i /gl
ECHO                 (deutsches Windows)
ECHO.
ECHO Die Reihenfolge der Parameter ist beliebig.
ECHO Bei einem Fehler wird ERRORLEVEL auf 1 gesetzt, sonst auf 0.
ECHO Algorithmus nach Christian Zeller (Zellers Kongruenz).
ECHO.
EXIT /B 0





::*************************************************************************
::*                                                                       *
::*                         Unterprogramm-Sektion                         *
::*                                                                       *
::*************************************************************************

:CheckParam
SET p1=&SET p2=&SET p3=&SET p4=&SET p5=&SET p6=&SET pd=&SET pn=&SET p4ok=&SET p5ok=&SET dsep=


IF "%~1"=="/?" EXIT /B 1
IF "%~1"=="" EXIT /B 2


FOR %%i IN (%1 %2 %3 %4) DO (
   SET pn=%%~i

   IF DEFINED pn (
      IF /I "!pn:~0,3!"=="/m:" (
         SET p4=!pn:~3!
      ) ELSE IF /I "!pn!"=="/i" (
          SET p6=i
      ) ELSE IF "!pn:~0,1!"=="/" (
          SET p5=!pn:~1,2!
      ) ELSE (
          SET pd=!pn!
      )
   )
)


IF NOT DEFINED pd (
   IF /I "%p6%"=="i" (
      SET /P pd=
   ) ELSE (
      EXIT /B 2
   )
)


IF NOT DEFINED p4 (
   SET p4=DD.MM.YYYY
   SET dsep=.
) ELSE (
   FOR /F "tokens=1 delims=DMYdmy" %%i IN ("%p4%") DO (
      SET dsep=%%i
      SET dsep=!dsep:~0,1!
   )
   IF NOT DEFINED dsep EXIT /B 3

   FOR /F "tokens=1 delims=.-/" %%i IN ("!dsep!") DO EXIT /B 3

   FOR %%i IN (DD!dsep!MM!dsep!YYYY MM!dsep!DD!dsep!YYYY YYYY!dsep!MM!dsep!DD YYYY!dsep!DD!dsep!MM DD!dsep!YYYY!dsep!MM MM!dsep!YYYY!dsep!DD) DO (
      IF /I "%p4%"=="%%i" SET p4ok=1
   )
   IF NOT DEFINED p4ok EXIT /B 3
)


IF NOT DEFINED p5 (
   SET p5=n
) ELSE (
   FOR %%i IN (n gs gl es el) DO (
      IF /I "%p5%"=="%%i" SET p5ok=1
   )
   IF NOT DEFINED p5ok EXIT /B 3
)


SET pd=%pd:~0,10%
IF "%pd:~9,1%"=="" EXIT /B 3

FOR /F "tokens=1-3 delims=%dsep%" %%i IN ("%p4%") DO (
   IF /I "%%i"=="DD" SET p1=%pd:~0,2%
   IF /I "%%i"=="MM" SET p2=%pd:~0,2%
   IF /I "%%i"=="YYYY" SET p3=%pd:~0,4%

   IF /I NOT "%%i"=="YYYY" (
      IF /I "%%j"=="DD" SET p1=%pd:~3,2%
      IF /I "%%j"=="MM" SET p2=%pd:~3,2%
      IF /I "%%j"=="YYYY" SET p3=%pd:~3,4%
   ) ELSE (
      IF /I "%%j"=="DD" SET p1=%pd:~5,2%
      IF /I "%%j"=="MM" SET p2=%pd:~5,2%
   )

   IF /I "%%k"=="YYYY" (
      SET p3=%pd:~6,4%
   ) ELSE (
      IF /I "%%k"=="DD" SET p1=%pd:~8,2%
      IF /I "%%k"=="MM" SET p2=%pd:~8,2%
   )
)

FOR /F "tokens=1 delims=0123456789" %%i IN ("%p1%%p2%%p3%") DO EXIT /B 3


FOR %%i IN (%p1% %p2% %p3%) DO (
   IF %%i LEQ 0 EXIT /B 4
)

IF %p1% GTR 31 EXIT /B 4
IF %p2% GTR 12 EXIT /B 4

IF %p2%==02 (
   IF %p1% GTR 29 EXIT /B 4

   SET /A isleapyr=p3%%4

   IF !isleapyr! GTR 0 (
       IF %p1% GTR 28 EXIT /B 4
   ) ELSE (
      SET /A isleapyr=p3%%100

      IF !isleapyr! EQU 0 (
         SET /A isleapyr=p3%%400

         IF !isleapyr! GTR 0 (
            IF %p1% GTR 28 EXIT /B 4
         )
      )
   )
) ELSE (
   FOR %%i IN (04 06 09 11) DO (
      IF %p2%==%%i (
         IF %p1% GTR 30 EXIT /B 4
      )
   )
)


EXIT /B 0
