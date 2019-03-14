@echo off


::Funktion FileSelector
::
::Funktion um eine Dateiauswahlbox darzustellen. Zur Darstellung wird die
::Funktion ShowMenu verwendet. Der Pfad und der Inhalt des gerade bearbeiteten
::Verzeichnisses wird in eine Datei geschrieben, die von ShowMenu ausgelesen
::und dargestellt wird. Ein Verzeichnis oder eine Datei kann anhand der
::Zeilennummer ausgewaehlt werden. Verzeichnisse werden durch das Zeichen
::markiert. Um in ein angezeigtes Verzeichnis zu wechseln, nur zugehoerige
::Nummer eingeben. Um das Verzeichnis selbst auszuwaehlen, CX (X=Nummer des
::Eintrags) eingeben. Das Wechseln des uebergebenen Verzeichnisses kann durch
::den Parameter /s unterbunden werden. Der Text der anzuzeigenden Eingabe-
::aufforderung kann optional mit dem Parameter /i uebergeben werden. Durch die
::Uebergabe eines Modus-Parameters kann festgelegt werden, ob nur Dateien, nur
::Verzeichnisse oder beides ausgewaehlt werden koennen und ob diese bereits
::existieren muessen oder auch neu angelegt werden koennen. Neue Verzeichnisse
::und Dateien koennen ueber einen Dialog angelegt werden, der ueber die Taste
::T (Tools) erreichbar ist. Dort koennen auch Dateien und Verzeichnisse
::geloescht werden.
::
::
::Aufruf   : call :FileSelector Pfad /m Maske [/i Eingabeaufforderung]
::                              [/f | /fd | /fn | /fdn | /d | /dn]
::                              [/r Ergebnisdatei] [/s]
::
::Parameter: Pfad
::             Der Pfad zum Verzeichnis, das angezeigt werden soll. Wenn als
::             leerer Parameter uebergeben, wird der Arbeitsplatz angezeigt.
::
::           Maske
::             Die Maske fuer die anzuzeigenden Dateien. Verwendung von
::             Wildcards ist moeglich.
::
::           Eingabeaufforderung
::             Die Eingabeaufforderung, die angezeigt werden soll. Standard
::             ist "Welches Verzeichnis/Welche Datei mîchten Sie îffnen? "
::
::           Modus
::             /f   - Nur bereits existierende Dateien koennen ausgewaehlt
::                    werden. Das ist der Standardmodus.
::             /fd  - Nur bereits existierende Dateien und Verzeichnisse
::                    koennen ausgewaehlt werden.
::             /fn  - Nur Dateien koennen ausgewaehlt werden. Zum Anlegen von
::                    neuen Dateien steht der Steuerbefehl T (Tools) zur
::                    Verfuegung. In diesem Sub-Menue koennen neue leere
::                    Dateien angelegt und danach ausgewaehlt werden. Neue
::                    leere Verzeichnisse koennen auch angelegt, aber danach
::                    nicht ausgewaehlt werden.
::             /fdn - Es koennen Dateien und Verzeichnisse ausgewaehlt werden.
::                    Zum Anlegen von neuen leeren Dateien und Verzeichnissen
::                    steht der Steuerbefehl T (Tools) zur Verfuegung. Die
::                    ueber dieses Sub-Menue angelegten leeren Dateien und
::                    Verzeichnisse koennen nach dem Schliessen des Submenues
::                    ausgewaehlt werden.
::             /d   - Es koennen nur bereits existierende Verzeichnisse aus-
::                    gewaehlt werden.
::             /dn  - Es koennen nur Verzeichnisse ausgewaehlt werden. Zum
::                    Anlegen von neuen Verzeichnissen steht der Steuerbefehl
::                    T (Tools) zur Verfuegung. In diesem Sub-Menue koennen
::                    neue leere Verzeichnisse angelegt und danach ausgewaehlt
::                    werden. Neue leere Dateien koennen auch angelegt, aber
::                    danach nicht ausgewaehlt werden.
::
::           Ergebnisdatei
::             Optional kann Name und Pfad zu einer Datei uebergeben werden,
::             in die das Ergebnis der Auswahl geschrieben werden soll.
::
::           Statischer Pfad
::             Durch Uebergabe des Parameters /s kann das Wechseln des
::             Verzeichnisses verboten werden.
::
::Rueckgabe: In der Variablen selectedObject wird der Name inkl. Pfad der
::           ausgewaehlten Datei/des ausgewaehlten Verzeichnisses zurueck-
::           gegeben. Wurde der Dialog abgebrochen, kein Pfad oder keine
::           Maske angegeben, ist selectedFile nicht definiert. Falls mit
::           dem Parameter /r eine Ergebnisdatei uebergeben wurde, enthaelt
::           diese ebenfalls die ausgewaehlte Datei/das ausgewaehlte
::           Verzeichnis. Wurde der Dialog abgebrochen, existiert die Datei
::           nicht.


:FileSelector
setlocal disabledelayedexpansion

cls


set "objectPath="
set "objectMask="
set "inputStr=Welches Verzeichnis/Welche Datei mîchten Sie îffnen? "
set "keepDir=0"

set "showTools="
set "allowFiles=1"
set "allowFilesDirs=0"
set "allowNewFiles=0"
set "allowNewFilesDirs=0"
set "allowOnlyExistingDirs=0"
set "allowOnlyDirs=0"

set "selectedObject="
set "outputFile="


::Parameter uebernehmen
:FileSelectorGetParamsLoop
  set "param=%~1"

  if /i "%param%" equ "/m" (
    set "objectMask=%~2" & shift
  ) else if /i "%param%" equ "/i" (
    set "inputStr=%~2" & shift
  ) else if /i "%param%" equ "/s" (
    set "keepDir=1"
  ) else if "%param%" equ "/f" (
    set "allowFiles=1" & set "allowFilesDirs=0" & set "allowNewFiles=0" & set "allowNewFilesDirs=0" & set "allowOnlyDirs=0" & set "allowOnlyExistingDirs=0"
  ) else if "%param%" equ "/fd" (
    set "allowFilesDirs=1" & set "allowFiles=0" & set "allowNewFiles=0" & set "allowNewFilesDirs=0" & set "allowOnlyDirs=0" & set "allowOnlyExistingDirs=0"
  ) else if "%param%" equ "/fn" (
    set "allowNewFiles=1" & set "allowFiles=0" & set "allowFilesDirs=0" & set "allowNewFilesDirs=0" & set "allowOnlyDirs=0" & set "allowOnlyExistingDirs=0"
  ) else if "%param%" equ "/fdn" (
    set "allowNewFilesDirs=1" & set "allowFiles=0" & set "allowFilesDirs=0" & set "allowNewFiles=0" & set "allowOnlyDirs=0" & set "allowOnlyExistingDirs=0"
  ) else if "%param%" equ "/d" (
    set "allowOnlyExistingDirs=1" & set "allowFiles=0" & set "allowFilesDirs=0" & set "allowNewFiles=0" & set "allowNewFilesDirs=0" & set "allowOnlyDirs=0"
  ) else if "%param%" equ "/dn" (
    set "allowOnlyDirs=1" & set "allowFiles=0" & set "allowFilesDirs=0" & set "allowNewFiles=0" & set "allowNewFilesDirs=0" & set "allowOnlyExistingDirs=0"
  ) else if /i "%param%" equ "/r" (
    set "outputFile=%~2" & shift
  ) else if "%param:~0,1%" neq "/" if not defined objectPath (
    set "objectPath=%~f1"
  )

  shift
if "%~1" neq "" goto :FileSelectorGetParamsLoop


::Bei fehlendem Parameter objectMask Ruecksprung.
::Bei leerem Parameter objectPath wird der Arbeitsplatz angezeigt.
::Falls die Parameter Carret-Zeichen (^) enthalten, wurden sie bei der
::Uebergabe aus einem Batch-Skript vom Kommando-Interpreter verdoppelt.
::Das muss wieder rueckgaengig gemacht werden.
if not defined objectMask goto :FileSelectorExit
set "objectMask=%objectMask:^^=^%"

if not defined objectPath goto :FileSelectorChkOutputFile
set "objectPath=%objectPath:^^=^%"

:FileSelectorChkOutputFile
if not defined outputFile goto :FileSelectorSetMenuFile
set "outputFile=%outputFile:^^=^%"


:FileSelectorSetMenuFile
::Menue-Datei festlegen. Durch Anhaengen einer Zufallszahl koennen mehrere
::Instanzen von FileSelector nacheinander gestartet werden.
set "menuFile=%TEMP%\FileSelectorMenu_%random%.txt"


::Maximale Laenge der Titelzeile (und damit auch der Breite der
::Fileselectorbox) festlegen
set "maxTitleLen=72"


::Parameter fuer ShowMenu, um den Zugang zum Tools-Menue freizuschalten
set "toolsStr=/u T "T=Tools" FileSelectorTools"


::Vorhandene Laufwerke ermitteln
::Dazu ein temporaeres VBS-Skript benutzen
set "vbsCode=%TEMP%\GetDrives.vbs"

> "%vbsCode%" echo Dim fso, dc, d, s
>>"%vbsCode%" echo Set fso = CreateObject("Scripting.FileSystemObject")
>>"%vbsCode%" echo Set dc = fso.Drives
>>"%vbsCode%" echo For Each d in dc
>>"%vbsCode%" echo   s = s ^& d.DriveLetter ^& ": "
>>"%vbsCode%" echo Next
>>"%vbsCode%" echo WScript.Echo s

::Skript ausfuehren, um die Laufwerksbuchstaben zu ermitteln
for /f "delims=" %%d in ('cscript /nologo "%vbsCode%"') do (
  set "availDrives=%%d"
)

::Skript loeschen
del "%vbsCode%" 2>NUL


::Die Hauptschleife von FileSelector
::Wird solange ausgefuehrt, bis eine Datei/ein Verzeichnis ausgewaehlt
::oder der Dialog abgebrochen wurde.

:FileSelectorSelectLoop
  ::Rueckgabevariable loeschen
  set "selectedObject="

  ::Marker fuer spzielle Auswahl loeschen, falls die spezielle Auswahl mit
  ::einem Dateinamen ausgefuehrt wurde.
  set "dirSelected=0"


  ::Wenn objectPath definiert ist, ist ein Pfad darin enthalten.
  ::Falls es nicht definiert ist, wurde zum Arbeitsplatz (der Laufwerksauswahl)
  ::gewechselt und die Liste der vorhandenen Laufwerke muss angezeigt werden.
  if defined objectPath goto :FileSelectorProcessDir


  ::Statusmeldung ausgeben
  <NUL set /p "=Lese Laufwerksliste ein..."


  ::Menue mit dem Arbeitsplatz (der Laufwerksauswahl) generieren.
  ::Titelzeile schreiben
  > "%menuFile%" (<NUL set /p "=:::Arbeitsplatz" & echo.)

  ::Menueeintraege schreiben
  >>"%menuFile%" (for %%d in (%availDrives%) do (<NUL set /p "=::: %%d" & echo.))

  ::Eine Leerzeile an die Menue-Datei anhaengen, damit die Liste der
  ::Menueeintraege vollstaendig dargestellt wird.
  >>"%menuFile%" echo :::


  ::Bei der Laufwerksauswahl kein Tools-Menue anzeigen
  set "showTools="


  ::Menue anzeigen
  goto :FileSelectorShowMenu


  :FileSelectorProcessDir
  ::Statusmeldung ausgeben
  <NUL set /p "=Lese Verzeichnis ein..."


  ::Wenn vom Arbeitsplatz (der Laufwerksauswahl) auf ein Laufwerk gewechselt
  ::wurde, enthaelt objectPath als erstes Zeichen einen Backslash, dann den
  ::Laufwerksbuchstaben und dann den Doppelpunkt. Der fuehrende Backslash
  ::muss beseitigt werden.
  if "%objectPath:~0,1%" equ "\" (
    set "objectPath=%objectPath:~1%"
  )


  ::Falls das letzte Zeichen von objectPath ein Backslash ist, diesen entfernen.
  ::Das ist wichtig, falls aus einem Unterverzeichnis eine Ebene nach oben
  ::gewechselt wurde, dann ist das letzte Zeichen von objectPath ein Backslash.
  ::Falls objectPath ein Wurzelverzeichnis ist, wird spaeter beim Verzeichnislisting
  ::eine Hilfsvariable verwendet, an die dann wieder ein Backslash angehaengt wird.
  if "%objectPath:~-1%" equ "\" (
    set "objectPath=%objectPath:~0,-1%"
  )


  ::Suchmaske und Titelzeile initialisieren
  set "searchMask=%objectPath%\%objectMask%"
  set "fselTitle=%searchMask%"


  ::Evtl. die Laenge der Titelzeile kuerzen
  ::Beim Wurzelverzeichnis auf keinen Fall noetig
  set "fselTitleDrv=%fselTitle:~0,3%"
  set "fselTitlePath=%fselTitle:~3%"
  if not defined fselTitlePath goto :FileSelectorWriteMenu

  ::Korrekturfaktor auf Startwert setzen => es wurde noch nicht gekuerzt
  set "correction=0"

  :FileSelectorTrimTitleLoop
    ::Laenge von fselTitlePath ermitteln
    >"%TEMP%\strlen.txt" <NUL (set /p "=<%fselTitlePath%"&echo.&set /p "=<"&echo.)
    for /f "delims=:" %%l in ('findstr /o /b "<" "%TEMP%\strlen.txt"') do set /a "titleLen=%%l+correction"

    ::Wenn die Laenge jetzt stimmt, aber vorher schon einmal gekuerzt wurde,
    ::die gekuerzte Titelzeile zusammenbauen. Wenn die Zeile noch nicht
    ::gekuertzt werden musste, normal weitermachen.
    if %titleLen% leq %maxTitleLen% (
      del "%TEMP%\strlen.txt" 2>NUL
      if %correction% neq 0 goto :FileSelectorTrimTitleLoopEnd
      goto :FileSelectorWriteMenu
    )

    ::Alles vor dem ersten Backslash, inkl. dem Backslash selbst, loeschen
    set "fselTitlePath=%fselTitlePath:*\=%"

    ::Korrekturfaktor aendern => es wurde gekuerzt
    ::und neuer Schleifendurchlauf
    set "correction=4"
    goto :FileSelectorTrimTitleLoop
  :FileSelectorTrimTitleLoopEnd

  ::Die gekuerzte Titelzeile zusammenbauen
  set "fselTitle=%fselTitleDrv%...\%fselTitlePath%"


  :FileSelectorWriteMenu
  ::Menue generieren
  ::Titelzeile schreiben
  >"%menuFile%" (<NUL set /p "=:::%fselTitle%" & echo.)

  ::Standardeintrag zum Wechsel ins Elternverzeichnis erzeugen.
  >>"%menuFile%" echo ::: ..

  ::Verzeichnis auslesen und Menueeintraege in Datei schreiben.
  ::Eintraege fuer Unterverzeichnisse erzeugen.
  ::Wenn ein Wurzelverzeichnis gelistet werden soll,
  ::objectPath um einen Backslash verlaengert als Pfad verwenden.
  ::DAS WAR NACH VIELEN EXPERIMENTEN DIE EINFACHSTE LOESUNG!!!
  set "objectPath2=%objectPath%"
  if "%objectPath2:~-1%" equ ":" set "objectPath2=%objectPath2%\"
  >>"%menuFile%" (for /f "delims=" %%d in ('dir /b /o:ne /a:-s-hd "%objectPath2%" 2^>NUL') do (<NUL set /p "=::: %%d" & echo.))

  ::Eintraege fuer Dateien erzeugen.
  >>"%menuFile%" (for /f "delims=" %%f in ('dir /b /o:ne /a:-s-h-d "%searchMask%" 2^>NUL') do (<NUL set /p "=:::  %%f" & echo.))

  ::Eine Leerzeile an die Menue-Datei anhaengen, damit die Liste der
  ::Menueeintraege vollstaendig dargestellt wird.
  >>"%menuFile%" echo :::


  ::Wenn erlaubt, Zugang zum Tools-Menue anzeigen
  if %allowNewFiles% equ 1 set "showTools=%toolsStr%"
  if %allowNewFilesDirs% equ 1 set "showTools=%toolsStr%"
  if %allowOnlyDirs% equ 1 set "showTools=%toolsStr%"


  :FileSelectorShowMenu
  ::Menue und Eingabeaufforderung anzeigen
  call :ShowMenu "%menuFile%" /i "%inputStr%" /t /c /h 12 /w %maxTitleLen% %showTools%
  set /a "selected=%errorlevel%"

  ::Bei Dialog-Abbruch Ruecksprung
  ::Wenn der Arbeitsplatz angezeigt wird, ist keine Auswahl eines Laufwerks
  ::mit der speziellen Auswahl moeglich.
  if %selected% equ 0 goto :FileSelectorEnd
  if %selected% lss 0 if not defined objectPath goto :FileSelectorSelectLoop
  if %selected% lss 0 set /a "selected=-selected" & set "dirSelected=1"
  if %selected% equ 2147483647 goto :FileSelectorSelectLoop


  ::Wenn objectPath NICHT definiert ist, ist das aktuelle Verzeichnis der
  ::Arbeitsplatz (die Laufwerksauswahl). Dann weitermachen.
  ::Wenn das aktuelle Verzeichnis ein Wurzelverzeichnis ist und auf den
  ::Arbeitsplatz gewechselt werden soll, einen Backslash an objectPath
  ::anhaengen, da sonst der Ausdruck %%~dpp in der FOR-Schleife
  ::das Elternverzeichnis des aktuellen Verzeichnisses des Laufwerks liefern
  ::wuerde. Dadurch wuerde objectPath nicht geloescht und der Wechsel zum
  ::Arbeitsplatz wuerde nicht funktionieren.
  if not defined objectPath goto :FileSelectorSelectLoopEnd
  if %selected% equ 2 if "%objectPath:~-1%" equ ":" set "objectPath=%objectPath%\"

:FileSelectorSelectLoopEnd


::Ausrufezeichen im Namen der Menue-Datei escapen, da jetzt ein Bereich
::mit ENABLEDELAYEDEXPANSION betreten wird
set "menuFile=%menuFile:!=^!%"


::Ausgewaehltes Element verarbeiten
setlocal enabledelayedexpansion

set "indexCnt=0"

for /f "usebackq tokens=1* delims=: " %%f in ("%menuFile%") do (
  set /a "indexCnt=!indexCnt!+1"

  if !indexCnt! equ %selected% (
    endlocal & set "menuFile=%menuFile:^!=!%"

    if "%%f" equ "" (
      if "%%g" equ ".." (
        if %keepDir% equ 1 goto :FileSelectorSelectLoop

        for /f "delims=" %%p in ("%objectPath%") do (
          if /i "%objectPath%" neq "%%~dpp" (
            set "objectPath=%%~dpp"
          ) else (
            set "objectPath="
          )

          goto :FileSelectorSelectLoop
        )
      )

      if %keepDir% equ 1 if %dirSelected% equ 0 goto :FileSelectorSelectLoop

      set "objectPath=%objectPath%\%%g"

      if %allowFilesDirs% equ 0 if %allowNewFilesDirs% equ 0 if %allowOnlyExistingDirs% equ 0 if %allowOnlyDirs% equ 0 goto :FileSelectorSelectLoop
      if %dirSelected% equ 0 goto :FileSelectorSelectLoop

      set "selectedObject=%objectPath%\%%g"
      goto :FileSelectorEnd
    )

    if %allowOnlyExistingDirs% equ 1 goto :FileSelectorSelectLoop
    if %allowOnlyDirs% equ 1 goto :FileSelectorSelectLoop

    if "%%g%" equ "" (
      set "selectedObject=%objectPath%\%%f"
      goto :FileSelectorEnd
    )

    set "selectedObject=%objectPath%\%%f %%g"
    goto :FileSelectorEnd
  )
)


:: Menue-Datei loeschen
:FileSelectorEnd
del "%menuFile%" 2>NUL

if not defined outputFile goto :FileSelectorExit
if not defined selectedObject goto :FileSelectorExit
>"%outputFile%" (<NUL set /p "=%selectedObject%" & echo.)


:FileSelectorExit
::Funktionsergebnis zurueckgeben
::und Ruecksprung
endlocal & set "selectedObject=%selectedObject%"
exit /b




::Funktion FileSelectorTools
::Wird aufgerufen, wenn der Benutzers T fuer Tools aus der Zeile
::mit den Steuerbefehlen der Dateiauswahlbox eingibt.
:FileSelectorTools
setlocal disabledelayedexpansion

set "toolsMenuFile=%TEMP%\FileSelectorToolsMenu.txt"
set "reloadFileSelector=0"


:FileSelectorToolsSelectLoop
  cls

  ::Auswahlmenue generieren.
  > "%toolsMenuFile%" echo :::Tools
  >>"%toolsMenuFile%" echo :::Neue Datei
  >>"%toolsMenuFile%" echo :::Neues Verzeichnis
  >>"%toolsMenuFile%" echo :::Datei lîschen
  >>"%toolsMenuFile%" echo :::Verzeichnis lîschen
  >>"%toolsMenuFile%" echo :::

  call :ShowMenu "%toolsMenuFile%" /i "Was mîchten Sie tun? " /t /c /w 30
  set /a "selected=%errorlevel%"

  ::Bei Dialog-Abbruch Ruecksprung
  if %selected% equ 0 goto :FileSelectorToolsEnd

  ::Spezielle Auswahl wird nicht gebraucht
  if %selected% lss 0 set /a "selected=-selected"

  if %selected% equ 2 goto :FileSelectorToolsCreateFile
  if %selected% equ 3 goto :FileSelectorToolsCreateDirectory
  if %selected% equ 4 goto :FileSelectorToolsDeleteFile
  if %selected% equ 5 goto :FileSelectorToolsDeleteDirectory
goto :FileSelectorToolsSelectLoop


:FileSelectorToolsCreateFile
::Eingabeaufforderung anzeigen
echo. & set /p "fileName=Geben Sie den Namen der Datei ein: "
if not defined fileName goto :FileSelectorToolsSelectLoop

::Dateinamen auf ungueltige Zeichen pruefen
::Wenn ungueltige Zeichen enthalten sind, Fehlermeldung ausgeben,
::auf Tastendruck warten und dann Tools-Menue nochmal anzeigen
set "fileName=%fileName:"=*%"

for /f %%i in ('set /p "=%fileName%" ^<NUL ^| findstr "[\*?:<>|\\/]"') do (
  echo. & echo Ein Dateiname darf keines der Zeichen *?^<^>^|\/:^" enthalten! & pause>NUL
  goto :FileSelectorToolsSelectLoop
)

::Wenn die Datei noch nicht existiert, ist alles OK
if not exist "%objectPath%\%fileName%" goto :FileSelectorToolsWriteFile

::Sonst Auswahlmenue generieren, was jetzt zu tun ist
> "%toolsMenuFile%" echo :::Existierende Datei Åberschreiben?
>>"%toolsMenuFile%" echo :::Ja
>>"%toolsMenuFile%" echo :::Nein
>>"%toolsMenuFile%" echo :::

::Auswahlmenue anzeigen
call :ShowMenu "%toolsMenuFile%" /i "Die Datei existiert bereits. öberschreiben? " /t /c
set /a "selected=%errorlevel%"

::Bei Auswahl von Abbruch Ende
::Spezielle Auswahl wird nicht gebraucht
::Bei Auswahl von Ja wird die Datei angelegt
::Bei Auswahl von Nein nochmal Tools-Menue anzeigen
if %selected% equ 0 goto :FileSelectorToolsEnd
if %selected% lss 0 set /a "selected=-selected"
if %selected% equ 2 goto :FileSelectorToolsWriteFile
goto :FileSelectorToolsSelectLoop

:FileSelectorToolsWriteFile
::Leere Datei anlegen
type NUL>"%objectPath%\%fileName%" || pause>NUL

::Inhalt der Dateiauswahlbox muss neu geladen werden
set "reloadFileSelector=1"
echo.
goto :FileSelectorToolsEnd


:FileSelectorToolsCreateDirectory
::Eingabeaufforderung anzeigen
echo. & set /p "dirName=Geben Sie den Namen des Verzeichnisses ein: "
if not defined dirName goto :FileSelectorToolsSelectLoop

::Verzeichnisnamen auf ungueltige Zeichen pruefen
::Wenn ungueltige Zeichen enthalten sind, Fehlermeldung ausgeben,
::auf Tastendruck warten und dann Tools-Menue nochmal anzeigen
set "dirName=%dirName:"=*%"

for /f %%i in ('set /p "=%dirName%" ^<NUL ^| findstr "[\*?:<>|\\/]"') do (
  echo. & echo Ein Verzeichnisname darf keines der Zeichen *?^<^>^|\/:^" enthalten! & pause > NUL
  goto :FileSelectorToolsSelectLoop
)

::Wenn das Verzeichnis noch nicht existiert ist alles OK
if not exist "%objectPath%\%dirName%" goto :FileSelectorToolsWriteDir

::Sonst Fehlermeldung ausgeben und auf Tastendruck warten,
::dann Tools-Menue nochmal anzeigen
echo. & echo Dieses Verzeichnis existiert bereits! & pause>NUL
goto :FileSelectorToolsSelectLoop

:FileSelectorToolsWriteDir
::Verzeichnis anlegen
md "%objectPath%\%dirName%" || pause>NUL

::Inhalt der Dateiauswahlbox muss neu geladen werden
set "reloadFileSelector=1"
echo.
goto :FileSelectorToolsEnd


:FileSelectorToolsDeleteFile
::Dateiauswahlbox anzeigen. Durch Parameter /f koennen nur existierende
::Dateien ausgewaehlt werden und durch /s kann das Verzeichnis nicht
::gewechselt werden. Bei Abbruch Ende
call :FileSelector "%objectPath%" /m *.* /i "WÑhlen Sie die Datei aus: " /f /s
if not defined selectedObject goto :FileSelectorToolsEnd

::Dialogbox zum bestaetigen des Loeschvorgangs anzeigen
call :FileSelectorToolsWriteConfirmationMenu
call :ShowMenu "%toolsMenuFile%" /i "Diese Datei lîschen? " /t /c /w 72
set /a "selected=%errorlevel%"

::Bei Eingabe von Abbruch oder Nein Ende
if %selected% equ 0 goto :FileSelectorToolsEnd
if %selected% lss 0 set /a "selected=-selected"
if %selected% equ 3 goto :FileSelectorToolsEnd

::Datei loeschen
::Falls ein Fehler aufgetreten ist, muss eine Taste gedrueckt werden,
::damit man die Fehlermeldung ansehen kann.
>NUL del /q "%selectedObject%" || pause>NUL

::Inhalt der Dateiauswahlbox muss neu geladen werden
set "reloadFileSelector=1"
echo.
goto :FileSelectorToolsEnd


:FileSelectorToolsDeleteDirectory
::Dateiauswahlbox anzeigen. Durch Parameter /d koennen nur existierende
::Verzeichnisse ausgewaehlt werden und durch /s kann das Verzeichnis nicht
::gewechselt werden. Bei Abbruch Ende
call :FileSelector "%objectPath%" /m *.* /i "WÑhlen Sie das Verzeichnis aus: " /d /s
if not defined selectedObject goto :FileSelectorToolsEnd

::Dialogbox zum bestaetigen des Loeschvorgangs anzeigen
call :FileSelectorToolsWriteConfirmationMenu
call :ShowMenu "%toolsMenuFile%" /i "Dieses Verzeichnis lîschen? " /t /c /w 72
set /a "selected=%errorlevel%"

::Bei Eingabe von Abbruch oder Nein Ende
if %selected% equ 0 goto :FileSelectorToolsEnd
if %selected% lss 0 set /a "selected=-selected"
if %selected% equ 3 goto :FileSelectorToolsEnd

::Verzeichnis und alle Unterverzeichnisse ohne Nachfrage loeschen.
::Falls ein Fehler aufgetreten ist, muss eine Taste gedrueckt werden,
::damit man die Fehlermeldungen ansehen kann.
>NUL rd /s /q "%selectedObject%" || pause>NUL

::Inhalt der Dateiauswahlbox muss neu geladen werden
set "reloadFileSelector=1"
echo.
goto :FileSelectorToolsEnd


:FileSelectorToolsEnd
::Menuedatei loeschen
>NUL del "%toolsMenuFile%"

::Ruecksprung mit Rueckgabewert
exit /b %reloadFileSelector%



::Stellt eine Dialogbox zum bestaetigen eines Loeschvorgangs dar.
::In der Titelzeile wird die Datei/das Verzeichnis angezeigt, das geloescht
::werden soll. Die Pfadlaenge wird bei Bedarf so gekuerzt, dass das Laufwerk
::und das hintere Ende des Pfades auf jeden Fall angezeigt wird.
:FileSelectorToolsWriteConfirmationMenu
setlocal

set "fstConfirmTitle=%selectedObject%"
set "fstConfirmTitleDrv=%selectedObject:~0,3%"
set "fstConfirmTitlePath=%selectedObject:~3%"
if not defined fstConfirmTitlePath goto :FileSelectorToolsWriteConfirmationFile

set "correction=9"

:FileSelectorToolsTrimTitleLoop
  >"%TEMP%\strlen.txt" (set /p "=<%fstConfirmTitlePath%"&echo.&set /p "=<"&echo.) <NUL
  for /f "delims=:" %%l in ('findstr /o /b "<" "%TEMP%\strlen.txt"') do set /a "fstTitleLen=%%l+correction"

  if %fstTitleLen% leq 72 (
    del "%TEMP%\strlen.txt" 2>NUL
    if %correction% neq 9 goto :FileSelectorToolsTrimTitleLoopEnd
    goto :FileSelectorToolsWriteConfirmationFile
  )

  set "fstConfirmTitlePath=%fstConfirmTitlePath:*\=%"

  set "correction=13"
  goto :FileSelectorToolsTrimTitleLoop
:FileSelectorToolsTrimTitleLoopEnd

::Die gekuerzte Titelzeile zusammenbauen
set "fstConfirmTitle=%fstConfirmTitleDrv%...\%fstConfirmTitlePath%"

:FileSelectorToolsWriteConfirmationFile
> "%toolsMenuFile%" (set /p "=:::%fstConfirmTitle% lîschen?" <NUL & echo.)
>>"%toolsMenuFile%" echo :::Ja
>>"%toolsMenuFile%" echo :::Nein
>>"%toolsMenuFile%" echo :::

exit /b




::Funktion ShowMenu
::
::Zeigt ein umrahmtes Menue an, dessen Eintraege innerhalb des Rahmens
::zentriert dargestellt werden. Die Hoehe des Rahmens passt sich der Anzahl der
::Menueeintraege an, kann aber auch ueber den Parameter /h fest eingestellt
::werden. Umfasst das Menue mehr Eintraege als in den Rahmen passen, kann
::geblaettert werden. Die Breite des Rahmens passt sich dem laengsten Eintrag
::im Menue, der Titelzeile (Parameter /t) oder der Zeile mit den Steuerbefehlen
::an, je nach dem, was laenger ist. Die Rahmenbreite kann aber auch durch den
::Parameter /w festgelegt werden. Der Rahmen wird aber mindestens so breit
::sein, das die Zeile mit den Steuerbefehlen vollstaendig hineinpasst. Menue-
::eintraege, die nicht komplett in den Rahmen passen, werden abgeschnitten,
::ebenso die Titelzeile. Der Rahmen kann wahlweise zentriert auf dem Bildschirm
::dargestellt werden. Ein Menueeintrag wird ueber seine Nummer ausgewaehlt.
::Die Menueeintraege muessen der Funktion in einer Datei uebergeben werden. Die
::Nummerierung (inkl. fuehrender Nullen) wird von ShowMenu durchgefuehrt. Ein
::Doppelpunkt als erstes Zeichen eines Menueeintrags ist nicht mîglich. Die
::letzte Zeile der Datei muss drei Doppelpunkte (eine Leerzeile) enthalten. Nur
::so kann, fuer den Fall, dass der letzte Menueeintrag der laengste ist, die
::Breite des Rahmens korrekt berechnet werden und nur dann werden alle Menue-
::eintraege dargestellt. Das Einfuegen von Leerzeilen ins Menue ist zwar
::moeglich, diese werden aber auch nummeriert. Durch den Parameter /u kann der
::Zeile mit den Stuerbefehlen EIN benutzerdefinierter Eintrag hinzugefuegt
::werden. Dazu muss der Text des Eintrags, die Taste fuer seine Auswahl und
::der Name der Funktion, die dadurch aufgerufen werden soll, uebergeben
::werden. Die Ausloesetaste kann nicht c oder C sein, da dieser Code intern
::fuer die spezielle Auswahl benutzt wird, d.h. man gibt CX ein (X steht fÅr
::die Nummer des auszuwaehlenden Menueeintrags). Anstatt der Zeilennummer des
::Menueeintrags in der Menuedatei wird dann die negative Zeilennummer zurueck-
::gegeben. Sinn der Sache ist folgender: Wenn ShowMenu z.B. dazu benutzt wird,
::um eine Dateiauswahlbox darzustellen, kann durch Eingabe der Nummer X eines
::Menueintrags, der ein Verzeichnis darstellt, in dieses Verzeichnis gewechselt
::werden. Durch Eingabe von CX kann das aufrufende Skript an der dann zurueck-
::gegebenen negativen Nummer X des Menueeintrags erkennen, dass nicht in das
::ausgewaehlte Verzeichnis gewechselt werden soll, sondern der Verzeichnisname
::selbst ausgewaehlt werden soll.
::
::Die Funktion gibt in ERRORLEVEL die Zeilennummer des ausgewaehlten Menue-
::eintrags in der Menuedatei zurueck. Unter Verwendung der speziellen Auswahl
::ist diese Nummer negativ. Wenn in der Menuedatei eine Titelzeile enthalten
::ist, ist die Nummer des ersten Menueeintrags 2. Ohne Titelzeile ist die
::Nummer gleich 1. Im Query-Mode (Parameter /q) wird die Laenge der laengsten
::Zeile in der Menuebox (die Breite des beschreibbaren Bereichs) zurueck-
::gegeben. Bei Auswahl von A oder a (fuer Abbruch) wird 0 zurueckgeliefert.
::Wenn das Menue neu geladen werden muss, wird 2147483647 zurueckgegeben.
::Bei einem Fehler in den Parametern wird -2147483648 zurueckgeliefert.
::
::Die Reihenfolge der Parameter ist beliebig.
::
::
::Aufruf   : call :ShowMenu Datei [/i Eingabeaufforderung] [/s Startzeile]
::                          [/h Zeilenanzahl] [/w Spaltenanzahl]
::                          [/u Taste Eintrag Funktion]
::                          [/t] [/c] [/q]
::
::Parameter: Datei
::             Pfad zur Datei, die die Menueeintraege enthaelt.
::             Pro Zeile der Datei wird eine Zeile des Menues angegeben.
::             Eine Zeile muss mit drei Doppelpunkten (:::) beginnen,
::             danach kommt der Text der Menuezeile. Die lezte Zeile der
::             Datei muss eine Leerzeile sein.
::             Beispiel: :::Eintrag 1
::                       :::Eintrag 2
::                       :::
::
::           Eingabeaufforderung
::             Der Text, der als Eingabeaufforderung angezeigt werden soll.
::             Standard ist Auswahl?
::
::           Startzeile
::             Gibt die Zeile an, die als erstes im Menue erscheinen soll.
::             Standard ist 1.
::
::           Zeilenanzahl
::             Gibt die Anzahl Zeilen des beschreibbaren Bereichs in der
::             Menuebox an.
::             Zulaessige Werte sind 1-12 mit Titelzeile (Standard 12) und
::             1-14 ohne Titelzeile (Standard 14). Selbst wenn die Menuebox
::             mit allen Eintraegen auf den Schirm passen wuerde, kann
::             hiermit eine Maximalhoehe vorgegeben werden.
::
::           Spaltenanzahl
::             Gibt die Anzahl Spalten des beschreibbaren Bereichs in der
::             Menuebox an.
::             Zulaessige Werte sind 58-74 (58 ist die Breite der Zeile mit
::             den Steuerbefehlen). Ohne diesen Parameter passt sich die Breite
::             der Menuebox an die Laenge des laengsten Menueeintrags bzw. der
::             Titelzeile an.
::
::           Benutzerdefinierter Steuerbefehl
::             Durch Uebergabe des Parameter /u ist es moeglich, EINEN
::             benutzerdefinierten Eintrag in der Zeile mit den Steuerbefehlen
::             anzuzeigen. Taste ist der Buchstabe, um die benutzerdefinierte
::             Funktion, deren Namen durch Funktion festgelegt wird,
::             aufzurufen. Eintrag ist der Text, der in der Zeile mit den
::             Steuerbefehlen angezeigt wird. Taste darf nicht den Wert c oder
::             C haben, da dieser Code intern fuer die spezielle Auswahl
::             benutzt wird. Die aufzurufende Funktion muss einen von 0
::             verschiedenen Wert zurueckliefern, wenn das Menu neu geladen
::             werden muss (ShowMenu gibt dann 2147483647 an seinen Aufrufer
::             zurueck), sonst 0.
::
::           Titelzeile
::             Wenn der Parameter /t angegeben wird, wird die erste Zeile der
::             Menue-Datei als Titelzeile in der Menuebox angezeigt. Dadurch
::             wird bei Auswahl des ersten Menueeintrags 2 zurueckgegeben,
::             bei Auswahl des zweiten 3 usw.
::
::           Center
::             Wenn der Parameter /c angegeben wird, wird die Menuebox auf dem
::             Bildschirm zentriert. Standard ist linksbuendige Darstellung.
::
::           Query-Mode
::             Durch die Angabe des Parameters /q kann die Breite des
::             beschreibbaren Bereichs der Menuebox erfragt werden. Das kann
::             z.B. dazu benutzt werden, um die Titelzeile zu formatieren.
::
::Rueckgabe: In ERRORLEVEL wird die Zeilennummer innerhalb der Menuedatei des
::           ausgewaehlten Menueeintrags zurueckgegeben. Diese Zeilenummer ist
::           negativ, wenn ein Eintrag mit der speziellen Auswahl (CX, X ist
::           die Eintragsnummer) ausgewaehlt wurde. Wenn eine Titelzeile
::           angezeigt wird, ist die Zeilennummer des ersten Menueeintrags 2
::           (bzw. -2), ohne Titelzeile ist sie 1 (bzw. -1). Bei Auswahl von
::           a oder A fuer Abbruch wird 0 zurueckgegeben. Wenn das Menue neu
::           geladen werden muss, wird 2147483647 zurueckgeliefert und bei
::           Fehlern in den Parametern der Wert -2147483648.
::
::Bedienung: a oder A - bricht die Funktion ab.
::           u        - blaettert eine Zeile nach unten
::           o        - blaettert eine Zeile nach oben
::           U        - blaettert eine Seite nach unten
::           O        - blaettert eine Seite nach oben
::           b oder B - blaettert zum Anfang der Liste
::           e oder E - blaettert zum Ende der Liste


:ShowMenu
setlocal disabledelayedexpansion

::Die Rahmenelemente definieren
set "ulCorner=…"
set "urCorner=ª"
set "llCorner=»"
set "lrCorner=º"
set "hBar=Õ"
set "vBar=∫"


::Die Zeile mit den Steuerbefehlen definieren
set "ctrlLine=A=Abbruch"
set "ctrlLineLen=9
set "scrollCtrls=  o/O=nach oben  u/U=nach unten  B=Beginn  E=Ende"


::Zur Sicherheit Variabeln loeschen
set "inputStr=Auswahl? "
set "center=0"
set "specialChoice=0"

set "startLine="
set "contentLines="
set "contentCols="
set "readTitleStr="
set "titleStr="
set "titleLine="
set "queryMode="
set "menuFile="
set "userKey="
set "userOption="
set "userHandler="


::Parameter ermitteln
:ShowMenuChkParamsLoop
  set "param=%~1"

  if /i "%param%" equ "/u" (
    set "userKey=%~2" & set "userOption=%~3" & set "userHandler=%~4" & shift & shift & shift
  ) else if /i "%param%" equ "/h" (
    set /a "contentLines=%~2" & shift
  ) else if /i "%param%" equ "/w" (
    set /a "contentCols=%~2" & shift
  ) else if /i "%param%" equ "/s" (
    set /a "startLine=%~2"& shift
  ) else if /i "%param%" equ "/i" (
    set "inputStr=%~2" & shift
  ) else if /i "%param%" equ "/t" (
    set "readTitleStr=1"
  ) else if /i "%param%" equ "/q" (
    set "queryMode=1"
  ) else if /i "%param%" equ "/c" (
    set /a "center=1"
  ) else if "%param:~0,1%" neq "/" if not defined menuFile (
    set "menuFile=%~f1%"
  )

  shift
if "%~1" neq "" goto :ShowMenuChkParamsLoop


::Wenn der Name der Menue-Datei ein ^ enthaelt, wurde es vom Kommando-
::Interpreter beim Aufruf von ShowMenu aus einem Batch-Skript durch ^^
::ersetzt. Das muss wieder rueckgaengig gemacht werden.
set "menuFile=%menuFile:^^=^%"


::Bei falschen Parametern RÅcksprung
if not defined menuFile exit /b -2147483648
if not exist "%menuFile%" exit /b -2147483648


::Wenn eine Titelzeile ausgegeben werden soll, die maximale Anzahl von
::Menueeintraegen um 2 niedriger festsetzen und Titelzeile einlesen.
if defined readTitleStr (
  set /p "titleStr=" <"%menuFile%"
  set "maxContentLines=12"
) else (
  set "maxContentLines=14"
)


::Ausrufezeichen im Namen der Menue-Datei escapen, da im weiteren Verlauf
::des Skripts die Variable nur noch in Bereichen mit ENABLEDELAYEDEXPANSION
::benutzt wird, in denen Ausrufezeichen im Variableninhalt Probleme machen.
set "menuFile=%menuFile:!=^!%"


::Breite des laengsten Menueeintrags und der Titelzeile ermitteln.
::Dabei auch gleich die Anzahl der Menueeintraege zaehlen.
setlocal enabledelayedexpansion

set "longestLine=0"
set "chrOffset=0"
set "numItems=0"
set "titleLineLen=0"

for /f "delims=:" %%i in ('findstr /o /b ":::" "%menuFile%"') do (
  set /a "numItems+=1"

  set /a "lineLen=%%i-chrOffset-5"
  set /a "chrOffset=%%i"

  if !longestLine! lss !lineLen! (
    if not defined titleStr (
      set /a "longestLine=lineLen"
    ) else (
      if !numItems! equ 2 (
        set /a "titleLineLen=lineLen"
      ) else (
        set /a "longestLine=lineLen"
      )
    )
  )
)

endlocal & set "longestLine=%longestLine%" & set "numItems=%numItems%" & set "titleLineLen=%titleLineLen%"


::Durch die zusÑtzliche Leerzeile in der Menue-Datei muss numItems korrigiert
::werden. Die Hoehe der Korrektur haengt davon ab, ob in der Menue-Datei eine
::Titelzeile enthalten ist.
if defined titleStr (
  set /a "numItems-=2"
) else (
  set /a "numItems-=1"
)


::Hoehe der Menuebox ermitteln
if defined contentLines goto :ShowMenuCheckContentLinesBounds

::Hoehe der Menuebox abhaengig von der Anzahl der Menueeintraege setzen.
if %numItems% leq %maxContentLines% (
  set /a "contentLines=numItems"
) else (
  set /a "contentLines=maxContentLines"
)

::Keine Leerzeilen nach den Menueeintraegen ausgeben
set "trailingBlankLines=0"
goto :ShowMenuSetCtrlLine

:ShowMenuCheckContentLinesBounds
::Obere und untere Grenze der uebergebenen contentLines pruefen.
if %contentLines% lss 1 set "contentLines=1"
if %contentLines% gtr %maxContentLines% set /a "contentLines=maxContentLines"

::Anzahl Leerzeilen ausrechnen, die nach den Menueeintraegen ausgegeben werden
::muessen, um eine konstante Hoehe der Menuebox zu erreichen, wenn weniger
::Menueeintraege als Zeilen vorhanden sind.
set /a "trailingBlankLines=contentLines-numItems"


:ShowMenuSetCtrlLine
::Falls nicht alle Menueeintraege in die Menuebox passen, die Zeile mit den
::Steuerbefehlen um die Scroll-Steuerung verlaengern und Laenge anpassen
if %numItems% gtr %contentLines% (
  set "ctrlLine=%ctrlLine%%scrollCtrls%"
  set "ctrlLineLen=58"
)


::Falls eine benutzerdefinierte Steuerfunktion uebergeben wurde, die Zeile mit
::den Steuerbefehlen und deren Laenge anpassen. Die benutzerdefinierte Taste
::darf nicht c oder C sein, weil dieser Code die fÅr die spezielle Auswahl
::reserviert ist.
if not defined userKey goto :ShowMenuCountNumItemsDigits
if not defined userOption exit /b -2147483648
if not defined userHandler exit /b -2147483648

set "userKey=%userKey:~0,1%"
if /i "%userKey%" equ "C" exit /b -2147483648

set "ctrlLine=%ctrlLine%  %userOption%"
>"%TEMP%\strlen.txt" (set /p "=<%ctrlLine%"&echo.&set /p "=<"&echo.) <NUL
for /f "delims=:" %%l in ('findstr /o /b "<" "%TEMP%\strlen.txt"') do set /a "ctrlLineLen=%%l-3"
del "%TEMP%\strlen.txt" > NUL

if %ctrlLineLen% gtr 74 (
  set "ctrlLine=%ctrlLine:~0,74%"
  set "ctrlLineLen=74"
)


:ShowMenuCountNumItemsDigits
::Bestimmen, wie viele Stellen numItems hat und fuer die Nummerierung
::String fuer fuehrende Nullen generieren.
set "numItems2=%numItems%"
set "numDigits=1"
set "zeros="

:ShowMenuCountDigitsLoop
  set /a "numItems2/=10"
if %numItems2% gtr 0 (set /a "numDigits+=1" & set "zeros=%zeros%0" & goto :ShowMenuCountDigitsLoop)


::longestLine um die Anzahl Zeichen fuer die Nummerierung und den Zwischenraum
::zwischen Nummern und Menueeintraegen erhoehen
set /a "longestLine+=numDigits+2"


::Wenn die breiteste Zeile oder die Titelzeile breiter als zulaessig,
::auf Maximalwert setzen
if %longestLine% gtr 74 set "longestLine=74"
if %titleLineLen% gtr 74 set "titleLineLen=74"


::Innere Breite der Menuebox ermitteln.
if defined contentCols goto :ShowMenuCheckContentColsBounds

::Breite der Menuebox auf das Maximum aus longestLine, ctrlLineLen und,
::falls eine Titelzeile ausgegeben werden soll, titleLineLen setzen
if %longestLine% gtr %ctrlLineLen% (
  set /a "innerBoxWidth=longestLine"
) else (
  set /a "innerBoxWidth=ctrlLineLen"
)

if %titleLineLen% gtr %innerBoxWidth% (
  set /a "innerBoxWidth=titleLineLen"
)

goto :ShowMenuCheckQueryMode

:ShowMenuCheckContentColsBounds
::Ober- und Untergrenze von contentCols pruefen
if %contentCols% lss %ctrlLineLen% set /a "contentCols=ctrlLineLen"
if %contentCols% gtr 74 set "contentCols=74"

set /a "innerBoxWidth=contentCols"

::longestLine und titleLineLen an die vorgegebene Breite der Menuebox anpassen
if %longestLine% gtr %innerBoxWidth% set /a "longestLine=innerBoxWidth"
if %titleLineLen% gtr %innerBoxWidth% set /a "titleLineLen=innerBoxWidth"


:ShowMenuCheckQueryMode
::Wenn nur die Laenge der laengsten Zeile abgefragt werden soll, Ruecksprung
if defined queryMode exit /b %innerBoxWidth%


::innerBoxWidth um das minimale Padding erhoehen
set /a "innerBoxWidth+=4"


::Die maximal moegliche startLine ermitteln.
if defined titleStr goto :ShowMenuCalcMaxStartLineWithTitleLine

::Ohne Titelzeile
set /a "maxStartLine=numItems-contentLines+1"
if %maxStartLine% lss 1 set "maxStartLine=1"
goto :ShowMenuSetStartLine

::Mit Titelzeile
:ShowMenuCalcMaxStartLineWithTitleLine
set /a "maxStartLine=numItems-contentLines+2"
if %maxStartLine% lss 2 set "maxStartLine=2"


::Den ersten anzuzeigenden Menueeintrag ermitteln
:ShowMenuSetStartLine
if defined startLine goto :ShowMenuCheckStartLineBounds

::Keine startLine uebergeben => startLine auf Standardwert setzen
::Dabei beruecksichtigen, ob eine Titelzeile angezeigt werden soll oder nicht.
if not defined titleStr (
  set "startLine=1"
) else (
  set "startLine=2"
)

goto :ShowMenuCalcPaddings

::Obere und untere Grenze von uebergebener startLine pruefen
:ShowMenuCheckStartLineBounds
if not defined titleStr (
  if %startLine% lss 1 set "startLine=1"
) else (
  if %startLine% lss 2 set "startLine=2"
)

::Bei zu grossem Wert auf den Maximalwert setzen
if %startLine% gtr %maxStartLine% set /a "startLine=maxStartLine"


:ShowMenuCalcPaddings
::Paddings errechnen
set /a "itemLinePadding=(innerBoxWidth-longestLine)/2"
set /a "ctrlLinePadding=(innerBoxWidth-ctrlLineLen)/2"


::DELAYEDEXPANSION einschalten, damit man im folgenden Bereich FOR-Schleifen
::statt GOTO-Schleifen benutzen kann
setlocal enabledelayedexpansion

::Zeile fuer vertikalen Rand der Menuebox und die Leerzeilen erzeugen
set "vertBorder="
set "blankLine="

for /l %%i in (1,1,%innerBoxWidth%) do (
  set "vertBorder=!vertBorder!%hBar%"
  set "blankLine=!blankLine! "
)


::Die Leerzeichen fuer Padding von den Zeilen mit den Menuepunkten erzeugen
set "itemLinePaddingBlanks="

for /l %%i in (1,1,%itemLinePadding%) do (
  set "itemLinePaddingBlanks=!itemLinePaddingBlanks! "
)


::Die Leerzeichen fuer Padding von der Zeile mit den Steuerbefehlen erzeugen
set "ctrlLinePaddingBlanks="

for /l %%i in (1,1,%ctrlLinePadding%) do (
  set "ctrlLinePaddingBlanks=!ctrlLinePaddingBlanks! "
)


::Die Leerzeichen fuer die Zentrierung der Menuebox erzeugen
set "screenPadding=0"
set "screenPaddingBlanks="

if %center% neq 0 (
  set /a "screenPadding=(80-(innerBoxWidth+2))/2"

  for /l %%i in (1,1,!screenPadding!) do (
    set "screenPaddingBlanks=!screenPaddingBlanks! "
  )
)

::DELAYEDEXPANSION ausschalten, damit man die Titelzeile und die Zeile mit
::den Steuerbefehlen, die evtl. Ausrufezeichen enthalten koennten, problemlos
::zusammensetzen kann.
endlocal & set "vertBorder=%vertBorder%" & set "blankLine=%blankLine%" & set "itemLinePaddingBlanks=%itemLinePaddingBlanks%" & set "ctrlLinePaddingBlanks=%ctrlLinePaddingBlanks%" & set "screenPaddingBlanks=%screenPaddingBlanks%"


::Zeile mit den Steuerbefehlen inkl. Padding erzeugen
set "ctrlLine=%ctrlLinePaddingBlanks%%ctrlLine%%blankLine%"
call set "ctrlLine=%%ctrlLine:~0,%innerBoxWidth%%%"


::Evtl. Titelzeile inkl. Padding erzeugen. Die drei fuehrenden Doppelpunkte
::abschneiden. Sonderzeichen escapen. Klammern muessen escaped werden, weil
::die Ausgabebefehle fuer die Titelzeile in einem geklammerten IF-ELSE-Block
::stehen.
if not defined titleStr goto :ShowMenuWithoutTitleLine

set /a "titleLineWidth=innerBoxWidth-4"

set "titleLine=%titleStr:~3%%blankLine%"
call set "titleLine=  %%titleLine:~0,%titleLineWidth%%%  "
set "titleLine=%titleLine:^=^^%"
set "titleLine=%titleLine:&=^&%"
set "titleLine=%titleLine:(=^(%"
set "titleLine=%titleLine:)=^)%"


::Die Hauptschleife von ShowMenu muss mit DELAYEDEXPANSION laufen. Dort wird
::zwar zweimal DELAYEDEXPANSION ausgeschaltet, aber ohne die jetzt folgende
::ENABLEDELAYEDEXPANSION-Zeile muesste dreimal umgeschaltet werden.
:ShowMenuWithoutTitleLine
setlocal enabledelayedexpansion


::Hauptschleife von ShowMenu, in der die Menuebox dargestellt wird
::Falls die Menuebox bildschirmfuellend wird, andere Ausgabebefehle benutzen
:ShowMenuSelectLoop
  cls

  ::DELAYEDEXPANSION ausschalten, damit Ausrufezeichen in der Titelzeile
  ::keine Probleme machen.
  setlocal disabledelayedexpansion

  ::Oberen Rand, eine Leerzeile und evtl. Titelzeile und
  ::noch eine Leerzeile ausgeben
  if %innerBoxWidth% equ 78 (
    <NUL set /p =%ulCorner%%vertBorder%%urCorner%
    <NUL set /p =%vBar%%blankLine%%vBar%
    if defined titleLine <NUL set /p =%vBar%%titleLine%%vBar%
    if defined titleLine <NUL set /p =%vBar%%blankLine%%vBar%
  ) else (
    echo %screenPaddingBlanks%%ulCorner%%vertBorder%%urCorner%
    echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
    if defined titleLine echo %screenPaddingBlanks%%vBar%%titleLine%%vBar%
    if defined titleLine echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
  )

  endlocal


  ::Inhalt der Menuebox ausgeben.
  ::Prozedur zum Ausgeben der Menueeintraege nur fuer die Zeilen aufrufen,
  ::die ausgegeben werden muessen (wird schneller ausgefuehrt). Bei der
  ::Ausgabe DELAYEDEXPANSION ausschalten, damit Ausrufezeichen in den
  ::Menueeintraegen keine Probleme machen. Durch die zwei verschiedenen
  ::Methoden, um itemLine zusammenzusetzen, wird sichergestellt, das die
  ::Menueeintraege auch Doppelpunkte enthalten koennen (ausser als erstes
  ::Zeichen, da Doppelpunkte von der FOR-Schleife als Trennzeichen behandelt
  ::werden).
  set "itemCntr=0"
  set /a "maxItem=startLine+contentLines-1"
  if %maxItem% gtr %numItems% (if defined titleLine (set /a "maxItem=numItems+1") else (set /a "maxItem=numItems"))

  for /f "usebackq tokens=1,* delims=" %%i in ("%menuFile%") do (
    set /a "itemCntr+=1"

    if !itemCntr! geq %startLine% if !itemCntr! leq %maxItem% (
      setlocal disabledelayedexpansion
      set "itemLine=%%i"
      call :ShowMenuPrintMenuItems
      endlocal
    )
  )


  ::Leerzeilen ausgeben, um eine konstante Hoehe der Menuebox
  ::zu erreichen.
  for /l %%i in (1,1,%trailingBlankLines%) do (
    if %innerBoxWidth% equ 78 (
      <NUL set /p =%vBar%%blankLine%%vBar%
    ) else (
      echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
    )
  )


  ::Zwei Leerzeilen, die Zeile mit den Steuerbefehlen, noch eine Leerzeile
  ::und den unteren Rand der Menuebox ausgeben
  if %innerBoxWidth% equ 78 (
    <NUL set /p =%vBar%%blankLine%%vBar%
    <NUL set /p =%vBar%%blankLine%%vBar%
    <NUL set /p =%vBar%%ctrlLine%%vBar%
    <NUL set /p =%vBar%%blankLine%%vBar%
    <NUL set /p =%llCorner%%vertBorder%%lrCorner%
  ) else (
    echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
    echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
    echo %screenPaddingBlanks%%vBar%%ctrlLine%%vBar%
    echo %screenPaddingBlanks%%vBar%%blankLine%%vBar%
    echo %screenPaddingBlanks%%llCorner%%vertBorder%%lrCorner%
  )


  ::Zwei Leerzeilen und die Eingabeaufforderung ausgeben
  echo. & echo.
  set "selected="
  set /p "selected=%inputStr%"


  ::Bei leerer Eingabe Menue nochmal darstellen.
  if not defined selected goto :ShowMenuSelectLoop


  ::Sonderzeichen aus der Eingabe herausfiltern.
  ::Vermeidet zufaelliges Cross-Site Scripting.
  set "selected=%selected:!= %"
  set "selected=%selected:^= %"
  set "selected=%selected:"= %"
  set "selected=%selected:'= %"
  set "selected=%selected:`= %"
  set "selected=%selected:&= %"
  set "selected=%selected:<= %"
  set "selected=%selected:>= %"
  set "selected=%selected:|= %"


  ::Besteht die Eingabe nur aus zulaessigen Zeichen?
  ::Wenn nicht, Menue nochmal darstellen
  for /f %%i in ('echo %selected%^|findstr /i "[^aoubec0-9%userKey%]"') do (
    goto :ShowMenuSelectLoop
  )


  ::Auf die Eingabe von Steuerbefehlen reagieren
  if /i "%selected%" equ "A" set "selected=0" & goto :ShowMenuSelectLoopBreak
  if "%selected%" equ "o" if defined titleLine (if %startLine% gtr 2 set /a "startLine-=1") else (if %startLine% gtr 1 set /a "startLine-=1") & goto :ShowMenuSelectLoop
  if "%selected%" equ "u" if %startLine% lss %maxStartLine% set /a "startLine+=1" & goto :ShowMenuSelectLoop
  if "%selected%" equ "O" set /a "startLine-=contentLines-1" & if defined titleLine (if !startLine! lss 2 set /a "startLine=2") else (if !startLine! lss 1 set /a "startLine=1") & goto :ShowMenuSelectLoop
  if "%selected%" equ "U" set /a "startLine+=contentLines-1" & if !startLine! gtr %maxStartLine% set /a "startLine=maxStartLine" & goto :ShowMenuSelectLoop
  if /i "%selected%" equ "B" if defined titleLine (set "startLine=2") else (set "startLine=1") & goto :ShowMenuSelectLoop
  if /i "%selected%" equ "E" set /a "startLine=maxStartLine" & goto :ShowMenuSelectLoop

  if not defined userKey goto :ShowMenuCheckSpecialChoice
  ::Benutzerdefinierte Funktion aufrufen. Wenn der Rueckgabewert ungleich 0
  ::ist, mit dem Code fÅr Menue neu laden back to caller.
  if /i "%selected%" equ "%userKey%" (call :%userHandler% && goto :ShowMenuSelectLoop || set "selected=2147483647" & goto :ShowMenuSelectLoopBreak)

  ::Auf Code fuer spezielle Auswahl pruefen
  :ShowMenuCheckSpecialChoice
  if /i "%selected:~0,1%" equ "C" set "selected=%selected:~1%" & set "specialChoice=1"


  ::Fuehrende Nullen entfernen, damit selected nicht als Oktalzahl
  ::angesehen wird
  for /f "tokens=* delims=0" %%n in ("%selected%") do set "selected=%%n"


  ::Eingabe in eine Zahl wandeln
  if defined titleLine (set /a "selected=%selected%+1") else (set /a "selected=%selected%")


  ::Ist die eingegebene Nummer des Menueeintrags im zulaessigen Bereich?
  ::Wenn ja, diese Nummer als Funktionsergebnis zurueckgeben.
  ::Wenn nicht, Menue nochmal darstellen.
  if defined titleLine (
    set /a "numItems2=numItems+1"

    for /l %%i in (2,1,!numItems2!) do (
      if "%selected%" equ "%%i" goto :ShowMenuSelectLoopBreak
    )
  ) else (
    for /l %%i in (1,1,%numItems%) do (
      if "%selected%" equ "%%i" goto :ShowMenuSelectLoopBreak
    )
  )
goto :ShowMenuSelectLoop


:ShowMenuSelectLoopBreak
::Wenn eine spezielle Auswahl gemacht wurde, die Zeilennummer des
::Menueeintrags als negativen Wert zurueckgeben
if %specialChoice% equ 1 set /a "selected=-selected"


::Umgebung mit ENABLEDELAYEDEXPANSION schliessen und Variable selected
::an die Umgebung eine Ebene hoeher uebergeben.
endlocal & set "selected=%selected%"


:ShowMenuEnd
::Ruecksprung und Funktionsergebnis in ERRORLEVEL zurueckliefern
exit /b %selected%



::Die Prozedur zum Ausgeben der Menueeintraege. Durch die Auslagerung in ein
::Unterprogramm koennen die Variablen, aus denen ein Menueeintrag zusammen-
::gesetzt wird, veraendert werden und die Eintraege koennen Ausrufezeichen
::enthalten.
:ShowMenuPrintMenuItems
if defined titleLine (set /a "itemCntr2=itemCntr-1") else (set /a "itemCntr2=itemCntr")
set "itemIndex=%zeros%%itemCntr2%"
call set "itemIndex=%%itemIndex:~-%numDigits%%%"

set "itemLine=%itemLine:~3%"
set "itemLine=%itemIndex%  %itemLine%%blankLine%"
call set "itemLine=%itemLinePaddingBlanks%%%itemLine:~0,%longestLine%%%%blankLine%"
call set "itemLine=%%itemLine:~0,%innerBoxWidth%%%"
set "itemLine=%itemLine:^=^^%"
set "itemLine=%itemLine:&=^&%"
set "itemLine=%itemLine:(=^(%"
set "itemLine=%itemLine:)=^)%"

if %innerBoxWidth% geq 78 (
  <NUL set /p =%vBar%%itemLine%%vBar%
) else (
  echo %screenPaddingBlanks%%vBar%%itemLine%%vBar%
)

exit /b
