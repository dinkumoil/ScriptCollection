'==============================================================================
' Klasse zum Parsen von JSON-Dateien
'
' Autor: Andreas Heim
' Datum: 30.06.2016
'
' Umsetzung: Endlicher deterministischer Automat (Finite State Machine)
'
'==============================================================================


Const EolStyleWindows    = 0
Const EolStyleUnix       = 1

Const IndentStyleSpace   = 0
Const IndentStyleTab     = 1

Const BracketStyleAllman = 0
Const BracketStyleJava   = 1


Class clsJsonFile
  'Variablen für die Zustände der Finite State Machine
  Private intContextLookup
  Private intStateError
  Private intStateAwaitKey
  Private intStateReadKey
  Private intStateKeyRead
  Private intStateAwaitValue
  Private intStateReadConst
  Private intStateConstRead
  Private intStateReadNumber
  Private intStateNumberRead
  Private intStateReadString
  Private intStateReadEscapeChar
  Private intStateStringRead
  Private intStateReadArray
  Private intStateArrayRead
  Private intStateReadObject
  Private intStateObjectRead

  'Arrays für die Zustandsdaten
  Private arrStateAwaitKey,   arrStateAwaitValue
  Private arrStateReadKey,    arrStateKeyRead
  Private arrStateReadConst,  arrStateConstRead
  Private arrStateReadNumber, arrStateNumberRead
  Private arrStateReadString, arrStateReadEscapeChar, arrStateStringRead
  Private arrStateReadArray,  arrStateArrayRead
  Private arrStateReadObject, arrStateObjectRead

  'Tabelle aller Zustandsdaten
  Private arrStateTable

  'Steuervariablen des Parsers
  Private intState
  Private intLCnt
  Private strItem, strKey

  'Allgemeine Variablen
  Private objFSO, objRegEx, dicJsonFile
  Private arrStack, intStackPtr
  Private varCurDataStore
  Private intDebugLevel
  Private intEolStyle
  Private intIndent
  Private intIndentStyle
  Private intBracketStyle
  Private bolAllStrValues


  '----------------------------------------------------------------------------
  'Constructor
  '----------------------------------------------------------------------------
  Private Sub Class_Initialize()
    'Einstellung des Debug-Levels
    '0, 1 oder 2 = keine, wenig und viele Debugausgaben
    intDebugLevel           = 0

    'Zeichen für Zeilenende
    intEolStyle             = EolStyleWindows

    'Stil für Position von öffnenden Klammern
    intBracketStyle         = BracketStyleAllman

    'Einrückung per Leerzeichen oder Tab
    intIndentStyle          = IndentStyleSpace

    'Einrückungstiefe für die Textausgabe (Anzahl Leerzeichen bzw. Tabs)
    intIndent               = 2

    'Legt fest ob alle Eigenschaften eines Objekts als String geschrieben werden
    bolAllStrValues         = False

    'Stack und Stackpointer
    arrStack                = Array()
    intStackPtr             = -1

    'Regular Expressions zum decodieren der Eingabedaten
    Set objRegEx            = New RegExp
    objRegEx.Global         = False
    objRegEx.IgnoreCase     = True

    'Datencontainer für das komplette JSON-File
    Set dicJsonFile         = CreateObject("Scripting.Dictionary")
    dicJsonFile.CompareMode = vbTextCompare

    'Objekt für Dateisystemoperationen
    Set objFSO              = CreateObject("Scripting.FileSystemObject")

    'Codes der Zustände
    'Der Code -1 entspricht einem Fehlerzustand und beendet die Verarbeitung
    'der Eingabedaten.
    'Der Code -2 wird verwendet, um Kontextabhängig darüber zu entscheiden,
    'ob der Folgezustand "Erwarte Schlüssel" oder "Erwarte Wert" sein soll.
    intContextLookup       = -2
    intStateError          = -1
    intStateAwaitKey       =  0
    intStateReadKey        =  1
    intStateKeyRead        =  2
    intStateAwaitValue     =  3
    intStateReadConst      =  4
    intStateConstRead      =  5
    intStateReadNumber     =  6
    intStateNumberRead     =  7
    intStateReadString     =  8
    intStateReadEscapeChar =  9
    intStateStringRead     = 10
    intStateReadArray      = 11
    intStateArrayRead      = 12
    intStateReadObject     = 13
    intStateObjectRead     = 14

    'Die Zustandstabellen haben folgendes Layout:
    '
    'Index 0:    String der eine Regular Expression enthalten kann, die zur
    '            Validierung eines vollständig eingelesenen Tokens verwendet wird
    '
    'Ab Index 1: Arrays mit zwei Elementen, die eine Beziehung zwischen Eingabe-
    '            zeichen und sich daraus ergebenden Folgezuständen herstellen.
    '
    '            Index 0: String mit Regular Expression, die die Eingabezeichen
    '                     codiert, die zum Übergang in den ersten Folgezustand
    '                     aus dem Array in Index 1 führen
    '
    '            Index 1: Array mit Folgezuständen, die als Reaktion auf das Ein-
    '                     lesen der im Index 0 codierten Eingabezeichen sequentiell
    '                     durchlaufen werden sollen.

    'Zustand "Erwarte Schlüssel"
    arrStateAwaitKey       = Array("", _
                                   Array("""",          Array(intStateReadKey)) _
                                  )

    'Zustand "Lese Schlüssel"
    arrStateReadKey        = Array("^[^""]*$", _
                                   Array("[^""]",       Array(intStateReadKey)), _
                                   Array("""",          Array(intStateKeyRead))  _
                                  )

    'Zustand "Schlüssel vollständig gelesen"
    arrStateKeyRead        = Array("", _
                                   Array(":",           Array(intStateAwaitValue)) _
                                  )

    'Zustand "Erwarte Wert"
    arrStateAwaitValue     = Array("", _
                                   Array("[ntf]",       Array(intStateReadConst)),  _
                                   Array("[0-9-]",      Array(intStateReadNumber)), _
                                   Array("""",          Array(intStateReadString)), _
                                   Array("\[",          Array(intStateReadArray)),  _
                                   Array("\]",          Array(intStateArrayRead)),  _
                                   Array("\{",          Array(intStateReadObject)), _
                                   Array("\}",          Array(intStateObjectRead))  _
                                  )

    'Zustand "Lese Konstante"
    arrStateReadConst      = Array("^(null|true|false)$", _
                                   Array("[a-z]",       Array(intStateReadConst)), _
                                   Array(",",           Array(intStateConstRead, intContextLookup)),  _
                                   Array("\]",          Array(intStateConstRead, intStateArrayRead)), _
                                   Array("\}",          Array(intStateConstRead, intStateObjectRead)) _
                                  )

    'Zustand "Konstante vollständig gelesen"
    arrStateConstRead      = Array("", _
                                   Array(",",           Array(intContextLookup)),  _
                                   Array("\]",          Array(intStateArrayRead)), _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zustand "Lese Zahl"
    arrStateReadNumber     = Array("^-{0,1}(0|([1-9][0-9]*))(\.[0-9]+)*(E(\+|-){0,1}[0-9]+)*$", _
                                   Array("[0-9\.E\+-]", Array(intStateReadNumber)), _
                                   Array(",",           Array(intStateNumberRead, intContextLookup)),   _
                                   Array("\]",          Array(intStateNumberRead, intStateArrayRead)),  _
                                   Array("\}",          Array(intStateNumberRead, intStateObjectRead))  _
                                  )

    'Zustand "Zahl vollständig gelesen"
    arrStateNumberRead     = Array("", _
                                   Array(",",           Array(intContextLookup)),  _
                                   Array("\]",          Array(intStateArrayRead)), _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zustand "Lese Zeichenkette"
    arrStateReadString     = Array("", _
                                   Array("[^\\""]",     Array(intStateReadString)),     _
                                   Array("\\",          Array(intStateReadEscapeChar)), _
                                   Array("""",          Array(intStateStringRead))      _
                                  )

    'Zustand "Lese Escape-Zeichen in Zeichenkette"
    arrStateReadEscapeChar = Array("", _
                                   Array("\\",          Array(intStateReadString)), _
                                   Array("/",           Array(intStateReadString)), _
                                   Array("""",          Array(intStateReadString)), _
                                   Array("b",           Array(intStateReadString)), _
                                   Array("f",           Array(intStateReadString)), _
                                   Array("n",           Array(intStateReadString)), _
                                   Array("r",           Array(intStateReadString)), _
                                   Array("t",           Array(intStateReadString)), _
                                   Array("u",           Array(intStateReadString))  _
                                  )

    'Zustand "Zeichenkette vollständig gelesen"
    arrStateStringRead     = Array("", _
                                   Array(",",           Array(intContextLookup)),  _
                                   Array("\]",          Array(intStateArrayRead)), _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zustand "Lese Array"
    arrStateReadArray      = Array("", _
                                   Array("[ntf]",       Array(intStateReadConst)),  _
                                   Array("[0-9-]",      Array(intStateReadNumber)), _
                                   Array("""",          Array(intStateReadString)), _
                                   Array("\[",          Array(intStateReadArray)),  _
                                   Array("\]",          Array(intStateArrayRead)),  _
                                   Array("\{",          Array(intStateReadObject)), _
                                   Array("\}",          Array(intStateObjectRead))  _
                                  )

    'Zustand "Array vollständig gelesen"
    arrStateArrayRead      = Array("", _
                                   Array(",",           Array(intContextLookup)),  _
                                   Array("\]",          Array(intStateArrayRead)), _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zustand "Lese Objekt"
    arrStateReadObject     = Array("", _
                                   Array("""",          Array(intStateReadKey)),   _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zustand "Objekt vollständig gelesen"
    arrStateObjectRead     = Array("", _
                                   Array(",",           Array(intContextLookup)),  _
                                   Array("\]",          Array(intStateArrayRead)), _
                                   Array("\}",          Array(intStateObjectRead)) _
                                  )

    'Zusammenfassung aller Zustandstabellen
    arrStateTable          = Array(arrStateAwaitKey,       _
                                   arrStateReadKey,        _
                                   arrStateKeyRead,        _
                                   arrStateAwaitValue,     _
                                   arrStateReadConst,      _
                                   arrStateConstRead,      _
                                   arrStateReadNumber,     _
                                   arrStateNumberRead,     _
                                   arrStateReadString,     _
                                   arrStateReadEscapeChar, _
                                   arrStateStringRead,     _
                                   arrStateReadArray,      _
                                   arrStateArrayRead,      _
                                   arrStateReadObject,     _
                                   arrStateObjectRead      _
                                  )
  End Sub


  '----------------------------------------------------------------------------
  'Destructor
  '----------------------------------------------------------------------------
  Private Sub Class_Terminate()
    Clear
    Set objFSO      = Nothing
    Set dicJsonFile = Nothing
  End Sub


  '----------------------------------------------------------------------------
  'Alle aus einem evtl. vorher eingelesenen JSON-File stammenden Daten verwerfen
  '----------------------------------------------------------------------------
  Private Sub Clear
    dicJsonFile.RemoveAll

    'Anfangszustand der Finite State Machine setzen
    Set varCurDataStore = dicJsonFile
    intState            = intStateAwaitValue
    intLCnt             = 0
    strItem             = ""
    strKey              = ""
  End Sub


  '----------------------------------------------------------------------------
  'Liefert die Root-Ebene der JSON-Datenstruktur (Array oder Dictionary)
  '----------------------------------------------------------------------------
  Public Property Get Content
    If IsObject(dicJsonFile.Item("")) Then
      Set Content = dicJsonFile.Item("")
    ElseIf IsArray(dicJsonFile.Item("")) Then
      Content = dicJsonFile.Item("")
    Else
      Content = NULL
    End If
  End Property


  '----------------------------------------------------------------------------
  'Liest/Setzt den Zeilenende-Stil für die Ausgabe der JSON-Daten
  '----------------------------------------------------------------------------
  Public Property Get EolStyle
    EolStyle = intEolStyle
  End Property


  Public Property Let EolStyle(intValue)
    intEolStyle = intValue
  End Property


  '----------------------------------------------------------------------------
  'Liest/Setzt den Klammer-Stil für die Ausgabe der JSON-Daten
  '----------------------------------------------------------------------------
  Public Property Get BracketStyle
    BracketStyle = intBracketStyle
  End Property


  Public Property Let BracketStyle(intValue)
    intBracketStyle = intValue
  End Property


  '----------------------------------------------------------------------------
  'Liest/Setzt den Einrückungs-Stil für die Ausgabe der JSON-Daten
  '----------------------------------------------------------------------------
  Public Property Get IndentStyle
    IndentStyle = intIndentStyle
  End Property


  Public Property Let IndentStyle(intValue)
    intIndentStyle = intValue
  End Property


  '----------------------------------------------------------------------------
  'Liest/Setzt die Einrückungstiefe für die Ausgabe der JSON-Daten
  '----------------------------------------------------------------------------
  Public Property Get Indent
    Indent = intIndent
  End Property


  Public Property Let Indent(intValue)
    intIndent = intValue
  End Property


  '----------------------------------------------------------------------------
  'Liest/Setzt ob bei der Ausgabe der JSON-Daten
  'alle Eigenschaften als String geschrieben werden
  '----------------------------------------------------------------------------
  Public Property Get AllStrValues
    AllStrValues = bolAllStrValues
  End Property


  Public Property Let AllStrValues(bolValue)
    bolAllStrValues = bolValue
  End Property


  '----------------------------------------------------------------------------
  'Erstellt eine Textrepräsentation der JSON-Datenstruktur und verdeutlicht
  'die Hierarchieebenen durch Einrückungen
  '----------------------------------------------------------------------------
  Public Function ToString(bolClean)
    If dicJsonFile.Count > 0 Then
      ToString = PrintJsonDic(dicJsonFile, "", bolClean)
    Else
      ToString = ""
    End If
  End Function


  '----------------------------------------------------------------------------
  'Speichert das Datenmodell als JSON-Datei
  '----------------------------------------------------------------------------
  Public Sub SaveToFile(ByRef strFilePath, intEncoding)
    Dim objOutStream

    Set objOutStream = objFSO.OpenTextFile(strFilePath, 2, True, intEncoding)
    objOutStream.Write ToString(True)
    objOutStream.Close
  End Sub


  '----------------------------------------------------------------------------
  'Lädt einen JSON-String und parst die Datenstruktur
  '----------------------------------------------------------------------------
  Public Function LoadFromString(ByRef strJsonString)
    Dim arrJsonString, intIdx

    'Evtl. bereits vorhandenes JSON-Datenmodell verwerfen
    Clear

    'Prüfen ob der übergebene String nicht leer ist
    If strJsonString = "" Then
      LoadFromString = False
      Exit Function
    End If

    'JSON-String in Zeilen aufteilen
    arrJsonString = Split(strJsonString, EolStr())

    'Alle Zeilen des Strings parsen
    For intIdx = 0 To UBound(arrJsonString)
      intLCnt = intLCnt + 1
      If Not ParseLine(arrJsonString(intIdx)) Then Exit For
    Next

    'Wenn der Endzustand "Objekt gelesen" und der Stack leer ist, war das
    'Einlesen des JSON-Strings erfolgreich
    LoadFromString = (intState = intStateObjectRead Or intState = intStateArrayRead) And _
                     intStackPtr = -1
  End Function


  '----------------------------------------------------------------------------
  'Lädt ein JSON-File und parst die Datenstruktur
  '----------------------------------------------------------------------------
  Public Function LoadFromFile(ByRef strFilePath, intEncoding)
    Dim objInStream

    'Evtl. bereits vorhandenes JSON-Datenmodell verwerfen
    Clear

    'Existenz der zu ladenden Datei prüfen
    If Not objFSO.FileExists(strFilePath) Then
      LoadFromFile = False
      Exit Function
    End If

    'JSON-File öffnen
    Set objInStream = objFSO.OpenTextFile(strFilePath, 1, False, intEncoding)

    'Alle Zeilen der Datei parsen
    Do While Not objInStream.AtEndOfStream
      intLCnt = intLCnt + 1
      If Not ParseLine(objInStream.ReadLine) Then Exit Do
    Loop

    'JSON-File schließen
    objInStream.Close

    'Wenn der Endzustand "Objekt gelesen" und der Stack leer ist, war das
    'Einlesen des JSON-Files erfolgreich
    LoadFromFile = (intState = intStateObjectRead Or intState = intStateArrayRead) And _
                   intStackPtr = -1
  End Function


  '----------------------------------------------------------------------------
  'Parst JSON-Daten in einem String und erstellt ein Datenmodell
  '----------------------------------------------------------------------------
  Private Function ParseLine(ByRef strLine)
    Dim intLineLen, intCCnt, strChar
    Dim arrStateData, intCnt, intStateCnt
    Dim strCheck, intNewState
    Dim varParentDataStore

    intLineLen = Len(strLine)

    'Über alle Zeichen einer Zeile iterieren
    For intCCnt = 1 To intLineLen
      'Ein einzelnes Zeichen aus der Zeile extrahieren und
      'die aktuelle Zustandstabelle laden
      strChar      = Mid(strLine, intCCnt, 1)
      arrStateData = arrStateTable(intState)

      'Tabulatorzeichen werden nie verarbeitet
      If strChar <> vbTab Then
        'Leerzeichen werden nur verarbeitet, wenn gerade ein String eingelesen wird
        If strChar <> " " Or intState = intStateReadKey Or intState = intStateReadString Then
          DebugOutput 1, strLine, strChar, intState

          'Überprüfen ob das aktuelle Zeichen auf eines der Suchmuster aus der
          'aktuellen Zustandstabelle passt, wenn ja die Schleife abbrechen
          For intCnt = 1 To UBound(arrStateData)
            objRegEx.Pattern = arrStateData(intCnt)(0)
            If objRegEx.Test(strChar) Then Exit For
          Next

          'Wenn ein Treffer gefunden wurde...
          If intCnt <= UBound(arrStateData) Then
            '...alle vorgegebenen Folgezustände sequentiell durchlaufen
            For intStateCnt = 0 To UBound(arrStateData(intCnt)(1))
              'Den Folgezustand ermitteln
              intNewState = arrStateData(intCnt)(1)(intStateCnt)

              'Wenn der neue Zustand nur unter Berücksichtigung des aktuellen
              'Kontexts (Einlesen eines Array oder eines Objekts) ermittelt
              'werden kann, jetzt den neuen Zustand endgültig bestimmen
              If intNewState = intContextLookup Then
                intNewState = AwaitStateByContext(varCurDataStore)
              End If

              'Je nach ermitteltem Folgezustand verschiedene Prüfungen und
              'Aktionen ausführen, bevor der Zustand gewechselt wird
              Select Case intNewState
                'Folgezustand: Schlüssel einlesen
                Case intStateReadKey
                  'Dieser Token-Typ wird durch ein Anführungszeichen eingeleitet,
                  'das nicht zum Token gehört. Deshalb erst dann Zeichen in das
                  'Token übernehmen, wenn der aktuelle Zustand bereits "Lese
                  'Schlüssel" ist.
                  If intNewState = intState Then
                    strItem = strItem & strChar
                  End If

                'Folgezustand: Konstanten- bzw. Zahlwert einlesen
                Case intStateReadConst, _
                     intStateReadNumber
                  'Für diese beiden Token-Typen existiert kein spezielles Einleitungs-
                  'zeichen, das aktuelle Zeichen kann deshalb direkt übernommen werden.
                  strItem  = strItem & strChar

                'Folgezustand: Zeichenkette einlesen
                Case intStateReadString
                  'Dieser Token-Typ wird durch ein Anführungszeichen eingeleitet,
                  'das nicht zum Token gehört. Deshalb erst dann Zeichen in das
                  'Token übernehmen, wenn der aktuelle Zustand bereits "Lese
                  'Zeichenkette" ist. Wenn der aktuelle Zustand "Escape-Zeichen
                  'in Zeichenkette einlesen" ist, kann das Zeichen ebenfalls
                  'übernommen werden
                  If intNewState = intState Or _
                     intState    = intStateReadEscapeChar Then
                    strItem = strItem & strChar
                  End If

                'Folgezustand: Escape-Zeichen in Zeichenkette einlesen
                Case intStateReadEscapeChar
                  'Das Einleitungszeichen für ein Escape-Zeichen gehört zum Token
                  'und kann deshalb direkt übernommen werden.
                  strItem  = strItem & strChar

                'Folgezustand: Array einlesen
                Case intStateReadArray
                  'Der aktuelle Datenspeicher (Array oder Objekt) wird zusammen mit
                  'dem Namen des neuen Arrays auf dem Stack abgelegt, damit evtl.
                  'noch weitere Daten darin gespeichert werden können, nachdem das
                  'neue Array vollständig eingelesen wurde.
                  Call StackPush(strKey, varCurDataStore)

                  'Neues Array als Datenspeicher anlegen
                  varCurDataStore = Array()

                  'Der Name des neuen Arrays ist jetzt "verbrannt", der String muss
                  'deshalb gelöscht werden, damit die weitere Verarbeitung fehlerfrei
                  'abläuft
                  strKey          = ""

                'Folgezustand: Objekt einlesen
                Case intStateReadObject
                  'Der aktuelle Datenspeicher (Array oder Objekt) wird zusammen mit
                  'dem Namen des neuen Objekts auf dem Stack abgelegt, damit evtl.
                  'noch weitere Daten darin gespeichert werden können, nachdem das
                  'neue Objekt vollständig eingelesen wurde.
                  Call StackPush(strKey, varCurDataStore)

                  'Neues Objekt als Datenspeicher anlegen
                  Set varCurDataStore         = CreateObject("Scripting.Dictionary")
                  varCurDataStore.CompareMode = vbTextCompare

                  'Der Name des neuen Objekts ist jetzt "verbrannt", der String muss
                  'deshalb gelöscht werden, damit die weitere Verarbeitung fehlerfrei
                  'abläuft
                  strKey                      = ""

                'Folgezustand: Schlüssel, Konstante, Zahl oder String
                '              wurde vollständig eingelesen
                Case intStateKeyRead, _
                     intStateConstRead, _
                     intStateNumberRead, _
                     intStateStringRead
                  'Die Regular Expression zur Validierung des Tokens laden
                  strCheck = arrStateData(0)

                  'Wenn es keine Regular Expression zur Validierung des Tokens
                  'gibt, kann der Folgezustand einfach übernommen werden.
                  'Ansonsten die Validierung ausführen
                  If strCheck <> "" Then
                    objRegEx.Pattern = strCheck

                    If Not objRegEx.Test(strItem) Then
                      DebugOutput 2, strLine, strChar, intNewState
                      intNewState = intStateError
                    End If
                  End If

                  'Wenn der Folgezustand übernommen werden kann, das aktuelle
                  'Token speichern
                  Select Case intNewState
                    Case intStateKeyRead
                      'Den Schlüsselnamen speichern und den Einlesepuffer löschen
                      If strItem <> "" Then
                        strKey  = strItem
                        strItem = ""

                      Else
                        DebugOutput 3, strLine, strChar, intNewState
                        intNewState = intStateError
                      End If

                    Case intStateConstRead, _
                         intStateNumberRead, _
                         intStateStringRead
                      'Wenn zuvor ein Schlüsselname eingelesen wurde oder gerade
                      'die Elemente eines Arrays eingelesen werden (die keinem
                      'Schlüssel zugeordnet sind) ...
                      If strKey <> "" Or IsArrayContext(varCurDataStore) Then
                        '...wird der aktuelle Wert im aktuellen Datenspeicher abgelegt
                        If StoreData(varCurDataStore, strKey, strItem) Then
                          'Den "verbrannten" Schlüsselnamen und den Einlesepuffer löschen
                          strKey  = ""
                          strItem = ""

                        'Falls das Speichern des Wertes fehlgeschlagen ist, wird ein
                        'Fehlerzustand gesetzt, der zum Abbruch der Verarbeitung führt
                        Else
                          DebugOutput 4, strLine, strChar, intNewState
                          intNewState = intStateError
                        End If

                      Else
                        DebugOutput 5, strLine, strChar, intNewState
                        intNewState = intStateError
                      End If
                  End Select

                'Folgezustand: Array oder Objekt wurde vollständig eingelesen
                Case intStateArrayRead, _
                     intStateObjectRead
                  'Den neuen Zustand nur übernehmen, wenn der Stack nicht leer ist
                  'Ansonsten einen Fehlerzustand setzen, der zum Abbruch der Ver-
                  'arbeitung führt
                  If intStackPtr >= 0 Then
                    'Den Schlüsselnamen sowie den Datenspeicher, denen das Array bzw.
                    'das Objekt zugeordnet sind, vom Stack holen
                    Call StackPop(strKey, varParentDataStore)

                    'Wenn der Schlüsselname nicht leer ist oder es sich bei dem Parent-
                    'Datenspeicher um ein Array oder um das Root-Objekt der JSON-Daten-
                    'struktur handelt ...
                    If strKey <> "" Or IsArrayContext(varParentDataStore) _
                                    Or IsSameObject(varParentDataStore, dicJsonFile) Then
                      '...wird der aktuelle Wert im Parent-Datenspeicher abgelegt
                      If StoreData(varParentDataStore, strKey, varCurDataStore) Then
                        'Jetzt den Parent-Datenspeicher zum aktuellen Datenspeicher machen
                        If IsObject(varParentDataStore) Then
                          Set varCurDataStore = varParentDataStore
                        Else
                          varCurDataStore = varParentDataStore
                        End If

                        'Den "verbrannten" Schlüsselnamen und den Einlesepuffer löschen
                        strKey  = ""
                        strItem = ""

                      'Falls das Speichern des Wertes fehlgeschlagen ist, wird ein
                      'Fehlerzustand gesetzt, der zum Abbruch der Verarbeitung führt
                      Else
                        DebugOutput 6, strLine, strChar, intNewState
                        intNewState = intStateError
                      End If
                    End If

                  Else
                    DebugOutput 7, strLine, strChar, intNewState
                    intNewState = intStateError
                  End If
              End Select

              DebugOutput 8, strLine, strChar, intNewState
            Next

          'Wenn das aktuell eingelesene Zeichen auf keines der Suchmuster aus der
          'aktuellen Zustandstabelle passt, liegt ein Formatfehler vor. Die weitere
          'Verarbeitung durch Setzen eines Fehlerzustands abbrechen
          Else
            DebugOutput 9, strLine, strChar, intNewState
            intNewState = intStateError
          End If

          intState = intNewState

          'Wenn kein Fehlerzustand gesetzt ist, den neuen Zustand übernehmen.
          'Bei einem Fehlerzustand die betroffene Zeile und die Zeichennummer
          'ausgeben und abbrechen
          If intState = intStateError Then
            PrintErrorMessage intLCnt, intCCnt
            Exit For
          End If
        End If
      End If
    Next

    'Bei einem Fehlerzustand abbrechen
    ParseLine = (intState <> intStateError)
  End Function


  '----------------------------------------------------------------------------
  'Daten (ggf. in Form von Schlüssel/Wert-Paaren) in einem Datenspeicher ablegen
  '----------------------------------------------------------------------------
  Private Function StoreData(ByRef varDataStore, ByRef varKey, ByRef varValue)
    StoreData = False

    'Beim Speichern von Werten muss unterschieden werden, ob es sich bei dem
    'Zieldatenspeicher um ein Dictionary oder um ein Array handelt

    'Datenspeicher ist ein Dictionary
    If IsObject(varDataStore) Then
      'Im Dictionary kann der Wert nur gespeichert werden, wenn der zugehörige
      'Schlüssel noch nicht existiert
      If Not varDataStore.Exists(varKey) Then
        'Debug-Ausgaben
        If intDebugLevel > 0 then
          If IsObject(varValue) Then
            WScript.Echo "Add to Object: " & varKey & "=Object" & vbNewLine
          ElseIf IsArray(varValue) Then
            WScript.Echo "Add to Object: " & varKey & "=Array" & vbNewLine
          Else
            WScript.Echo "Add to Object: " & varKey & "=" & varValue & vbNewLine
          End If
        End If

        'Schlüssel und Wert speichern
        Call varDataStore.Add(varKey, varValue)
        StoreData = True
      End If

    'Datenspeicher ist ein Array
    ElseIf IsArray(varDataStore) Then
      'Das Array muss zuerst vergrößert werden
      ReDim Preserve varDataStore(UBound(varDataStore) + 1)

      'Debug-Ausgaben
      If intDebugLevel > 0 then
        If IsObject(varValue) Then
          WScript.Echo "Add to Array[" & UBound(varDataStore) & "]: Object" & vbNewLine
        ElseIf IsArray(varValue) Then
          WScript.Echo "Add to Array[" & UBound(varDataStore) & "]: Array" & vbNewLine
        Else
          WScript.Echo "Add to Array[" & UBound(varDataStore) & "]: " & varValue & vbNewLine
        End If
      End If

      'Hier muss unterschieden werden, ob es sich bei dem zu speichernden Wert
      'um ein Objekt oder um ein Array bzw. einen skalaren Wert handelt
      If IsObject(varValue) Then
        Set varDataStore(UBound(varDataStore)) = varValue
      Else
        varDataStore(UBound(varDataStore)) = varValue
      End If

      StoreData = True
    End If
  End Function


  '----------------------------------------------------------------------------
  'Schlüsselname und Parent-Datenspeicher auf dem Stack ablegen
  '----------------------------------------------------------------------------
  Private Sub StackPush(ByRef strName, ByRef varDataStore)
    intStackPtr = intStackPtr + 1

    'Stack-Array vergrößern
    ReDim Preserve arrStack(intStackPtr)
    arrStack(intStackPtr) = Array("", NULL)

    'Datenspeicher ist ein Dictionary
    If IsObject(varDataStore) Then
      'Debug-Ausgaben
      If intDebugLevel > 0 Then
        WScript.Echo "Push Name: " & strName
        WScript.Echo "Push Parent DataStore: Object" & vbNewLine
      End If

      'Daten speichern
      arrStack(intStackPtr)(0)     = strName
      Set arrStack(intStackPtr)(1) = varDataStore

    'Datenspeicher ist ein Array
    ElseIf IsArray(varDataStore) Then
      'Debug-Ausgaben
      If intDebugLevel > 0 Then
        WScript.Echo "Push Name: " & strName
        WScript.Echo "Push Parent DataStore: Array" & vbNewLine
      End If

      'Daten speichern
      arrStack(intStackPtr)(0) = strName
      arrStack(intStackPtr)(1) = varDataStore
    End If
  End Sub


  '----------------------------------------------------------------------------
  'Schlüsselname und Parent-Datenspeicher vom Stack holen
  '----------------------------------------------------------------------------
  Private Sub StackPop(ByRef strName, ByRef varDataStore)
    'Operation nur ausführen, wenn der Stack nicht leer ist
    If intStackPtr >= 0 Then
      'Schlüsselname holen
      strName = arrStack(intStackPtr)(0)

      'Datenspeicher ist ein Dictionary
      If IsObject(arrStack(intStackPtr)(1)) Then
        'Debug-Ausgaben
        If intDebugLevel > 0 Then
          WScript.Echo "Pop Name: " & strName
          WScript.Echo "Pop Parent DataStore: Object" & vbNewLine
        End If

        'Datenspeicher holen
        Set varDataStore = arrStack(intStackPtr)(1)

      'Datenspeicher ist ein Array
      ElseIf IsArray(arrStack(intStackPtr)(1)) Then
        'Debug-Ausgaben
        If intDebugLevel > 0 Then
          WScript.Echo "Pop Name: " & strName
          WScript.Echo "Pop Parent DataStore: Array" & vbNewLine
        End If

        'Datenspeicher holen
        varDataStore = arrStack(intStackPtr)(1)
      End If

      'Stackpointer korrigieren
      intStackPtr = intStackPtr - 1

    'Bei leerem Stack nichtinitialisierte Variablen zurückliefern
    Else
      strName      = NULL
      varDataStore = NULL
    End If
  End Sub


  '----------------------------------------------------------------------------
  'Aus dem Kontext (verwendeter Datenspeicher) ableiten, ob als nächstes ein
  'Schlüssel oder ein Wert erwartet wird
  '----------------------------------------------------------------------------
  Private Function AwaitStateByContext(ByRef varDataStore)
    If IsArrayContext(varDataStore) Then
      AwaitStateByContext = intStateAwaitValue
    Else
      AwaitStateByContext = intStateAwaitKey
    End If
  End Function


  '----------------------------------------------------------------------------
  'Ermittelt ob der übergebene Datenspeicher ein Array ist
  '----------------------------------------------------------------------------
  Private Function IsArrayContext(ByRef varDataStore)
    IsArrayContext = IsArray(varDataStore)
  End Function


  '----------------------------------------------------------------------------
  'Ermittelt ob zwei Objektreferenzen auf das gleiche Objekt zeigen
  '----------------------------------------------------------------------------
  Private Function IsSameObject(varObj1, varObj2)
    If IsObject(varObj1) And IsObject(varObj2) Then
      IsSameObject = (varObj1 Is varObj2)
    Else
      IsSameObject = False
    End If
  End Function


  '----------------------------------------------------------------------------
  'Wandelt ein JSON-Objekt in seine Textdarstellung um
  '----------------------------------------------------------------------------
  Private Function PrintJsonDic(ByRef dicDataStore, strIndent, bolClean)
    Dim strKey, varValue, strLine

    strLine = ""

    For Each strKey In dicDataStore.Keys
      If strLine <> "" Then
        strLine = strLine & "," & EolStr()
      End If

      If IsObject(dicDataStore.Item(strKey)) Then
        Set varValue = dicDataStore.Item(strKey)
      Else
        varValue = dicDataStore.Item(strKey)
      End If

      If IsObject(varValue) Then
        Select Case intBracketStyle
          Case BracketStyleAllman:
            If strKey <> "" Then
              strLine = strLine & strIndent & """" & strKey & """:" & EolStr()
            End If

            strLine = strLine & strIndent & "{" & EolStr() & _
                                PrintJsonDic(varValue, strIndent & IndentStr(), bolClean) & _
                                strIndent & "}"

          Case Else  'BracketStyleJava:
            If strKey <> "" Then
              strLine = strLine & strIndent & """" & strKey & """: {" & EolStr()
            Else
              strLine = strLine & strIndent & "{" & EolStr()
            End If

            strLine = strLine & PrintJsonDic(varValue, strIndent & IndentStr(), bolClean) & _
                                strIndent & "}"
        End Select

      ElseIf IsArray(varValue) Then
        Select Case intBracketStyle
          Case BracketStyleAllman:
            If strKey <> "" Then
              strLine = strLine & strIndent & """" & strKey & """: " & EolStr()
            Else
              strLine = ""
            End If

            strLine = strLine & strIndent & "[" & EolStr() & _
                                PrintJsonArr(varValue, strIndent & IndentStr(), bolClean) & _
                                strIndent & "]"

          Case Else  'BracketStyleJava:
            If strKey <> "" Then
              strLine = strLine & strIndent & """" & strKey & """: [" & EolStr()
            Else
              strLine = strLine & strIndent & "[" & EolStr()
            End If

            strLine = strLine & PrintJsonArr(varValue, strIndent & IndentStr(), bolClean) & _
                                strIndent & "]"
        End Select

      ElseIf IsNumeric(varValue) And Not bolAllStrValues Then
        strLine = strLine & strIndent & """" & strKey & """: " & varValue

      ElseIf LCase(varValue) = "null"  Or _
             LCase(varValue) = "true"  Or _
             LCase(varValue) = "false" Then
        strLine = strLine & strIndent & """" & strKey & """: " & varValue

      Else
        strLine = strLine & strIndent & """" & strKey & """" & ": """ & varValue & """"
      End If
    Next

    If strLine <> "" Then strLine = strLine & EolStr()

    PrintJsonDic = strLine
  End Function


  '----------------------------------------------------------------------------
  'Wandelt ein JSON-Array in seine Textdarstellung um
  '----------------------------------------------------------------------------
  Private Function PrintJsonArr(ByRef arrDataStore, strIndent, bolClean)
    Dim intCnt, varValue, strLine
    Dim intLocalIndentLen, strLocalIndent

    intCnt  = 0
    strLine = ""

    intLocalIndentLen = Len(CStr(UBound(arrDataStore))) + 4

    If Not bolClean Then
      strLocalIndent  = String(intLocalIndentLen, " ")
    Else
      strLocalIndent  = ""
    End If

    For Each varValue In arrDataStore
      If strLine <> "" Then
        strLine = strLine & "," & EolStr()
      End If

      If IsObject(varValue) Then
        strLine = strLine & strIndent & JsonArrPrefix(bolClean, intLocalIndentLen, intCnt) & "{" & EolStr() & _
                            PrintJsonDic(varValue, strIndent & strLocalIndent & IndentStr(), bolClean) & _
                            strIndent & strLocalIndent & "}"

      ElseIf IsArray(varValue) Then
        strLine = strLine & strIndent & JsonArrPrefix(bolClean, intLocalIndentLen, intCnt) & "[" & EolStr() & _
                            PrintJsonArr(varValue, strIndent & strLocalIndent & IndentStr(), bolClean) & _
                            strIndent & strLocalIndent & "]"

      ElseIf IsNumeric(varValue) And Not bolAllStrValues Then
        strLine = strLine & strIndent & JsonArrPrefix(bolClean, intLocalIndentLen, intCnt) & varValue

      ElseIf LCase(varValue) = "null"  Or _
             LCase(varValue) = "true"  Or _
             LCase(varValue) = "false" Then
        strLine = strLine & strIndent & JsonArrPrefix(bolClean, intLocalIndentLen, intCnt) & varValue

      Else
        strLine = strLine & strIndent & JsonArrPrefix(bolClean, intLocalIndentLen, intCnt) & """" & varValue & """"
      End If

      intCnt = intCnt + 1
    Next

    If strLine <> "" Then strLine = strLine & EolStr()

    PrintJsonArr = strLine
  End Function


  '----------------------------------------------------------------------------
  'Erzeugt ein EOL-Zeichen entsprechend dem eingestellten Zeilenende-Stil
  '----------------------------------------------------------------------------
  Private Function EolStr()
    Select Case intEolStyle
      Case EolStyleWindows
        EolStr = vbCrLf

      Case Else  'EolStyleUnix
        EolStr = vbLf
    End Select
  End Function


  '----------------------------------------------------------------------------
  'Erzeugt eine Einrückung entsprechend dem eingestellten Einrückungs-Stil
  '----------------------------------------------------------------------------
  Private Function IndentStr()
    Select Case intIndentStyle
      Case IndentStyleSpace
        IndentStr = String(intIndent, " ")

      Case Else  'IndentStyleTab
        IndentStr = String(intIndent, vbTab)
    End Select
  End Function


  '----------------------------------------------------------------------------
  'Erzeugt bei Anforderung den Index eines JSON-Arrayelements für dessen Text-
  'darstellung und berücksichtigt dabei die größere notwendige Einrückung
  '----------------------------------------------------------------------------
  Private Function JsonArrPrefix(bolClean, intBasePadding, intCnt)
    If Not bolClean Then
      JsonArrPrefix = "[" & intCnt & "]:" & String(intBasePadding - 3 - Len(CStr(intCnt)), " ")
    Else
      JsonArrPrefix = ""
    End If
  End Function


  '----------------------------------------------------------------------------
  'Ausgabe einer Fehlermeldung mit Zeilen- und Zeichennummer
  '----------------------------------------------------------------------------
  Private Sub PrintErrorMessage(intLCnt, intCCnt)
    WScript.Echo "Error in line " & intLCnt & ", character " & intCCnt
  End Sub


  '----------------------------------------------------------------------------
  'Ausgabe von Debug-Informationen
  '----------------------------------------------------------------------------
  Private Sub DebugOutput(intId, ByRef strLine, strChar, intState)
    If intDebugLevel > 1 Then
      WScript.Echo intId
      WScript.Echo "Line:  " & strLine
      WScript.Echo "Char:  " & strChar
      WScript.Echo "State: " & intState
      WScript.Echo
    End If
  End Sub
End Class
