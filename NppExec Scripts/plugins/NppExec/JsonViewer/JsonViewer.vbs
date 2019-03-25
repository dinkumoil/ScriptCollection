'Alle Variablen müssen vor der ersten Verwendung mit Dim deklariert werden
Option Explicit


'Externen Code in das Script einbinden
Include ".\IO.vbs"
Include ".\ClassJsonFile.vbs"


'===============================================================================
' Hauptprogramm
'===============================================================================

'Variablendeklaration
Dim objJsonFile

'Initialisierung
Set objJsonFile = New clsJsonFile

'Abbruch wenn zu wenig Parameter übergeben wurden
If WScript.Arguments.Count < 1 Then
  WScript.StdErr.WriteLine "Fehlender Parameter"
  WScript.Quit 1
End If

'Abbruch wenn beim Einlesen der JSON-Datei oder
'der Erstellung des Datenmodells ein Fehler auftritt
If Not objJsonFile.LoadFromFile(WScript.Arguments(0), AsAnsi) Then
  WScript.StdErr.WriteLine "Ungültiges Dateiformat"
  WScript.Quit 2
End If

'Aus dem Datenmodell eine formatierte Textdarstellung erzeugen und ausgeben
WScript.StdOut.WriteLine objJsonFile.ToString(True)

'Beenden
WScript.Quit 0



'===============================================================================
' Routine zum Einbinden von externem Code
'===============================================================================

Sub Include(ByRef strFilePath)
  Dim objFSO, objFileStream, strScriptFilePath, strAbsFilePath, strCode

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  strScriptFilePath = objFSO.GetParentFolderName(WScript.ScriptFullName)
  strAbsFilePath    = objFSO.GetAbsolutePathName(objFSO.BuildPath(strScriptFilePath, strFilePath))

  Set objFileStream = objFSO.OpenTextFile(strAbsFilePath, 1, False, 0)
  strCode           = objFileStream.ReadAll
  objFileStream.Close

  ExecuteGlobal strCode
End Sub
