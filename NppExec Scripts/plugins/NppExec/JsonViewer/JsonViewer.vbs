Option Explicit


'Include external code modules
Include ".\IO.vbs"
Include ".\ClassJsonFile.vbs"


'===============================================================================
' Main program
'===============================================================================

'Variables declaration
Dim objJsonFile

'Variables initialization
Set objJsonFile = New clsJsonFile

'Break on missing parameter
If WScript.Arguments.Count < 1 Then
  WScript.StdErr.WriteLine "Missing parameter"
  WScript.Quit 1
End If

'Break if reading the JSON file or createing the data model fails
If Not objJsonFile.LoadFromFile(WScript.Arguments(0), AsAnsi) Then
  WScript.StdErr.WriteLine "Invalid file format"
  WScript.Quit 2
End If

'Turn data model of JSON file into a well formatted text and send it to StdOut
WScript.StdOut.WriteLine objJsonFile.ToString(True)

'Quit
WScript.Quit 0



'===============================================================================
' Routine for including external code modules
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
