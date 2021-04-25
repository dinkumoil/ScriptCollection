Option Explicit


'===============================================================================
' Base configuration
'===============================================================================

'Variables declaration
Dim objFSO, objWshShell
Dim strPluginDirPath, strNppPluginDirPath


'Variables initalization
Set objFSO      = CreateObject("Scripting.FileSystemObject")
Set objWshShell = CreateObject("WScript.Shell")


'-------------------------------------------------------------------------------
' Retrieve commandline arguments
'-------------------------------------------------------------------------------
If Not ParseCommandline() Then WScript.Quit 1


'-------------------------------------------------------------------------------
' Check arguments
'-------------------------------------------------------------------------------
If Not objFSO.FolderExists(strPluginDirPath)    Then WScript.Quit 2
If Not objFSO.FolderExists(strNppPluginDirPath) Then WScript.Quit 3


'-------------------------------------------------------------------------------
' Install plugin and delete downloaded files
'-------------------------------------------------------------------------------
objWshShell.Run "xcopy.exe " & _
                  "/reiscqy " & _
                  Quote(strPluginDirPath) & " " & _
                  Quote(strNppPluginDirPath), _
                0, _
                True

objWshShell.Run "cmd.exe " & _
                  "/c ""rd /s /q " & _
                  Quote(strPluginDirPath) & """", _
                0, _
                True



'===============================================================================
' Retrieve commandline arguments
'===============================================================================

Function ParseCommandLine
  strPluginDirPath    = ""
  strNppPluginDirPath = ""

  If WScript.Arguments.Named.Exists("P") Then
    strPluginDirPath = WScript.Arguments.Named("P")
  End If

  If WScript.Arguments.Named.Exists("N") Then
    strNppPluginDirPath = WScript.Arguments.Named("N")
  End If

  ParseCommandLine = (strPluginDirPath <> "" And strNppPluginDirPath <> "")
End Function


'===============================================================================
' Surround a string with double quotes
'===============================================================================

Function Quote(ByRef strString)
  Quote = """" & strString & """"
End Function
