strAlertFile    = "E:\ProcessCreationLog.txt"

ForReading      = 1
ForWriting      = 2
ForAppending    = 8

CreateNew       = True

AsASCII         = 0
AsUnicode       = -1
AsSystemDefault = -2


Set listArgs = WScript.Arguments

If listArgs.Count > 0 Then
  Set FSO = CreateObject("Scripting.FileSystemObject")

  Set AlertFile = FSO.OpenTextFile(strAlertFile, ForAppending, CreateNew, AsASCII)
  AlertFile.WriteLine("Process ""cmd.exe"" has been created with PID: " & listArgs(0))
  AlertFile.Close
  Set AlertFile = Nothing

  Set FSO = Nothing
End If

Set listArgs = Nothing
