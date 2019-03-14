'===============================================================================
' FilenameMatch-Klasse
'===============================================================================

Class clsFileNameMatch
  Private objFSO


  Private Sub Class_Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
  End Sub


  Private Sub Class_Terminate()
    Set objFSO = Nothing
  End Sub


  Private Function FindMatch(ByRef strString, ByRef strPattern, ByVal intIdxStr, ByVal intIdxPattern)
    Dim chrString, chrPattern

    While (intIdxStr <= Len(strString)) And (intIdxPattern <= Len(strPattern))
      chrString  = Mid(strString, intIdxStr, 1)
      chrPattern = Mid(strPattern, intIdxPattern, 1)

      If chrPattern <> "*" Then
        FindMatch = (UCase(chrString) = UCase(chrPattern) Or chrPattern = "?")
        If Not FindMatch Then Exit Function

        intIdxStr     = intIdxStr + 1
        intIdxPattern = intIdxPattern + 1
      Else
        FindMatch = (intIdxPattern = Len(strPattern))
        If FindMatch Then Exit Function

        Do
          FindMatch = FindMatch(strString, strPattern, intIdxStr, intIdxPattern + 1)
          intIdxStr = intIdxStr + 1
        Loop Until FindMatch Or intIdxStr > Len(strString)

        Exit Function
      End If
    Wend

    If (intIdxStr <= Len(strString)) Or (intIdxPattern <= Len(strPattern)) Then FindMatch = False
    If (intIdxStr > Len(strString)) And (intIdxPattern = Len(strPattern)) And (Right(strPattern, 1) = "*") Then FindMatch = True
  End Function


  Public Function FileNameMatch(ByRef strFileName, ByRef strPattern)
    Dim strName, strExt, strPatternName, strPatternExt

    strName = objFSO.GetBaseName(strFileName)
    strExt  = objFSO.GetExtensionName(strFileName)

    strPatternName = objFSO.GetBaseName(strPattern)
    strPatternExt  = objFSO.GetExtensionName(strPattern)

    If strPatternName = "" Then strPatternName = "*"
    If strPatternExt  = "" Then strPatternExt  = "*"

    FileNameMatch = FindMatch(strName, strPatternName, 1, 1) And FindMatch(strExt, strPatternExt, 1, 1)
  End Function
End Class
