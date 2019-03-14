'===============================================================================
' Ini file class
'===============================================================================

Class clsIniFile
  Private objFSO
  Private dicIniFile


  Private Sub Class_Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set dicIniFile         = CreateObject("Scripting.Dictionary")
    dicIniFile.CompareMode = vbTextCompare
  End Sub


  Private Sub Class_Terminate()
    Clear
    Set objFSO     = Nothing
    Set dicIniFile = Nothing
  End Sub


  Private Sub Clear
    dicIniFile.RemoveAll
  End Sub


  Private Function IsArrayDimed(ByRef arrArray, intDimension)
    Dim intUBound

    If Not IsArray(arrArray) Or intDimension < 1 Then
      IsArrayDimed = False
      Exit Function
    End If

    On Error Resume Next

    intUBound = UBound(arrArray, intDimension)
    IsArrayDimed = (Err.Number = 0)

    On Error Goto 0
  End Function


  Public Function LoadFile(strFilePath, intEncoding)
    Dim objInStream
    Dim objRegEx, colMatches
    Dim strLine, strSection, strKey, strValue
    Dim dicSection

    Clear

    strFilePath = objFSO.GetAbsolutePathName(strFilePath)

    If Not objFSO.FileExists(strFilePath) Then
      LoadFile = False
      Exit Function
    End If

    Set objRegEx        = New RegExp
    objRegEx.Global     = False
    objRegEx.IgnoreCase = True

    Set objInStream = objFSO.OpenTextFile(strFilePath, 1, False, intEncoding)

    Do While Not objInStream.AtEndOfStream
      strLine = objInStream.ReadLine

      objRegEx.Pattern = "^\[(.+)\]$"
      Set colMatches   = objRegEx.Execute(strLine)

      If colMatches.Count > 0 Then
        strSection = colMatches(0).SubMatches(0)
        Call AddSection(strSection)
      ElseIf strSection <> "" Then
        objRegEx.Pattern = "^([^;].+)=(.*)$"
        Set colMatches   = objRegEx.Execute(strLine)

        If colMatches.Count > 0 Then
          strKey   = colMatches(0).SubMatches(0)
          strValue = colMatches(0).SubMatches(1)
          Call AddKeyValue(strSection, strKey, strValue)
        End If
      End If
    Loop

    objInStream.Close

    LoadFile = True
  End Function


  Public Function SaveFile(strFilePath, intEncoding, bolOverwrite)
    Dim objOutStream
    Dim strSection, strKey

    strFilePath = objFSO.GetAbsolutePathName(strFilePath)

    If objFSO.FileExists(strFilePath) And Not bolOverwrite Then
      SaveFile = False
      Exit Function
    End If

    Set objOutStream = objFSO.OpenTextFile(strFilePath, 2, True, intEncoding)

    For Each strSection In Sections
      Call objOutStream.WriteLine("[" & strSection & "]")

      For Each strKey In Keys(strSection)
        Call objOutStream.WriteLine(strKey & "=" & Value(strSection, strKey))
      Next
    Next

    objOutStream.Close

    SaveFile = True
  End Function


  Public Function SectionExists(strSection)
    SectionExists = dicIniFile.Exists(strSection)
  End Function


  Public Function AddSection(strSection)
    Dim dicSection

    If Not dicIniFile.Exists(strSection) Then
      Set dicSection         = CreateObject("Scripting.Dictionary")
      dicSection.CompareMode = vbTextCompare

      Call dicIniFile.Add(strSection, dicSection)
      AddSection = True
    Else
      AddSection = False
    End If
  End Function


  Public Function DeleteSection(strSection)
    If dicIniFile.Exists(strSection) Then
      Call dicIniFile.Remove(strSection)
      DeleteSection = True
    Else
      DeleteSection = False
    End If
  End Function


  Public Function ClearSection(strSection)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)
      dicSection.RemoveAll
      ClearSection = True
    Else
      ClearSection = False
    End If
  End Function


  Public Function KeyExists(strSection, strKey)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)
      KeyExists = dicSection.Exists(strKey)
    Else
      KeyExists = False
    End If
  End Function


  Public Function AddKeyValue(strSection, strKey, strValue)
    Value(strSection, strKey) = strValue
    AddKeyValue = (Value(strSection, strKey) = strValue)
  End Function


  Public Function DeleteKey(strSection, strKey)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)

      If dicSection.Exists(strKey) Then
        Call dicSection.Remove(strKey)
        DeleteKey = True
      Else
        DeleteKey = False
      End If
    Else
      DeleteKey = False
    End If
  End Function


  Public Function ClearKey(strSection, strKey)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)

      If dicSection.Exists(strKey) Then
        dicSection.Item(strKey) = ""
        ClearKey = True
      Else
        ClearKey = False
      End If
    Else
      ClearKey = False
    End If
  End Function


  Public Property Get Sections
    Sections = dicIniFile.Keys
  End Property


  Public Property Get Keys(ByRef strSection)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)
      Keys = dicSection.Keys
    Else
      Keys = Array()
    End If
  End Property


  Public Property Get Value(ByRef strSection, ByRef strKey)
    Dim dicSection

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)

      If dicSection.Exists(strKey) Then
        Value = dicSection.Item(strKey)
      Else
        Value = ""
      End If
    Else
      Value = ""
    End If
  End Property


  Public Property Let Value(ByRef strSection, ByRef strKey, ByRef strValue)
    Dim dicSection

    Call AddSection(strSection)

    Set dicSection = dicIniFile.Item(strSection)

    If Not dicSection.Exists(strKey) Then
      Call dicSection.Add(strKey, strValue)
    Else
      dicSection.Item(strKey) = strValue
    End If
  End Property


  Public Property Get KeyValues(ByRef strSection)
    Dim dicSection, strKey, intIdx, arrResult

    If dicIniFile.Exists(strSection) Then
      Set dicSection = dicIniFile.Item(strSection)

      If dicSection.Count > 0 Then
        ReDim arrResult(dicSection.Count - 1, 1)
        intIdx = 0

        For Each strKey In dicSection.Keys
          arrResult(intIdx, 0) = strKey
          arrResult(intIdx, 1) = dicSection.Item(strKey)
          intIdx = intIdx + 1
        Next
      Else
        arrResult = Array()
      End If
    Else
      arrResult = Array()
    End If

    KeyValues = arrResult
  End Property


  Public Property Let KeyValues(ByRef strSection, ByRef arrKeyValues)
    Dim dicSection
    Dim intArrLength, intCnt

    If Not IsArrayDimed(arrKeyValues, 1) Then Exit Property
    If UBound(arrKeyValues, 1) < 0 Then Exit Property

    If Not IsArrayDimed(arrKeyValues, 2) Then
      If Not IsArrayDimed(arrKeyValues(0), 1) Then Exit Property
      If UBound(arrKeyValues(0), 1) < 1 Then Exit Property
    Else
      If UBound(arrKeyValues, 2) < 1 Then Exit Property
    End If

    Call AddSection(strSection)

    Set dicSection = dicIniFile.Item(strSection)

    If Not IsArrayDimed(arrKeyValues, 2) Then
      'Manually created 2-dimensional array
      For intCnt = 0 To UBound(arrKeyValues)
        Value(strSection, arrKeyValues(intCnt)(0)) = arrKeyValues(intCnt)(1)
      Next
    Else
      'Native 2-dimensional VBScript array
      For intCnt = 0 To UBound(arrKeyValues)
        Value(strSection, arrKeyValues(intCnt, 0)) = arrKeyValues(intCnt, 1)
      Next
    End If
  End Property
End Class
