'===============================================================================
' Datenbank-Klasse
'===============================================================================

Class clsDataBase
  '-----------------------------------------------------------------------------
  ' Private Variablen
  '-----------------------------------------------------------------------------

  Private strStdDataPath

  Private strScriptPath
  Private strDataPath
  Private strConnString

  Private objFSO
  Private objDBConnection
  Private objCommand


  '-----------------------------------------------------------------------------
  ' Private Methoden
  '-----------------------------------------------------------------------------

  Private Sub Class_Initialize()
    On Error Resume Next

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strStdDataPath = "Data"
    strDataPath    = strStdDataPath
    strScriptPath  = objFSO.GetParentFolderName(WScript.ScriptFullName)

    strConnString  = "Provider=MSDASQL.1; " &_
                     "FileDSN=@1; "

    Set objCommand      = Nothing
    Set objDBConnection = Nothing
  End Sub


  Private Sub Class_Terminate()
    On Error Resume Next

    Close()
    Set objFSO = Nothing
  End Sub


  Private Function BuildConnString
    BuildConnString = Replace(strConnString, "@1", DSNPath)
  End Function


  Private Function ToBooleanParameter(ByRef varValue, intDirection)
    Dim objResult, bolValue

    Set objResult = CreateObject("ADODB.Parameter")
    bolValue      = CBool(varValue)

    With objResult
      .Type      = adBoolean
      .Direction = intDirection
      .Value     = bolValue
    End With

    Set ToBooleanParameter = objResult
  End Function


  Private Function ToByteParameter(ByRef varValue, intDirection)
    Dim objResult, bytValue

    Set objResult = CreateObject("ADODB.Parameter")
    bytValue      = CByte(varValue)

    With objResult
      .Type      = adUnsignedTinyInt
      .Direction = intDirection
      .Value     = bytValue
    End With

    Set ToByteParameter = objResult
  End Function


  Private Function ToIntegerParameter(ByRef varValue, intDirection)
    Dim objResult, intValue

    Set objResult = CreateObject("ADODB.Parameter")
    intValue      = CInt(varValue)

    With objResult
      .Type      = adInteger
      .Direction = intDirection
      .Value     = intValue
    End With

    Set ToIntegerParameter = objResult
  End Function


  Private Function ToLongParameter(ByRef varValue, intDirection)
    Set ToLongParameter = ToIntegerParameter(varValue, intDirection)
  End Function


  Private Function ToSingleParameter(ByRef varValue, intDirection)
    Dim objResult, sngValue

    Set objResult = CreateObject("ADODB.Parameter")
    sngValue      = CSng(varValue)

    With objResult
      .Type      = adSingle
      .Direction = intDirection
      .Value     = sngValue
    End With

    Set ToSingleParameter = objResult
  End Function


  Private Function ToDoubleParameter(ByRef varValue, intDirection)
    Dim objResult, dblValue

    Set objResult = CreateObject("ADODB.Parameter")
    dblValue      = CDbl(varValue)

    With objResult
      .Type      = adDouble
      .Direction = intDirection
      .Value     = dblValue
    End With

    Set ToDoubleParameter = objResult
  End Function


  Private Function ToCurrencyParameter(ByRef varValue, intDirection)
    Dim objResult, curValue

    Set objResult = CreateObject("ADODB.Parameter")
    curValue      = CDbl(varValue)

    With objResult
      .Type      = adDouble
      .Direction = intDirection
      .Value     = curValue
    End With

    Set ToCurrencyParameter = objResult
  End Function


  Private Function ToDateParameter(ByRef varValue, intDirection)
    Dim objResult, datValue

    Set objResult = CreateObject("ADODB.Parameter")
    datValue      = CDate(varValue)

    With objResult
      .Type      = adDate
      .Direction = intDirection
      .Value     = datValue
    End With

    Set ToDateParameter = objResult
  End Function


  Private Function ToStringParameter(ByRef varValue, intDirection)
    Dim objResult, strValue

    Set objResult = CreateObject("ADODB.Parameter")
    strValue      = CStr(varValue)

    With objResult
      .Type      = adVarChar
      .Direction = intDirection
      .Value     = strValue
      .Size      = Len(strValue)
    End With

    Set ToStringParameter = objResult
  End Function


  Private Function ToSqlInputParameter(ByRef varValue)
    Set ToSqlInputParameter = Eval("To" & TypeName(varValue) & "Parameter(varValue, adParamInput)")
  End Function


  ' Für Debugging-Zwecke
  Private Sub HandleResults(ByRef objResultSet, ByVal intAffected)
    Dim objDbConnection, objCommand, objDbError, objParam
    Dim intItemCnt, intRecCnt
    Dim intIdx, strLine

    Set objDbConnection = objResultSet.ActiveConnection
    Set objCommand      = objResultSet.ActiveCommand
    intItemCnt          = 0

    Do While Not (objResultSet Is Nothing)
      intItemCnt = intItemCnt + 1

      If objResultSet.State <> adStateClosed Then
        intRecCnt = 0
        
        Do While Not objResultSet.EOF
          intRecCnt = intRecCnt + 1
          strLine   = ""
          
          For intIdx = 0 To objResultSet.Fields.Count - 1
            strLine = strLine & " | " & objResultSet(intIdx)
          Next
          
          WScript.Echo strLine
          objResultSet.MoveNext
        Loop
        
        WScript.Echo "Item " & intItemCnt & ": Recordset has " & intRecCnt & " records and " & objResultSet.Fields.Count & " fields."

      ElseIf objDbConnection.Errors.Count > 0 Then
        For Each objDbError In objDbConnection.Errors
          WScript.Echo "Item " & intItemCnt & ": Error " & objDbError.Number & " " & objDbError.Description
        Next

      Else
        WScript.Echo "Item " & intItemCnt & ": " & intAffected & " records intAffected."
      End If

      On Error Resume Next
      Set objResultSet = objResultSet.NextRecordset(intAffected)

      If Err.Number <> 0 Then
        intItemCnt = intItemCnt + 1
        WScript.Echo "Item " & intItemCnt & ": Fatal Error " & Err.Number & " " & Err.Description

        For Each objParam In objCommand.Parameters
          Select Case objParam.Direction
            Case adParamReturnValue
              WScript.Echo "Return value: " & objParam.Value
            Case adParamOutput
              WScript.Echo "Output: " & objParam.Value
            Case adParamInputOutput
              WScript.Echo "Changed: " & objParam.Value
          End Select
        Next

        Exit Sub
      End If

      On Error Goto 0
    Loop
  End Sub


  '-----------------------------------------------------------------------------
  ' Private Properties
  '-----------------------------------------------------------------------------

  Private Property Get DSNPath
    On Error Resume Next

    Dim colFiles, objFile, strPath

    strDSNPath = ""
    strPath    = DataPath

    Set colFiles = objFSO.GetFolder(strPath).Files

    For Each objFile In colFiles
      If StrComp(objFSO.GetExtensionName(objFile.Name), "dsn", vbTextCompare) = 0 Then
        strDSNPath = objFSO.BuildPath(strPath, objFile.Name)
        Exit For
      End If
    Next

    DSNPath = strDSNPath
  End Property


  '-----------------------------------------------------------------------------
  ' Public Methoden
  '-----------------------------------------------------------------------------

  Public Function Open
    On Error Resume Next

    Open = False

    Set objDBConnection = CreateObject("ADODB.Connection")

    If Not objDBConnection Is Nothing Then
      objDBConnection.Open(BuildConnString)
      If not NewCommand Is Nothing Then Open = True
    End If
  End Function


  Public Sub Close()
    On Error Resume Next

    If Not objDBConnection Is Nothing Then
      If (objDBConnection.State And adStateOpen) = adStateOpen Then
        objDBConnection.Close()
      End If

      Set objCommand      = Nothing
      Set objDBConnection = Nothing
    End If
  End Sub


  Public Function NewCommand()
    Set objCommand = Nothing

    If Not objDBConnection Is Nothing Then
      If (objDBConnection.State And adStateOpen) = adStateOpen Then
        Set objCommand              = CreateObject("ADODB.Command")
        objCommand.ActiveConnection = objDBConnection
      End If
    End If

    Set NewCommand = objCommand
  End Function


  Public Function ExecuteQuery(ByRef strSql, ByRef arrParams)
    Dim intCnt, objResultSet, objResult

    Set objResult = Nothing

    Call NewCommand()

    With Command
      .CommandType = adCmdText
      .CommandText = strSql
    End With

    For intCnt = LBound(arrParams) To UBound(arrParams)
      Call Command.Parameters.Append(ToSqlInputParameter(arrParams(intCnt)))
    Next

    'On Error Resume Next
      Set objResultSet = Command.Execute

      If Err.Number = 0 Then
        Set objResult = objResultSet
      End If
    'On Error GoTo 0

    Set ExecuteQuery = objResult
  End Function


  Public Function SelectSingleValue(ByRef strSql, ByRef arrParams)
    Dim intCnt, varValue, objResultSet, varResult

    varResult        = Null
    Set objResultSet = ExecuteQuery(strSql, arrParams)

    If Not objResultSet Is Nothing Then
      If Not objResultSet.BOF And Not objResultSet.EOF Then
        varResult = objResultSet.Fields(0).Value
      End If

      If (objResultSet.State And adStateOpen) = adStateOpen Then
        objResultSet.Close
      End If
    End If

    SelectSingleValue = varResult
  End Function


  Public Function ExecuteNonQuery(ByRef strSql, ByRef arrParams)
    Dim objResultSet

    Set objResultSet = ExecuteQuery(strSql, arrParams)
    ExecuteNonQuery  = (Err.Number = 0)

    If Not objResultSet Is Nothing Then
      If (objResultSet.State And adStateOpen) = adStateOpen Then
        objResultSet.Close
      End If
    End If
  End Function


  Public Function AddTableToSchema(ByRef strTableName, ByRef arrColumnData)
    Dim strSchemaFilePath, objSchemaFile, intCnt

    AddTableToSchema = False

    Set objSchemaFile = New clsIniFile
    strSchemaFilePath = objFSO.BuildPath(DataPath, "schema.ini")

    If objSchemaFile.LoadFile(strSchemaFilePath, AsAnsi) Then
      objSchemaFile.KeyValues(strTableName) = objSchemaFile.KeyValues("StandardSettings")

      For intCnt = LBound(arrColumnData) to UBound(arrColumnData)
        objSchemaFile.Value(strTableName, "Col" & intCnt+1) = arrColumnData(intCnt)
      Next

      Call objSchemaFile.SaveFile(strSchemaFilePath, AsAnsi, True)

      AddTableToSchema = True
    End If
  End Function


  Public Function RemoveTableFromSchema(ByRef strTableName)
    Dim strSchemaFilePath, objSchemaFile, intCnt

    RemoveTableFromSchema = False

    Set objSchemaFile = New clsIniFile
    strSchemaFilePath = objFSO.BuildPath(DataPath, "schema.ini")

    If objSchemaFile.LoadFile(strSchemaFilePath, AsAnsi) Then
      If objSchemaFile.DeleteSection(strTableName) Then
        Call objSchemaFile.SaveFile(strSchemaFilePath, AsAnsi, True)
        RemoveTableFromSchema = True
      End If
    End If
  End Function


  '-----------------------------------------------------------------------------
  ' Public Properties
  '-----------------------------------------------------------------------------

  Public Property Get Command
    On Error Resume Next
    Set Command = objCommand
  End Property


  Public Property Get DataPath
    On Error Resume Next

    If strDataPath = strStdDataPath Then
      strDataPath = objFSO.BuildPath(strScriptPath, strStdDataPath)
    End If

    DataPath = strDataPath
  End Property


  Public Property Let DataPath(strValue)
    On Error Resume Next
    strDataPath = objFSO.GetAbsolutePathName(strValue)
  End Property

End Class
