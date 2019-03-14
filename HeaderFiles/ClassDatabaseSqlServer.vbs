'===============================================================================
' Database class
'===============================================================================

Class clsDataBase
  '-----------------------------------------------------------------------------
  ' Private fields
  '-----------------------------------------------------------------------------

  Private strScriptPath
  Private strConnString
  Private strSQLServerHostName
  Private strSQLServerInstanceName
  Private strUserName
  Private strUserPW
  Private strDatabase

  Private objFSO
  Private objDBConnection
  Private objCommand


  '-----------------------------------------------------------------------------
  ' Private methods
  '-----------------------------------------------------------------------------

  Private Sub Class_Initialize()
    On Error Resume Next

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
    strDatabase   = "master"

    strConnString = "Provider=SQLOLEDB; " & _
                    "Data Source=@1; " & _
                    "User ID=@2; " & _
                    "Password=@3; " & _
                    "Initial Catalog=@4; "

    Set objCommand      = Nothing
    Set objDBConnection = Nothing
  End Sub


  Private Sub Class_Terminate()
    On Error Resume Next

    Close()
    Set objFSO = Nothing
  End Sub


  Private Function BuildConnString
    Dim strResult

    strResult = ""

    If IsValid Then
      strResult = Replace(strConnString, "@1", SQLServer)
      strResult = Replace(strResult,     "@2", UserName)
      strResult = Replace(strResult,     "@3", UserPW)
      strResult = Replace(strResult,     "@4", Database)
    End If

    BuildConnString = strResult
  End Function


  Private Function NewCommand()
    Set objCommand = Nothing

    If Not objDBConnection Is Nothing Then
      If (objDBConnection.State And adStateOpen) = adStateOpen Then
        Set objCommand              = CreateObject("ADODB.Command")
        objCommand.ActiveConnection = objDBConnection
      End If
    End If

    Set NewCommand = objCommand
  End Function


  Private Function ReadScriptFile(ByRef strSqlScript, intFormat, ByRef strSql)
    On Error Resume Next

    Dim objInFile

    ReadScriptFile = False

    strSqlScript = objFSO.GetAbsolutePathName(strSqlScript)
    If Not objFSO.FileExists(strSqlScript) Then Exit Function

    If intFormat = AsAnsi Or intFormat = AsUnicode Or intFormat = AsSystemDefault Then
      Set objInFile = objFSO.OpenTextFile(strSqlScript, ForReading, False, intFormat)
      strSql        = objInFile.ReadAll
      objInFile.Close
    Else
      strSql = ReadUTF8File(strSqlScript)
    End If

    If Err.Number <> 0 Then Exit Function

    ReadScriptFile = True
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


  '-----------------------------------------------------------------------------
  ' Public methods
  '-----------------------------------------------------------------------------

  Public Function Open()
    On Error Resume Next

    Open = False

    If Not IsValid Then Exit Function

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


  Public Function SelectSingleValueByScriptFile(ByVal strSqlScript, ByRef arrParams, intFormat)
    On Error Resume Next

    Dim strSql

    SelectSingleValueByScriptFile = Null

    If ReadScriptFile(strSqlScript, intFormat, strSql) Then
      SelectSingleValueByScriptFile = SelectSingleValue(strSql, arrParams)
    End If
  End Function


  Public Function ExecuteNonQueryByScriptFile(ByVal strSqlScript, ByRef arrParams, intFormat)
    On Error Resume Next

    Dim strSql

    ExecuteNonQueryByScriptFile = False

    If ReadScriptFile(strSqlScript, intFormat, strSql) Then
      ExecuteNonQueryByScriptFile = ExecuteNonQuery(strSql, arrParams)
    End If
  End Function


  Public Function ExecuteQueryByScriptFile(ByVal strSqlScript, ByRef arrParams, intFormat)
    On Error Resume Next

    Dim strSql

    Set ExecuteQueryByScriptFile = Nothing

    If ReadScriptFile(strSqlScript, intFormat, strSql) Then
      Set ExecuteQueryByScriptFile = ExecuteQuery(strSql, arrParams)
    End If
  End Function


  Public Function SelectSingleValue(ByRef strSql, ByRef arrParams)
    Dim objResultSet

    SelectSingleValue = Null
    Set objResultSet  = ExecuteQuery(strSql, arrParams)

    If Not objResultSet Is Nothing Then
      If Not objResultSet.BOF And Not objResultSet.EOF Then
        SelectSingleValue = objResultSet.Fields(0).Value
      End If

      If (objResultSet.State And adStateOpen) = adStateOpen Then
        objResultSet.Close
      End If
    End If
  End Function


  Public Function ExecuteNonQuery(ByRef strSql, ByRef arrParams)
    Dim objResultSet

    ExecuteNonQuery  = False
    Set objResultSet = ExecuteQuery(strSql, arrParams)

    If Not objResultSet Is Nothing Then
      ExecuteNonQuery = True

      If (objResultSet.State And adStateOpen) = adStateOpen Then
        objResultSet.Close
      End If
    End If
  End Function


  Public Function ExecuteQuery(ByRef strSql, ByRef arrParams)
    On Error Resume Next

    Dim intCnt, objResultSet, intAffected

    Set ExecuteQuery = Nothing

    Call NewCommand()

    With Command
      .CommandType = adCmdText
      .CommandText = strSql
    End With

    For intCnt = LBound(arrParams) To UBound(arrParams)
      Call Command.Parameters.Append(ToSqlInputParameter(arrParams(intCnt)))
    Next

    Set objResultSet = Command.Execute(intAffected)

    If Err.Number = 0 Then
      Set ExecuteQuery = objResultSet
    End If
  End Function


  '-----------------------------------------------------------------------------
  ' Private properties
  '-----------------------------------------------------------------------------

   Private Property Get IsValid
    On Error Resume Next

    IsValid = SQLServer <> "" And _
              UserName  <> "" And _
              UserPW    <> "" And _
              Database  <> ""
   End Property


  '-----------------------------------------------------------------------------
  ' Public properties
  '-----------------------------------------------------------------------------

  Public Property Get Command
    On Error Resume Next
    Set Command = objCommand
  End Property


  Public Property Get SQLServer
    On Error Resume Next

    If SQLServerInstanceName = "" Then
      SQLServer = strSQLServerHostName
    Else
      SQLServer = strSQLServerHostName & "\" & SQLServerInstanceName
    End IF
  End Property


  Public Property Let SQLServer(strValue)
    On Error Resume Next

    Dim arrNameParts

    arrNameParts = Split(strValue, "\")

    If UBound(arrNameParts) >= 0 Then
      SQLServerHostName = arrNameParts(0)

      If UBound(arrNameParts) >= 1 Then
        SQLServerInstanceName = arrNameParts(1)
      End If
    End If
  End Property


  Public Property Get SQLServerHostName
    On Error Resume Next
    SQLServerHostName = strSQLServerHostName
  End Property


  Public Property Let SQLServerHostName(strValue)
    On Error Resume Next
    strSQLServerHostName = strValue
  End Property


  Public Property Get SQLServerInstanceName
    On Error Resume Next
    SQLServerInstanceName = strSQLServerInstanceName
  End Property


  Public Property Let SQLServerInstanceName(strValue)
    On Error Resume Next
    strSQLServerInstanceName = strValue
  End Property


  Public Property Get UserName
    On Error Resume Next
    UserName = strUserName
  End Property


  Public Property Let UserName(strValue)
    On Error Resume Next
    strUserName = strValue
  End Property


  Public Property Get UserPW
    On Error Resume Next
    UserPW = strUserPW
  End Property


  Public Property Let UserPW(strValue)
    On Error Resume Next
    strUserPW = strValue
  End Property


  Public Property Get Database
    On Error Resume Next
    Database = strDatabase
  End Property


  Public Property Let Database(strValue)
    On Error Resume Next
    strDatabase = strValue
  End Property

End Class
