'*******************************************************************************
'
' Registry access class.
'
' Provides access to 64 bit registry hives even to scripts started with the
' 32 bit Windows Script Host.
'
' Author: Andreas Heim
' Date  : June 2017
'
'
' Disclaimer
' ==========
'
' This code is released "as is" without any guarantee to work proper for a
' certain subject. The author is not responsible for any damage caused by
' running this code.
'
'
' License
' =======
'
' This code can be changed and used for free by everyone. It is not allowed to
' claim any rights on this code or to charge any fees for it. If you change this
' file or deploy it to other persons the disclaimer and license text has to be
' kept included as well as a reference to the original author.
'
'*******************************************************************************

'WMI impersonation levels
Const wbemImpersonationLevelAnonymous     = 1
Const wbemImpersonationLevelIdentify      = 2
Const wbemImpersonationLevelImpersonate   = 3
Const wbemImpersonationLevelDelegate      = 4

'WMI authentication levels
Const wbemAuthenticationLevelDefault      = 0
Const wbemAuthenticationLevelNone         = 1
Const wbemAuthenticationLevelConnect      = 2
Const wbemAuthenticationLevelCall         = 3
Const wbemAuthenticationLevelPkt          = 4
Const wbemAuthenticationLevelPktIntegrity = 5
Const wbemAuthenticationLevelPktPrivacy   = 6

'WMI WbemFlags
Const wbemFlagReturnWhenComplete          = &H00
Const wbemFlagReturnImmediately           = &H10

Const wbemFlagBidirectional               = &H00
Const wbemFlagForwardOnly                 = &H20

Const wbemFlagReturnErrorObject           = &H00
Const wbemFlagNoErrorObject               = &H40

Const wbemFlagDontSendStatus              = &H00
Const wbemFlagSendStatus                  = &H80

Const wbemFlagUseAmendedQualifiers        = &H20000



Class clsRegistry
  Private objCtx, objSWbemLocator, objSWbemServices, objRegistry
  Private arrSubKeyNames, arrValueNames, arrValueTypes, arrValues
  Private arrBinaryValue, intDWORDValue, strStringValue, strExpandedStringValue, arrMultiStringValues
  Private intLastSubKeysQueryHive, strLastSubKeysQueryPath, intLastValueNamesQueryHive, strLastValueNamesQueryPath


  Private Sub Class_Initialize()
    On Error Resume Next

    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    objCtx.Add "__ProviderArchitecture", GetOSArchitecture()

    Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
    objSWbemLocator.Security_.AuthenticationLevel = wbemAuthenticationLevelPktPrivacy
    objSWbemLocator.Security_.ImpersonationLevel  = wbemImpersonationLevelImpersonate

    Set objSWbemServices = objSWbemLocator.ConnectServer(".", "root\default", "", "", , , , objCtx)

    If IsObject(objSWbemServices) Then
      objSWbemServices.Security_.AuthenticationLevel = wbemAuthenticationLevelPktPrivacy
      objSWbemServices.Security_.ImpersonationLevel  = wbemImpersonationLevelImpersonate

      Set objRegistry = objSWbemServices.Get("StdRegProv")

      If Not IsObject(objRegistry) Then
        WScript.Echo "Unable to instantiate StdRegProv class"
        Exit Sub
      End If
    Else
      WScript.Echo "Connecting to WMI failed"
      Exit Sub
    End If

    strLastKeyNameQuery   = ""
    strLastValueNameQuery = ""
    arrSubKeyNames        = Array()
    arrValueNames         = Array()
    arrValueTypes         = Array()
    arrValues             = Array()
    arrBinaryValue        = Array()
    arrMultiStringValues  = Array()
  End Sub


  Private Sub Class_Terminate()
    On Error Resume Next
    Set objRegistry = Nothing
    Erase arrSubKeyNames
    Erase arrValueNames
    Erase arrValueTypes
    Erase arrValues
    Erase arrBinaryValue
    Erase arrMultiStringValues
  End Sub


  Private Function GetOSArchitecture()
    Dim objOS, strProperty

    GetOSArchitecture = 32

    Set objOS = GetObject("winmgmts:\\.\root\CIMV2:Win32_OperatingSystem=@")

    On Error Resume Next
    
    strProperty = objOS.CreationClassName
    If InStr(1, strProperty, "64", vbTextCompare) > 0 Then GetOSArchitecture = 64

    strProperty = objOS.OSArchitecture
    If InStr(1, strProperty, "64", vbTextCompare) > 0 Then GetOSArchitecture = 64
  End Function


  Public Function CreateKey(intHive, strKeyPath)
    On Error Resume Next
    If objRegistry.CreateKey(intHive, strKeyPath) = 0 Then
      CreateKey = True
    Else
      CreateKey = False
    End If
  End Function


  Public Function DeleteKey(intHive, strKeyPath)
    On Error Resume Next
    If objRegistry.DeleteKey(intHive, strKeyPath) = 0 Then
      DeleteKey = True
    Else
      DeleteKey = False
    End If
  End Function


  Public Function DeleteValue(intHive, strKeyPath, strValueName)
    On Error Resume Next
    If objRegistry.DeleteValue(intHive, strKeyPath, strValueName) = 0 Then
      DeleteValue = True
    Else
      DeleteValue = False
    End If
  End Function


  Public Function GetBinaryValue(intHive, strKeyPath, strValueName)
    On Error Resume Next

    Erase arrBinaryValue
    arrBinaryValue = Array()

    If objRegistry.GetBinaryValue(intHive, strKeyPath, strValueName, arrBinaryValue) = 0 Then
      GetBinaryValue = arrBinaryValue
    Else
      GetBinaryValue = Array()
    End If
  End Function


  Public Function SetBinaryValue(intHive, strKeyPath, strValueName, arrNewBinaryValue)
    On Error Resume Next

    If objRegistry.SetBinaryValue(intHive, strKeyPath, strValueName, arrNewBinaryValue) = 0 Then
      SetBinaryValue = True
    Else
      SetBinaryValue = False
    End If
  End Function


  Public Function GetDWORDValue(intHive, strKeyPath, strValueName)
    On Error Resume Next

    If objRegistry.GetDWORDValue(intHive, strKeyPath, strValueName, intDWORDValue) = 0 Then
      GetDWORDValue = intDWORDValue
    Else
      GetDWORDValue = -2147483648
    End If
  End Function


  Public Function SetDWORDValue(intHive, strKeyPath, strValueName, intNewDWORDValue)
    On Error Resume Next

    If objRegistry.SetDWORDValue(intHive, strKeyPath, strValueName, intNewDWORDValue) = 0 Then
      SetDWORDValue = True
    Else
      SetDWORDValue = False
    End If
  End Function


  Public Function GetStringValue(intHive, strKeyPath, strValueName)
    On Error Resume Next

    strStringValue = ""

    If objRegistry.GetStringValue(intHive, strKeyPath, strValueName, strStringValue) = 0 Then
      GetStringValue = strStringValue
    Else
      GetStringValue = ""
    End If
  End Function


  Public Function SetStringValue(intHive, strKeyPath, strValueName, strNewStringValue)
    On Error Resume Next

    If objRegistry.SetStringValue(intHive, strKeyPath, strValueName, strNewStringValue) = 0 Then
      SetStringValue = True
    Else
      SetStringValue = False
    End If
  End Function


  Public Function GetExpandedStringValue(intHive, strKeyPath, strValueName)
    On Error Resume Next

    strExpandedStringValue = ""

    If objRegistry.GetExpandedStringValue(intHive, strKeyPath, strValueName, strExpandedStringValue) = 0 Then
      GetExpandedStringValue = strExpandedStringValue
    Else
      GetExpandedStringValue = ""
    End If
  End Function


  Public Function SetExpandedStringValue(intHive, strKeyPath, strValueName, strNewExpandedStringValue)
    On Error Resume Next

    If objRegistry.SetExpandedStringValue(intHive, strKeyPath, strValueName, strNewExpandedStringValue) = 0 Then
      SetExpandedStringValue = True
    Else
      SetExpandedStringValue = False
    End If
  End Function


  Public Function GetMultiStringValue(intHive, strKeyPath, strValueName)
    On Error Resume Next

    Erase arrMultiStringValues
    arrMultiStringValues = Array()

    If objRegistry.GetMultiStringValue(intHive, strKeyPath, strValueName, arrMultiStringValues) = 0 Then
      GetMultiStringValue = arrMultiStringValues
    Else
      GetMultiStringValue = Array()
    End If
  End Function


  Public Function SetMultiStringValue(intHive, strKeyPath, strValueName, arrNewMultiStringValues)
    On Error Resume Next

    If objRegistry.SetMultiStringValue(intHive, strKeyPath, strValueName, arrNewMultiStringValues) = 0 Then
      SetMultiStringValue = True
    Else
      SetMultiStringValue = False
    End If
  End Function


  Public Function EnumSubKeys(intHive, strKeyPath)
    On Error Resume Next

    intLastSubKeysQueryHive = intHive
    strLastSubKeysQueryPath = strKeyPath

    Erase arrSubKeyNames
    arrSubKeyNames = Array()

    If objRegistry.EnumKey(intHive, strKeyPath, arrSubKeyNames) = 0 Then
      EnumSubKeys = True
    Else
      EnumSubKeys = False
    End If
  End Function


  Public Function EnumValues(intHive, strKeyPath)
    On Error Resume Next

    intLastValueNamesQueryHive = intHive
    strLastValueNamesQueryPath = strKeyPath

    Erase arrValueNames
    Erase arrValueTypes
    arrValueNames = Array()
    arrValueTypes = Array()

    If objRegistry.EnumValues(intHive, strKeyPath, arrValueNames, arrValueTypes) = 0 Then
      EnumValues = True
    Else
      EnumValues = False
    End If
  End Function


  Public Property Get HiveName(intHive)
    Select Case intHive
      Case HKEY_CLASSES_ROOT
        HiveName = "HKEY_CLASSES_ROOT"

      Case HKEY_CURRENT_USER
        HiveName = "HKEY_CURRENT_USER"

      Case HKEY_LOCAL_MACHINE
        HiveName = "HKEY_LOCAL_MACHINE"

      Case HKEY_USERS
        HiveName = "HKEY_USERS"

      Case HKEY_CURRENT_CONFIG
        HiveName = "HKEY_CURRENT_CONFIG"

      Case Else
        HiveName = "UnknownHive"
    End Select
  End Property


  Public Property Get TypeName(intType)
    Select Case intType
      Case REG_SZ
        TypeName = "REG_SZ"

      Case REG_EXPAND_SZ
        TypeName = "REG_EXPAND_SZ"

      Case REG_BINARY
        TypeName = "REG_BINARY"

      Case REG_DWORD
        TypeName = "REG_DWORD"

      Case REG_MULTI_SZ
        TypeName = "REG_MULTI_SZ"

      Case Else
        TypeName = "UnknownType"
    End Select
  End Property


  Public Property Get SubKeyNames
    SubKeyNames = arrSubKeyNames
  End Property


  Public Property Get ValueNames
    ValueNames = arrValueNames
  End Property


  Public Property Get ValueTypes
    ValueTypes = arrValueTypes
  End Property


  Public Property Get Values
    On Error Resume Next

    Dim i, j, arrTempValues

    Erase arrValues
    arrValues = Array()

    If Not IsEmpty(arrValueNames) Then
      ReDim arrValues(UBound(arrValueNames))

      For i = 0 To UBound(arrValueNames)
        Erase arrTempValues
        arrTempValues = Array()

        Select Case arrValueTypes(i)
          Case REG_SZ
            arrValues(i) = GetStringValue(intLastValueNamesQueryHive, strLastValueNamesQueryPath, arrValueNames(i))

          Case REG_EXPAND_SZ
            arrValues(i) = GetExpandedStringValue(intLastValueNamesQueryHive, strLastValueNamesQueryPath, arrValueNames(i))

          Case REG_BINARY
            arrTempValues = GetBinaryValue(intLastValueNamesQueryHive, strLastValueNamesQueryPath, arrValueNames(i))

            For j = 0 To UBound(arrTempValues)
              If j = 0 Then
                arrValues(i) = "&H" & Hex(arrTempValues(j))
              Else
                arrValues(i) = arrValues(i) & " &H" & Right("0" & Hex(arrTempValues(j)), 2)
              End If
            Next

          Case REG_DWORD
            arrValues(i) = CStr(GetDWORDValue(intLastValueNamesQueryHive, strLastValueNamesQueryPath, arrValueNames(i)))

          Case REG_MULTI_SZ
            arrTempValues = GetMultiStringValue(intLastValueNamesQueryHive, strLastValueNamesQueryPath, arrValueNames(i))

            For j = 0 To UBound(arrTempValues)
              If j = 0 Then
                arrValues(i) = arrTempValues(j)
              Else
                arrValues(i) = arrValues(i) & ";" & arrTempValues(j)
              End If
            Next

          Case Else
            MsgBox "Unknown value type. Key path: " & HiveName(intLastValueNamesQueryHive) & "\" & strLastValueNamesQueryPath & "\" & arrValueNames(i), vbOKOnly
       End Select
      Next
    End If

    Values = arrValues
  End Property


  Public Property Get Value(idx)
    Value = arrValues(idx)
  End Property


  Public Property Get SubKeyName(idx)
    SubKeyName = arrSubKeyNames(idx)
  End Property


  Public Property Get ValueName(idx)
    ValueName = arrValueNames(idx)
  End Property


  Public Property Get ValueType(idx)
    ValueType = arrValueTypes(idx)
  End Property


  Public Property Get HKEY_CLASSES_ROOT
    HKEY_CLASSES_ROOT = &H80000000
  End Property


  Public Property Get HKEY_CURRENT_USER
    HKEY_CURRENT_USER = &H80000001
  End Property


  Public Property Get HKEY_LOCAL_MACHINE
    HKEY_LOCAL_MACHINE = &H80000002
  End Property


  Public Property Get HKEY_USERS
    HKEY_USERS = &H80000003
  End Property


  Public Property Get HKEY_CURRENT_CONFIG
    HKEY_CURRENT_CONFIG = &H80000005
  End Property


  Public Property Get REG_SZ
    REG_SZ = 1
  End Property


  Public Property Get REG_EXPAND_SZ
    REG_EXPAND_SZ = 2
  End Property


  Public Property Get REG_BINARY
    REG_BINARY = 3
  End Property


  Public Property Get REG_DWORD
    REG_DWORD = 4
  End Property


  Public Property Get REG_MULTI_SZ
    REG_MULTI_SZ = 7
  End Property

End Class
