'*******************************************************************************
'
' Script to set firewall rules for MS SQL Server instances installed at the
' machine the script is executed on.
'
' The rule for SQL Server Browser is only set if a named instance of SQL Server
' has been found.
'
' The rule for NetBIOS name service is only set if the local machine is NOT a
' domain member.
'
' The two files ClassRegistry.vbs and ClassSQLServer.vbs are required to be
' stored in the same directory like this script.
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

'Variables have to be declared before first usage
Option Explicit


'Include external Code
Include ".\ClassRegistry.vbs"
Include ".\ClassSQLServer.vbs"


'Set name of firewall rules group
Const RULES_GROUP_NAME = "My Group"


'Variables declaration
Dim arrSQLServers, objSQLServer


'Restart script with elevated rights if required
Call CheckElevation()

'Scan registry for installed instances of MS SQL Server
arrSQLServers = ScanRegistryForSQLServerInstances()

'Set firewall rules for all instances of MS SQL Server found
If UBound(arrSQLServers) >= 0 Then
  For Each objSQLServer In arrSQLServers
    If MsgBox("Create firewall rules for MS SQL Server instance" & vbNewLine _
              & vbNewLine _
              & objSQLServer.InstanceName, _
              vbQuestion + vbYesNo + vbDefaultButton1, _
              "Set MS SQL Server firewall rules") = vbYes Then
      Call SetFirewallRules(objSQLServer, RULES_GROUP_NAME)
    End If
  Next
Else
  WScript.Echo "No installed instances of MS SQL Server found."
End If



'===============================================================================
' Set firewall rules for MS SQL Server
'===============================================================================

Sub SetFirewallRules(ByRef objSQLServer, ByRef strGroupName)
  On Error Resume Next

  'IP protocol constants => IANA protocol numbers
  Const NET_FW_IP_PROTOCOL_TCP    =  6
  Const NET_FW_IP_PROTOCOL_UDP    = 17
  Const NET_FW_IP_PROTOCOL_ICMPv4 =  1
  Const NET_FW_IP_PROTOCOL_ICMPv6 = 58

  'Profile types
  Const NET_FW_PROFILE2_DOMAIN    = &H0001
  Const NET_FW_PROFILE2_PRIVATE   = &H0002
  Const NET_FW_PROFILE2_PUBLIC    = &H0004
  Const NET_FW_PROFILE2_ALL       = &H7FFFFFFF

  'Actions
  Const NET_FW_ACTION_BLOCK       = 0
  Const NET_FW_ACTION_ALLOW       = 1

  'Remote address constants
  Const NET_FW_REMOTE_ADDR_ANY    = "*"

  'Rule names
  Const RULE_SQL_SERVER_APP_TCP   = "%1 (%2) via TCP"
  Const RULE_SQL_SERVER_APP_UDP   = "%1 (%2) via UDP"
  Const RULE_SQL_SERVER_BROWSER   = "MS SQL Server Browser"
  Const RULE_NETBIOS_NAMESERVICE  = "NetBios name service"
  Const RULE_PING_V4_IN           = "ICMPv4 Echo Request"
  Const RULE_PING_V6_IN           = "ICMPv6 Echo Request"

  Dim objFirewall, intCurrentProfiles, objCurRule, objNewRule

  Set objFirewall    = CreateObject("HNetCfg.FwPolicy2")
  intCurrentProfiles = objFirewall.CurrentProfileTypes

  'When possible we avoid adding firewall rules to the Public profile.
  'If Public is currently active and it is not the only active profile, we remove it from the bitmask
  If ((intCurrentProfiles And NET_FW_PROFILE2_PUBLIC) <> 0) And (intCurrentProfiles <> NET_FW_PROFILE2_PUBLIC) Then
    intCurrentProfiles = (intCurrentProfiles Xor NET_FW_PROFILE2_PUBLIC)
  End If

  'Register sqlservr.exe as authorized application for inbound connections via TCP
  Set objCurRule = objFirewall.Rules.Item(FormatString(RULE_SQL_SERVER_APP_TCP, Array(objSQLServer.FullName, objSQLServer.InstanceName)))

  If IsObject(objCurRule) Then
    If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
      objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
    End If
  Else
    Set objNewRule             = CreateObject("HNetCfg.FWRule")
    objNewRule.Name            = FormatString(RULE_SQL_SERVER_APP_TCP, Array(objSQLServer.FullName, objSQLServer.InstanceName))
    objNewRule.ApplicationName = objSQLServer.InstanceExePath
    objNewRule.Protocol        = NET_FW_IP_PROTOCOL_TCP
    objNewRule.RemoteAddresses = NET_FW_REMOTE_ADDR_ANY
    objNewRule.Action          = NET_FW_ACTION_ALLOW
    objNewRule.Grouping        = strGroupName
    objNewRule.Profiles        = intCurrentProfiles
    objNewRule.Enabled         = True
    objFirewall.Rules.Add(objNewRule)
  End If

  'Register sqlservr.exe as authorized application for inbound connections via UDP
  Set objCurRule = objFirewall.Rules.Item(FormatString(RULE_SQL_SERVER_APP_UDP, Array(objSQLServer.FullName, objSQLServer.InstanceName)))

  If IsObject(objCurRule) Then
    If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
      objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
    End If
  Else
    Set objNewRule             = CreateObject("HNetCfg.FWRule")
    objNewRule.Name            = FormatString(RULE_SQL_SERVER_APP_UDP, Array(objSQLServer.FullName, objSQLServer.InstanceName))
    objNewRule.ApplicationName = objSQLServer.InstanceExePath
    objNewRule.Protocol        = NET_FW_IP_PROTOCOL_UDP
    objNewRule.RemoteAddresses = NET_FW_REMOTE_ADDR_ANY
    objNewRule.Action          = NET_FW_ACTION_ALLOW
    objNewRule.Grouping        = strGroupName
    objNewRule.Profiles        = intCurrentProfiles
    objNewRule.Enabled         = True
    objFirewall.Rules.Add(objNewRule)
  End If

  If Not objSQLServer.IsDefaultInstance Then
    'Open port 1434 UDP (MS SQL Server Browser) for inbound connections
    Set objCurRule = objFirewall.Rules.Item(RULE_SQL_SERVER_BROWSER)

    If IsObject(objCurRule) Then
      If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
        objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
      End If
    Else
      Set objNewRule             = CreateObject("HNetCfg.FWRule")
      objNewRule.Name            = RULE_SQL_SERVER_BROWSER
      objNewRule.Protocol        = NET_FW_IP_PROTOCOL_UDP
      objNewRule.LocalPorts      = "1434"
      objNewRule.RemoteAddresses = NET_FW_REMOTE_ADDR_ANY
      objNewRule.Action          = NET_FW_ACTION_ALLOW
      objNewRule.Grouping        = strGroupName
      objNewRule.Profiles        = intCurrentProfiles
      objNewRule.Enabled         = True
      objFirewall.Rules.Add(objNewRule)
    End If
  End If

  If Not objSQLServer.IsDomainMember Then
    'Open port 137 UDP (NetBios name service) for inbound connections
    Set objCurRule = objFirewall.Rules.Item(RULE_NETBIOS_NAMESERVICE)

    If IsObject(objCurRule) Then
      If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
        objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
      End If
    Else
      Set objNewRule             = CreateObject("HNetCfg.FWRule")
      objNewRule.Name            = RULE_NETBIOS_NAMESERVICE
      objNewRule.Protocol        = NET_FW_IP_PROTOCOL_UDP
      objNewRule.LocalPorts      = "137"
      objNewRule.RemoteAddresses = NET_FW_REMOTE_ADDR_ANY
      objNewRule.Action          = NET_FW_ACTION_ALLOW
      objNewRule.Grouping        = strGroupName
      objNewRule.Profiles        = intCurrentProfiles
      objNewRule.Enabled         = True
      objFirewall.Rules.Add(objNewRule)
    End If
  End If

  'Allow inbound PING packets (ICMPv4)
  Set objCurRule = objFirewall.Rules.Item(RULE_PING_V4_IN)

  If IsObject(objCurRule) Then
    If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
      objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
    End If
  Else
    Set objNewRule               = CreateObject("HNetCfg.FWRule")
    objNewRule.Name              = RULE_PING_V4_IN
    objNewRule.Protocol          = NET_FW_IP_PROTOCOL_ICMPv4
    objNewRule.IcmpTypesAndCodes = "8:0"
    objNewRule.RemoteAddresses   = NET_FW_REMOTE_ADDR_ANY
    objNewRule.Action            = NET_FW_ACTION_ALLOW
    objNewRule.Grouping          = strGroupName
    objNewRule.Profiles          = intCurrentProfiles
    objNewRule.Enabled           = True
    objFirewall.Rules.Add(objNewRule)
  End If

  'Allow inbound PING packets (ICMPv6)
  Set objCurRule = objFirewall.Rules.Item(RULE_PING_V6_IN)

  If IsObject(objCurRule) Then
    If (objCurRule.Profiles And intCurrentProfiles) <> intCurrentProfiles Then
      objCurRule.Profiles = (objCurRule.Profiles Or intCurrentProfiles)
    End If
  Else
    Set objNewRule               = CreateObject("HNetCfg.FWRule")
    objNewRule.Name              = RULE_PING_V6_IN
    objNewRule.Protocol          = NET_FW_IP_PROTOCOL_ICMPv6
    objNewRule.IcmpTypesAndCodes = "128:0"
    objNewRule.RemoteAddresses   = NET_FW_REMOTE_ADDR_ANY
    objNewRule.Action            = NET_FW_ACTION_ALLOW
    objNewRule.Grouping          = strGroupName
    objNewRule.Profiles          = intCurrentProfiles
    objNewRule.Enabled           = True
    objFirewall.Rules.Add(objNewRule)
  End If
End Sub


'===============================================================================
' Scan registry for installed instances of MS SQL Server
'===============================================================================

Function ScanRegistryForSQLServerInstances()
  'At first we try machine's native hive
  ScanRegistryForSQLServerInstances = DoScanRegistryForSQLServerInstances(False)

  'If that failed we assume a 64 bit machine and try its 32 bit hive
  If UBound(ScanRegistryForSQLServerInstances) < 0 Then
    ScanRegistryForSQLServerInstances = DoScanRegistryForSQLServerInstances(True)
  End If
End Function


Function DoScanRegistryForSQLServerInstances(bolScan32BitHive)
  Dim strSQLServerExeName
  Dim strRegKeySQLServerBase
  Dim strRegKeySQLServerInstanceNames, strRegKeySQLServerInstanceSetup
  Dim strRegValueSQLBinRoot, strRegValueVersion, strRegValueEdition, strRegValueSP

  Dim objFSO, objRegistry
  Dim arrSQLServerInstances, arrSQLServerInstanceIds, intSQLServerInstanceIdx
  Dim arrSQLServerInstanceSetupKeys, arrSQLServerInstanceSetupValues, intSQLServerInstanceSetupKeyIdx

  strSQLServerExeName = "sqlservr.exe"

  If bolScan32BitHive Then
    strRegKeySQLServerBase = "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server"
  Else
    strRegKeySQLServerBase = "SOFTWARE\Microsoft\Microsoft SQL Server"
  End If

  strRegKeySQLServerInstanceNames = strRegKeySQLServerBase & "\Instance Names\SQL"
  strRegKeySQLServerInstanceSetup = strRegKeySQLServerBase & "\%1\Setup"
  strRegValueSQLBinRoot           = "SQLBinRoot"
  strRegValueVersion              = "Version"
  strRegValueEdition              = "Edition"
  strRegValueSP                   = "SP"

  Set objFSO      = CreateObject("Scripting.FileSystemObject")
  Set objRegistry = New clsRegistry
  arrSQLServers   = Array()

  If objRegistry.EnumValues(objRegistry.HKEY_LOCAL_MACHINE, strRegKeySQLServerInstanceNames) Then
    arrSQLServerInstances   = objRegistry.ValueNames
    arrSQLServerInstanceIds = objRegistry.Values

    For intSQLServerInstanceIdx = 0 To UBound(arrSQLServerInstances)
      If arrSQLServerInstanceIds(intSQLServerInstanceIdx) <> "" Then
        strRegKeySQLServerInstanceSetup = FormatString(strRegKeySQLServerInstanceSetup, Array(arrSQLServerInstanceIds(intSQLServerInstanceIdx)))
        
        If objRegistry.EnumValues(objRegistry.HKEY_LOCAL_MACHINE, strRegKeySQLServerInstanceSetup) Then
          arrSQLServerInstanceSetupKeys   = objRegistry.ValueNames
          arrSQLServerInstanceSetupValues = objRegistry.Values

          ReDim Preserve arrSQLServers(UBound(arrSQLServers) + 1)
          Set arrSQLServers(UBound(arrSQLServers)) = New clsSQLServer

          With arrSQLServers(UBound(arrSQLServers))
            .InstanceName = arrSQLServerInstances(intSQLServerInstanceIdx)

            For intSQLServerInstanceSetupKeyIdx = 0 To UBound(arrSQLServerInstanceSetupKeys)
              If StrComp(arrSQLServerInstanceSetupKeys(intSQLServerInstanceSetupKeyIdx), strRegValueSQLBinRoot, vbTextCompare) = 0 Then
                .InstanceExePath = objFSO.BuildPath(arrSQLServerInstanceSetupValues(intSQLServerInstanceSetupKeyIdx), strSQLServerExeName)

              ElseIf StrComp(arrSQLServerInstanceSetupKeys(intSQLServerInstanceSetupKeyIdx), strRegValueVersion, vbTextCompare) = 0 Then
                .Version = arrSQLServerInstanceSetupValues(intSQLServerInstanceSetupKeyIdx)

              ElseIf StrComp(arrSQLServerInstanceSetupKeys(intSQLServerInstanceSetupKeyIdx), strRegValueEdition, vbTextCompare) = 0 Then
                .Edition = arrSQLServerInstanceSetupValues(intSQLServerInstanceSetupKeyIdx)

              ElseIf StrComp(arrSQLServerInstanceSetupKeys(intSQLServerInstanceSetupKeyIdx), strRegValueSP, vbTextCompare) = 0 Then
                .SPLevel = CInt(arrSQLServerInstanceSetupValues(intSQLServerInstanceSetupKeyIdx))
              End If
            Next
          End With

          If Not arrSQLServers(UBound(arrSQLServers)).IsValid Then
            ReDim Preserve arrSQLServers(UBound(arrSQLServers) - 1)
          End If
        End If
      End If
    Next
  End If

  DoScanRegistryForSQLServerInstances = arrSQLServers
End Function


'===============================================================================
' Check if script has to be restarted with elevated rights
'===============================================================================

Sub CheckElevation()
  If WScript.Arguments.Count = 0 Then Call RestartElevated()
  If StrComp(WScript.Arguments(0), "/elevated", vbTextCompare) <> 0 Then Call RestartElevated()
End Sub


'===============================================================================
' Restart script with elevated rights
'===============================================================================

Sub RestartElevated()
  Const SW_HIDE            =  0
  Const SW_SHOWNORMAL      =  1
  Const SW_SHOWMINIMIZED   =  2
  Const SW_SHOWMAXIMIZED   =  3
  Const SW_SHOWNOACTIVATE  =  4
  Const SW_SHOW            =  5
  Const SW_MINIMIZE        =  6
  Const SW_MAXIMIZE        =  3
  Const SW_SHOWMINNOACTIVE =  7
  Const SW_SHOWNA          =  8
  Const SW_RESTORE         =  9
  Const SW_SHOWDEFAULT     = 10
  
  Dim objShell, strApplication, strArguments

  Set objShell = CreateObject("Shell.Application")

  strApplication = WScript.FullName
  strArguments   = "/nologo """ & WScript.ScriptFullName & """ /elevated"

  objShell.ShellExecute strApplication, strArguments, "", "runas", SW_SHOWMINIMIZED

  WScript.Quit
End Sub


'===============================================================================
' Insert variable parts into a string 
'===============================================================================

Function FormatString(ByVal strString, ByRef arrItems)
  Dim intCnt, intStart, intPos
  Dim intDigits, strVar
  
  intStart  = 1
  intDigits = Len(CStr(UBound(arrItems) + 1))
  
  For intCnt = 0 To UBound(arrItems)
    strVar = "%" & Right(String(intDigits, "0") & intCnt+1, intDigits)
    intPos = InStr(intStart, strString, strVar, vbTextCompare)
    
    If intPos > 0 Then
      strString = Replace(strString, strVar, arrItems(intCnt), 1, 1, vbTextCompare)
      intStart  = intPos + Len(arrItems(intCnt))
    End If
  Next
  
  FormatString = strString
End Function


'===============================================================================
' Subroutine for including external code
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
