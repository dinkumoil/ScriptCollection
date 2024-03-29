'///////////////////////////////////////////////////////////////////////////////
'
' Header file for running processes and querying script's elevation state
'
' Author: Andreas Heim
'
' Required header files (to be included before):
'   - Utils.vbs
'
'///////////////////////////////////////////////////////////////////////////////



'===============================================================================
' Window show state constants for WshShell.Run and Shell.Application.ShellExecute
'===============================================================================

Const SW_HIDE            = 0
Const SW_SHOWNORMAL      = 1
Const SW_SHOWMINIMIZED   = 2
Const SW_SHOWMAXIMIZED   = 3
Const SW_SHOWNOACTIVATE  = 4
Const SW_SHOW            = 5
Const SW_MINIMIZE        = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA          = 8
Const SW_RESTORE         = 9
Const SW_SHOWDEFAULT     = 10



'===============================================================================
' System's ANSI and console code page
'===============================================================================

Const OEMCodePage  = "cp850"
Const ANSICodePage = "windows-1252"



'===============================================================================
' Retrieve bitness of OS
'===============================================================================

Function GetOSBitness
  GetOSBitness = CInt(GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth)
End Function



'===============================================================================
' Retrieve bitness of Script Host process
'===============================================================================

Function GetScriptHostBitness
  If GetOSBitness = 32 Then
    GetScriptHostBitness = 32
  ElseIf InStr(1, WScript.Path, "SysWOW64", vbTextCompare) > 0 Then
    GetScriptHostBitness = 32
  Else
    GetScriptHostBitness = 64
  End If
End Function



'===============================================================================
' Tells if script runs under console or GUI context
'===============================================================================

Function IsConsoleApp
  IsConsoleApp = (InStr(1, WScript.FullName, "cscript.exe", vbTextCompare) > 0)
End Function



'===============================================================================
' Ensures that the script runs in console mode
'===============================================================================

Sub EnsureConsoleMode
  Dim objFSO
  Dim strScriptHostPath, strCommandLine

  If Not IsConsoleApp Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strScriptHostPath = objFSO.BuildPath(WScript.Path, "cscript.exe")
    strCommandLine    = strScriptHostPath & " /nologo """ & WScript.ScriptFullName & """ & pause"

    Call ExecCommand("cmd.exe", Array("/c ", strCommandLine), "", SW_SHOWNORMAL, False)
    WScript.Quit
  End If
End Sub



'===============================================================================
' Ensures that the script runs in console mode as a 32 bit process
'===============================================================================

Sub EnsureConsoleMode32Bit
  Dim objFSO, objShell
  Dim strScriptHostPath, strCommandLine

  If Not IsConsoleApp Then
    Set objFSO   = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If GetScriptHostBitness = 32 Then
      strScriptHostPath = objFSO.BuildPath(WScript.Path, "cscript.exe")
    Else
      strScriptHostPath = objFSO.BuildPath(objShell.ExpandEnvironmentStrings("%SystemRoot%\SysWOW64"), "cscript.exe")
    End If

    strCommandLine = strScriptHostPath & " /nologo """ & WScript.ScriptFullName & """ & pause"

    Call ExecCommand("cmd.exe", Array("/c ", strCommandLine), "", SW_SHOWNORMAL, False)
    WScript.Quit
  End If
End Sub



'===============================================================================
' Check if script runs with elevated user rights
'===============================================================================

Function IsElevated()
  Dim objWshShell, strKey

  Set objWshShell = CreateObject("WScript.Shell")

  On Error Resume Next
  strKey = objWshShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")

  IsElevated = (Err.Number = 0)
End Function



'===============================================================================
' Run a command asynchronously with elevated user rights
'===============================================================================

Sub RunElevated(ByRef strCommand, ByRef strWorkFolder, ByVal intShowState, ByRef arrParams)
  Dim objShell
  Dim intCnt, strParam, strArguments

  Set objShell = CreateObject("Shell.Application")
  strArguments = ""

  For intCnt = 0 To UBound(arrParams)
    strParam = Trim(arrParams(intCnt))

    If InStr(strParam, " ") > 0 Then
      strArguments = strArguments & " " & Quote(strParam)
    Else
      strArguments = strArguments & " " & strParam
    End If
  Next

  objShell.ShellExecute strCommand, strArguments, strWorkFolder, "runas", intShowState
End Sub



'===============================================================================
' Restart script with elevated user rights
'===============================================================================

Sub RestartElevated(ByVal intShowState, ByRef arrParams)
  Dim objFSO, arrArguments(), intCnt

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  ReDim arrArguments(UBound(arrParams) + 2)

  arrArguments(0) = "/nologo"
  arrArguments(1) = WScript.ScriptFullName

  For intCnt = 2 To UBound(arrArguments)
    arrArguments(intCnt) = arrParams(intCnt - 2)
  Next

  Call RunElevated(WScript.FullName, objFSO.GetParentFolderName(WScript.ScriptFullName), intShowState, arrArguments)

  WScript.Quit 0
End Sub



'===============================================================================
' Terminate all running instances of a program and restart it afterwards
'===============================================================================

Function TerminateAndRestart(ByRef strExePath, ByRef arrParams, ByVal bolForce, ByVal intMaxWaitSeconds)
  Dim objFSO, objWshShell, objWMIService
  Dim colProcesses, objProcess, strInstanceQuery, intInstanceCount, intDelCount
  Dim colEvents, objEvent, strEventQuery, intInterval
  Dim strExeDir, strExeName
  Dim datWaitStart

  TerminateAndRestart = False

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" )
  intInterval       = 1

  If objFSO.FileExists(strExePath) Then
    strExeDir        = objFSO.GetParentFolderName(strExePath)
    strExeName       = objFSO.GetFileName(strExePath)

    strInstanceQuery = "SELECT * FROM Win32_Process" & _
                       " WHERE ExecutablePath = '" & EscapeForWMI(strExePath) & "'"

    strEventQuery    = "SELECT * FROM __InstanceOperationEvent" & _
                       " WITHIN " & intInterval & _
                       " WHERE TargetInstance ISA 'Win32_Process'" & _
                       " AND TargetInstance.ExecutablePath = '" & EscapeForWMI(strExePath) & "'"

    If bolForce Then
      Call ExecCommand("taskkill", Array("/f", "/im", strExeName), "", SW_HIDE, True)

      Do
        WScript.Sleep 100
        Set colProcesses = objWMIService.ExecQuery(strInstanceQuery)
      Loop Until colProcesses.Count = 0

      Call ExecCommand(strExePath, Array(), strExeDir, SW_SHOW, False)
    Else
      Set colProcesses = objWMIService.ExecQuery(strInstanceQuery)
      intInstanceCount = colProcesses.Count
      intDelCount      = 0

      For Each objProcess In colProcesses
        Call ExecCommand("taskkill", Array("/im", strExeName), "", SW_HIDE, False)

        Set colEvents = objWMIService.ExecNotificationQuery(strEventQuery)
        datWaitStart  = Now

        Do
          Set objEvent = colEvents.NextEvent()

          Select Case objEvent.Path_.Class
            Case "__InstanceDeletionEvent"
              intDelCount  = intDelCount + 1
              datWaitStart = Now

            Case Else
              If intMaxWaitSeconds >= 0 Then
                If DateDiff("s", datWaitStart, Now) >= intMaxWaitSeconds Then Exit Function
              End If
          End Select
        Loop Until intDelCount = intInstanceCount

        Call ExecCommand(strExePath, Array(), strExeDir, SW_SHOW, False)

        Exit For
      Next
    End If

    TerminateAndRestart = True
  End If
End Function



'===============================================================================
' Execute a command synchronously or asynchronously
'===============================================================================

Function ExecCommand(ByRef strCommand, ByRef arrParams, ByRef strWorkDir, ByVal intShowState, ByVal bolWait)
  Dim objWshShell
  Dim strParam, intCnt, strArguments, strCommandLine

  Set objWshShell = CreateObject("WScript.Shell")
  strArguments    = ""

  For intCnt = 0 To UBound(arrParams)
    strParam = Trim(arrParams(intCnt))

    If InStr(strParam, " ") > 0 Then
      strArguments = strArguments & " " & Quote(strParam)
    Else
      strArguments = strArguments & " " & strParam
    End If
  Next

  strCommandLine = Quote(strCommand) & " " & strArguments

  If strWorkDir <> "" Then objWshShell.CurrentDirectory = strWorkDir
  ExecCommand = objWshShell.Run(strCommandLine, intShowState, bolWait)
End Function



'===============================================================================
' Execute a command synchronously and capture its output
'===============================================================================

Function ExecAndCapture(ByRef strCommand, ByRef arrParams, ByRef strStdOutOutput, ByRef strStdErrOutput)
  Dim objWshShell, objExec
  Dim strParam, strArguments, intCnt, strLine

  strStdOutOutput = ""
  strStdErrOutput = ""
  strArguments    = ""

  For intCnt = 0 To UBound(arrParams)
    strParam = Trim(arrParams(intCnt))

    If InStr(strParam, " ") > 0 Then
      strArguments = strArguments & " " & Quote(strParam)
    Else
      strArguments = strArguments & " " & strParam
    End If
  Next

  Set objWshShell = CreateObject("WScript.Shell")
  Set objExec     = objWshShell.Exec(Quote(strCommand) & " " & strArguments)

  Do While Not objExec.StdOut.AtEndOfStream
    strLine         = objExec.StdOut.ReadLine
    strStdOutOutput = strStdOutOutput & strLine & vbCrLf
  Loop

  Do While Not objExec.StdErr.AtEndOfStream
    strLine         = objExec.StdErr.ReadLine
    strStdErrOutput = strStdErrOutput & strLine & vbCrLf
  Loop

  strStdOutOutput = ConvertEncoding(strStdOutOutput, OEMCodePage, ANSICodePage)
  strStdErrOutput = ConvertEncoding(strStdErrOutput, OEMCodePage, ANSICodePage)

  Do While objExec.Status = 0
    WScript.Sleep 100
  Loop

  If objExec.ExitCode <> 0 Then
    ExecAndCapture = objExec.ExitCode
  ElseIf strStdErrOutput = "" Then
    ExecAndCapture = objExec.ExitCode
  Else
    ExecAndCapture = -1
  End If
End Function



'===============================================================================
' Retrieve PID of the process that has run this script
'===============================================================================

Function GetParentProcessPID(ByRef strExecutablePath)
  Dim objWshShell, objWshExec
  Dim objIndicatorProcess, objCurProcess
  Dim strWMIQuery, strCurProcessPath, intCurProcessId

  Set objWshShell = CreateObject("WScript.Shell")
  strWMIQuery     = "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:Win32_Process.Handle="

  Set objWshExec          = objWshShell.Exec("rundll32 kernel32,Sleep")
  Set objIndicatorProcess = GetObject(strWMIQuery & objWshExec.ProcessId)
  Set objCurProcess       = objIndicatorProcess
  strCurProcessPath       = ""

  Do
    On Error Resume Next
    Set objCurProcess = GetObject(strWMIQuery & objCurProcess.ParentProcessId)
    On Error GoTo 0

    If Err.Number <> 0 Then Exit Do

    intCurProcessId    = objCurProcess.ProcessId
    strCurProcessPath  = objCurProcess.ExecutablePath
  Loop Until StrComp(strCurProcessPath, strExecutablePath, vbTextCompare) = 0

  objIndicatorProcess.Terminate

  If StrComp(strCurProcessPath, strExecutablePath, vbTextCompare) = 0 Then
    GetParentProcessPID = intCurProcessId
  Else
    GetParentProcessPID = -1
  End	If
End Function
