'///////////////////////////////////////////////////////////////////////////////
'
' Header file for unifying script output in console and GUI context.
'
' Author: Andreas Heim
'
' Required header files (to be included before):
'   - Utils.vbs
'   - Process.vbs
'
'///////////////////////////////////////////////////////////////////////////////



Class clsUniversalInOut
  Private objFSO, objStdIn, objStdOut, objStdErr


  Private Sub Class_Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set objStdIn  = objFSO.GetStandardStream(0)
    Set objStdOut = objFSO.GetStandardStream(1)
    Set objStdErr = objFSO.GetStandardStream(2)
  End Sub


  Private Sub Class_Terminate()
    Set objStdErr = Nothing
    Set objStdOut = Nothing
    Set objStdIn  = Nothing
    Set objFSO    = Nothing
  End Sub


  Private Function ConvertToOEM850(strString)
    On Error Resume Next

    ConvertToOEM850 = ConvertEncoding(strString, Enc_Windows_1252, Enc_DOS_850)
  End Function


  Private Function ConvertToANSI(strString)
    On Error Resume Next

    ConvertToANSI = ConvertEncoding(strString, Enc_DOS_850, Enc_Windows_1252)
  End Function


  Private Sub ErrMessageBox(strMessage, strTitle)
    On Error Resume Next

    MsgBox strMessage, vbCritical, strTitle
  End Sub


  Private Sub MessageBox(strMessage, strTitle)
    On Error Resume Next

    MsgBox strMessage, vbOK, strTitle
  End Sub


  Public Function ReadString(strPrompt)
    On Error Resume Next

    Dim strInput

    objStdOut.Write ConvertToOEM850(strPrompt)

    If Err.Number = 0 Then
      strInput = ConvertToANSI(objStdIn.ReadLine)
    Else
      Err.Clear
      strInput = InputBox(strPrompt)
    End If

    ReadString = strInput
  End Function


  Public Sub OutputText(strString, boolGenCRLF)
    On Error Resume Next

    Dim strText

    If boolGenCRLF Then
      strText = Replace(strString, "|", vbCRLF)
    Else
      strText = strString
    End If

    objStdOut.WriteLine ConvertToOEM850(strText)

    If Err.Number <> 0 Then
      Err.Clear
      MessageBox strText, WScript.ScriptName
    End If
  End Sub


  Public Sub OutputErrorMsg(strMessage, boolGenCRLF)
    On Error Resume Next

    Dim strMsg

    If boolGenCRLF Then
      strMsg = Replace(strMessage, "|", vbCRLF)
    Else
      strMsg = strMessage
    End If

    objStdErr.WriteLine ConvertToOEM850(strMsg)

    If Err.Number <> 0 Then
      Err.Clear
      ErrMessageBox strMsg, WScript.ScriptName
    End If
  End Sub


  Public Property Get DefaultSystemCodePage
    DefaultSystemCodePage = CStr(GetObject("winmgmts:root\cimv2:Win32_OperatingSystem=@").CodeSet)
  End Property


  Public Property Get ConsoleCodePage
    Dim strStdOutOutput, strStdErrOutput

    If ExecAndCapture("cmd.exe", Array("/c", "chcp"), strStdOutOutput, strStdErrOutput) = 0 Then
      ConsoleCodePage = Trim(Split(Split(strStdOutOutput, ":")(1), ".")(0))
    Else
      ConsoleCodePage = ""
    End If
  End Property


  Public Property Get StdIn
    Set StdIn = objStdIn
  End Property


  Public Property Get StdOut
    Set StdOut = objStdOut
  End Property


  Public Property Get StdErr
    Set StdErr = objStdErr
  End Property
End Class
