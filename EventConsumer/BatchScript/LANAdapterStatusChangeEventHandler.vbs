strDevconPath = """C:\Windows\System32\Devcon.exe"""


' Für XP SP2, SP1 und ohne SP
' LAN_CONNECTED    = "9"

' Ab XP SP3 und neuer
LAN_CONNECTED    = "2"

LAN_DISCONNECTED = "7"
MINIMIZED        = "7"
WAIT             = vbTrue

Set listArgs = WScript.Arguments

If listArgs.Count > 1 Then
  strWLANOn  = strDevconPath & " enable ""@" & listArgs(1) & """"
  strWLANOff = strDevconPath & " disable ""@" & listArgs(1) & """"

  Set WshShell = WScript.CreateObject("WScript.Shell")

  Select Case listArgs(0)
    Case LAN_CONNECTED
      WshShell.Run strWLANOff, MINIMIZED, WAIT

    Case LAN_DISCONNECTED
      WshShell.Run strWLANOn, MINIMIZED, WAIT

  End Select

  Set WshShell = Nothing
End If

Set listArgs = Nothing
