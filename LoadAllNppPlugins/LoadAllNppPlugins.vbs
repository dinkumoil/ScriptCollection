Option Explicit


'===============================================================================
' Base configuration
'===============================================================================

'Path to the EXE file of Notepad++
Const NPP_DIR_PATH = "C:\Program Files (x86)\Notepad++"

'Path to the directory where GUP.exe should store plugin files after unpacking
'them from the ZIP package
Const GUP_UNZIP_DIR_PATH = ".\Plugins"

'Path to the directory of WinApiExec.exe
Const WAE_DIR_PATH = ".\bin"

'URL of WinApiExec download
Const WAE_URL = "https://rammichael.com/wp-content/uploads/downloads/2015/08/winapiexec.zip"

'===============================================================================


'Include external code
Include ".\include\IO.vbs"
Include ".\include\ClassJsonFile.vbs"


'Variables declaration
Dim objJsonFile, objFSO, objWshShell
Dim strNppDirPath, strNppPath, strGupPath, strWaeDirPath, strWaeDownloadPath, strWaePath
Dim strPluginListDllPath, strPluginListJsonPath, strGupUnzipDirPath
Dim intHTTPStatusCode, strPrevWorkingDir, arrPlugins, intCnt


'Variables initalization
Set objFSO            = CreateObject("Scripting.FileSystemObject")
Set objWshShell       = CreateObject("WScript.Shell")
Set objJsonFile       = New clsJsonFile

strNppDirPath         = objFSO.GetAbsolutePathName(NPP_DIR_PATH)
strWaeDirPath         = objFSO.GetAbsolutePathName(WAE_DIR_PATH)
strGupUnzipDirPath    = objFSO.GetAbsolutePathName(GUP_UNZIP_DIR_PATH)

strNppPath            = objFSO.BuildPath(strNppDirPath, "notepad++.exe")
strGupPath            = objFSO.BuildPath(strNppDirPath, "updater\GUP.exe")
strWaePath            = objFSO.BuildPath(strWaeDirPath, "winapiexec.exe")
strWaeDownloadPath    = objFSO.BuildPath(objWshShell.ExpandEnvironmentStrings("%TEMP%"), "WinApiExec.zip")
strPluginListDllPath  = objFSO.BuildPath(strNppDirPath, "plugins\config\nppPluginList.dll")
strPluginListJsonPath = objFSO.BuildPath(objWshShell.ExpandEnvironmentStrings("%TEMP%"), "nppPluginList.json")


'-------------------------------------------------------------------------------
' Retrieve command line arguments
'-------------------------------------------------------------------------------
Call ParseCommandLine()


'-------------------------------------------------------------------------------
' If WinApiExec.exe is not available download its ZIP file and unpack it
'-------------------------------------------------------------------------------
If Not objFSO.FileExists(strWaePath) Then
  If Not objFSO.FolderExists(strWaeDirPath) Then
    WScript.Echo "Directory for storing download of WinApiExec not found."
    WScript.Quit
  End If
  
  If Not DownloadFile(WAE_URL, strWaeDownloadPath, intHTTPStatusCode) Then
    WScript.Echo "WinApiExec.exe not found and downloading it failed. HTTP status code: " & intHTTPStatusCode
    WScript.Quit
  End If

  Call UnzipFile(strWaeDownloadPath, objFSO.GetFileName(strWaePath), strWaeDirPath)
End If


'-------------------------------------------------------------------------------
' Delete and recreate GUP's plugin packages unzip directory
'-------------------------------------------------------------------------------
If objFSO.FolderExists(strGupUnzipDirPath) Then
  Call objFSO.DeleteFolder(strGupUnzipDirPath)
  WScript.Sleep 1000
End If

Call objFSO.CreateFolder(strGupUnzipDirPath)


'-------------------------------------------------------------------------------
' Extract plugin list from DLL file and parse it
'-------------------------------------------------------------------------------
If strPluginListDllPath <> "" Then
  If Not ExtractPluginList(strPluginListDllPath, strPluginListJsonPath) Then
    WScript.Echo "Extraction of plugin list from DLL file failed."
    WScript.Quit
  End If

ElseIf Not objFSO.FileExists(strPluginListJsonPath) Then
  WScript.Echo "JSON file with plugin list not found."
  WScript.Quit
End If

'Ensure Windows EOL style of JSON file
Call ConvertUTF8EOLFormat(strPluginListJsonPath, vbCrLf)

If Not objJsonFile.LoadFromString(ReadUTF8File(strPluginListJsonPath)) Then
  WScript.Echo "Failed to parse JSON file with plugin list."
  WScript.Quit
End If

arrPlugins = objJsonFile.Content.Item("npp-plugins")


'-------------------------------------------------------------------------------
' Download and unzip all plugin packages
'-------------------------------------------------------------------------------
WScript.Echo "Found " & UBound(arrPlugins)+1 & " plugins" & vbNewLine

strPrevWorkingDir            = objWshShell.CurrentDirectory
objWshShell.CurrentDirectory = objFSO.GetParentFolderName(strGupPath)

For intCnt = 0 To UBound(arrPlugins)
  WScript.Echo "Processing plugin " & intCnt+1 & ": " & arrPlugins(intCnt).Item("folder-name")

  'Download and unzip plugin ZIP package with GUP.exe
  objWshShell.Run Quote(strGupPath) & " " & _
                    "-unzipTo " & _
                    Quote(strNppPath) & " " & _
                    Quote(strGupUnzipDirPath) & " " & _
                    Quote(arrPlugins(intCnt).Item("folder-name") & " " & _
                          arrPlugins(intCnt).Item("repository") & " " & _
                          arrPlugins(intCnt).Item("id")), _
                  1, _
                  True
Next

objWshShell.CurrentDirectory = strPrevWorkingDir


'-------------------------------------------------------------------------------
' Cleanup
'-------------------------------------------------------------------------------
If strPluginListDllPath <> "" And objFSO.FileExists(strPluginListJsonPath) Then
  Call objFSO.DeleteFile(strPluginListJsonPath)
End If

If objFSO.FileExists(strWaeDownloadPath) Then
  Call objFSO.DeleteFile(strWaeDownloadPath)
End If




'===============================================================================
' Retrieve command line arguments
'===============================================================================

Sub ParseCommandLine()
  Dim objFSO, colArgs

  Set objFSO  = CreateObject("Scripting.FileSystemObject")
  Set colArgs = WScript.Arguments

  If colArgs.Named.Exists("N") Then
    strNppDirPath         = objFSO.GetAbsolutePathName(colArgs.Named.Item("N"))
    strNppPath            = objFSO.BuildPath(strNppDirPath, "notepad++.exe")
    strGupPath            = objFSO.BuildPath(strNppDirPath, "updater\GUP.exe")
    strPluginListDllPath  = objFSO.BuildPath(strNppDirPath, "plugins\config\nppPluginList.dll")
  End If

  If colArgs.Named.Exists("G") Then
    strGupPath = objFSO.GetAbsolutePathName(colArgs.Named.Item("G"))
  End If

  If colArgs.Named.Exists("L") Then
    strPluginListDllPath = objFSO.GetAbsolutePathName(colArgs.Named.Item("L"))
  End If

  If colArgs.Named.Exists("J") Then
    strPluginListJsonPath = objFSO.GetAbsolutePathName(colArgs.Named.Item("J"))
    strPluginListDllPath  = ""
  End If

  If colArgs.Named.Exists("P") Then
    strGupUnzipDirPath = objFSO.GetAbsolutePathName(colArgs.Named.Item("P"))
  End If
End Sub


'===============================================================================
' Download file
'===============================================================================

Function DownloadFile(ByRef strURL, ByRef strDownloadPath, ByRef intStatusCode)
  Dim objXMLHTTP, objStream

  Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
  Set objStream  = CreateObject("ADODB.Stream")

  objXMLHTTP.open "GET", strURL, false
  objXMLHTTP.send()

  intStatusCode = objXMLHTTP.status

  If intStatusCode <> 200 Then
    DownloadFile = False
    Exit Function
  End If

  With objStream
    .Type = 1 'binary
    .Open
    .Write objXMLHTTP.responseBody
    .SaveToFile strDownloadPath, 2 'overwrite
    .Close
  End With

  DownloadFile = True
End Function


'===============================================================================
' Unzip file
'===============================================================================

Sub UnzipFile(ByRef strZipFilePath, ByRef strFileNameToUnzip, ByRef strDestPath)
  Dim objShell, intOldFileCount

  Set objShell = CreateObject("Shell.Application")

  If strFileNameToUnzip <> "" Then
    With objShell
      intOldFileCount = .Namespace(strDestPath).Items.Count

      .Namespace(strDestPath).CopyHere .Namespace(strZipFilePath).Items.Item(strFileNameToUnzip), &H0614

      Do Until .Namespace(strDestPath).Items.Count = intOldFileCount + 1
        WScript.Sleep 500
      Loop
    End With
  Else
    With objShell
      intOldFileCount = .Namespace(strDestPath).Items.Count

      .Namespace(strDestPath).CopyHere .Namespace(strZipFilePath).Items, &H0614

      Do Until .Namespace(strDestPath).Items.Count = intOldFileCount + .Namespace(strZipFilePath).Items.Count
        WScript.Sleep 500
      Loop
    End With
  End If
End Sub


'===============================================================================
' Extract plugin list from DLL file
'===============================================================================

Function ExtractPluginList(ByRef strDllFilePath, ByRef strPluginListJsonPath)
  Dim objWshShell

  Set objWshShell = CreateObject("WScript.Shell")

  objWshShell.Run Quote(strWaePath) & " " & _
                    "k@LoadLibraryExW $u:" & Quote(strDllFilePath) & " 0 0x40 , " & _
                    "k@FindResourceW $$:1 101 256 , " & _
                    "k@LoadResource $$:1 $$:6 , " & _
                    "k@SizeofResource $$:1 $$:6 , " & _
                    "k@LockResource $$:11 , " & _
                    "k@CreateFileW $u:" & Quote(strPluginListJsonPath) & " " & _
                                   "0xC0000000 0x01 0 2 0x80 0 , " & _
                    "k@WriteFile $$:22 $$:19 $$:15 $b:4 0 , " & _
                    "k@CloseHandle $$:22 , " & _
                    "k@FreeLibrary $$:1", _
                  0, _
                  True
                  
  ExtractPluginList = (objFSO.FileExists(strPluginListJsonPath))
End Function


'===============================================================================
' Surround a string with double quotes
'===============================================================================

Function Quote(ByRef strString)
  Quote = """" & strString & """"
End Function


'===============================================================================
' Include external code
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
