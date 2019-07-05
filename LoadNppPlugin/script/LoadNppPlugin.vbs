Option Explicit


'===============================================================================
' Base configuration
'===============================================================================

'Path to directory with supporting applications
Const BIN_PATH = ".\bin"

'Path to the directory where plugin files should be stored after unpacking
'them from the ZIP package
Const UNZIP_DIR_PATH = ".\plugins"

'URLs of plugin list JSON files
Const PLUGIN_LIST_X86_URL = "https://github.com/notepad-plus-plus/nppPluginList/raw/master/src/pl.x86.json"
Const PLUGIN_LIST_X64_URL = "https://github.com/notepad-plus-plus/nppPluginList/raw/master/src/pl.x64.json"

'===============================================================================


'Include external code
Include ".\include\IO.vbs"
Include ".\include\ClassJsonFile.vbs"


'Variables declaration
Dim objFSO, objWshShell, objJsonFile
Dim strBinDirPath, str7ZipPath, strCurlPath, strPluginDownloadURL, strProxyURL
Dim strPluginListDownloadPath, strPluginDownloadDirPath, strPluginDownloadPath
Dim strUnzipDirPath, strUnzipPath
Dim strPluginURL, strPluginName, intPluginNameLength
Dim intHTTPStatusCode, arrPlugins, intCnt, strInput, intInput


'Variables initalization
Set objFSO                = CreateObject("Scripting.FileSystemObject")
Set objWshShell           = CreateObject("WScript.Shell")
Set objJsonFile           = New clsJsonFile

strBinDirPath             = objFSO.GetAbsolutePathName(BIN_PATH)
strUnzipDirPath           = objFSO.GetAbsolutePathName(UNZIP_DIR_PATH)

str7ZipPath               = objFSO.BuildPath(strBinDirPath, "7za.exe")
strCurlPath               = objFSO.BuildPath(strBinDirPath, "curl.exe")
strPluginListDownloadPath = objFSO.BuildPath(objWshShell.ExpandEnvironmentStrings("%TEMP%"), "nppPluginList.json")
strPluginDownloadDirPath  = objWshShell.ExpandEnvironmentStrings("%TEMP%")


'-------------------------------------------------------------------------------
' Retrieve commandline arguments
'-------------------------------------------------------------------------------
Call ParseCommandline()


'-------------------------------------------------------------------------------
' Create unzip directory if necessary
'-------------------------------------------------------------------------------
If Not objFSO.FolderExists(strUnzipDirPath) Then
  Call ForceDirectories(strUnzipDirPath)
End If


'-------------------------------------------------------------------------------
' Download plugin list JSON file and parse it
'-------------------------------------------------------------------------------
If Not DownloadFile(strPluginDownloadURL, strPluginListDownloadPath, intHTTPStatusCode) Then
  WScript.Echo "Downloading plugin list failed. Return code: " & intHTTPStatusCode
  CleanupAndQuit
End If

'Ensure Windows EOL style of JSON file
Call ConvertUTF8EOLFormat(strPluginListDownloadPath, vbCrLf)

'Parse JSON data
If Not objJsonFile.LoadFromString(ReadUTF8File(strPluginListDownloadPath)) Then
  WScript.Echo "Failed to parse JSON file with plugin list."
  CleanupAndQuit
End If

arrPlugins = objJsonFile.Content.Item("npp-plugins")


'-------------------------------------------------------------------------------
' Print list of plugins and prompt user to select one
'-------------------------------------------------------------------------------
WScript.Echo "            ******************************************************"
WScript.Echo "            *                                                    *"
WScript.Echo "            *         Downloader for Notepad++ plugins           *"
WScript.Echo "            *                                                    *"
WScript.Echo "            * Released to the public domain 2019 by Andreas Heim *"
WScript.Echo "            *                                                    *"
WScript.Echo "            ******************************************************"
WScript.Echo
WScript.Echo

'Retrieve length of longest plugin name
intPluginNameLength = 0

For intCnt = 0 To UBound(arrPlugins)
  strPluginName = arrPlugins(intCnt).Item("display-name")

  If Len(strPluginName) > intPluginNameLength Then
    intPluginNameLength = Len(strPluginName)
  End If
Next

'Print plugin list (number, plugin name, plugin URL)
For intCnt = 0 To UBound(arrPlugins)
  strPluginURL  = arrPlugins(intCnt).Item("repository")
  strPluginName = arrPlugins(intCnt).Item("display-name")

  WScript.Echo "[ " & String(Len(CStr(UBound(arrPlugins))) - Len(CStr(intCnt)), " ") & intCnt & " ]  " & _
               strPluginName & String(intPluginNameLength - Len(strPluginName) + 2, " ") & _
               strPluginURL
Next

WScript.Echo
WScript.Echo

Do While True
  WScript.StdOut.Write "Select plugin to download by its number (ENTER to cancel): "
  strInput = WScript.StdIn.ReadLine

  If strInput = "" Then CleanupAndQuit

  If IsNumeric(strInput) Then
    intInput = Fix(Abs(strInput))
    If intInput <= UBound(arrPlugins) Then Exit Do
  End If

  WScript.Echo "Please enter a valid number."
  WScript.Echo
Loop

WScript.Echo
WScript.Echo "Attempt to download selected plugin, please wait..."
WScript.Echo


'-------------------------------------------------------------------------------
' Download and unzip selected plugin package
'-------------------------------------------------------------------------------
strPluginURL  = arrPlugins(intInput).Item("repository")
strPluginName = arrPlugins(intInput).Item("folder-name")

strPluginDownloadPath = objFSO.BuildPath(strPluginDownloadDirPath, strPluginName & ".zip")

If Not DownloadFile(strPluginURL, strPluginDownloadPath, intHTTPStatusCode) Then
  WScript.Echo "Downloading plugin " & strPluginName & " failed. Return code: " & intHTTPStatusCode
  CleanupAndQuit
Else
  strUnzipPath = objFSO.BuildPath(strUnzipDirPath, strPluginName)

  Call UnzipFile(strPluginDownloadPath, strUnzipPath)
  Call objFSO.DeleteFile(strPluginDownloadPath)
End If

WScript.Echo "Plugin has been successfully downloaded."
WScript.Echo "See " & strUnzipPath


'-------------------------------------------------------------------------------
' Cleanup
'-------------------------------------------------------------------------------
Call CleanupAndQuit()



'===============================================================================
' Retrieve commandline arguments
'===============================================================================

Sub ParseCommandline
  Dim intCnt

  strPluginDownloadURL = PLUGIN_LIST_X86_URL
  strProxyURL          = ""

  For intCnt = 0 To WScript.Arguments.Unnamed.Count - 1
    If StrComp(WScript.Arguments.Unnamed(intCnt), "x86", vbTextCompare) = 0 Then
      strPluginDownloadURL = PLUGIN_LIST_X86_URL
      strUnzipDirPath      = objFSO.BuildPath(strUnzipDirPath, "x86")
      
    ElseIf StrComp(WScript.Arguments.Unnamed(intCnt), "x64", vbTextCompare) = 0 Then
      strPluginDownloadURL = PLUGIN_LIST_X64_URL
      strUnzipDirPath      = objFSO.BuildPath(strUnzipDirPath, "x64")
    End If
  Next

  If WScript.Arguments.Named.Exists("P") Then
    strProxyURL = WScript.Arguments.Named("P")
  End If
End Sub


'===============================================================================
' Download file
'===============================================================================

Function DownloadFile(ByRef strFileURL, ByRef strDstPath, ByRef intStatusCode)
  Dim objShell, strProxyArg

  Set objShell = CreateObject("WScript.Shell")

  If strProxyURL <> "" Then
    strProxyArg = "-x " & strProxyURL
  Else
    strProxyArg = ""
  End If

  intStatusCode = objShell.Run(Quote(strCurlPath) & " " & _
                                 strProxyArg & " " & _
                                 strFileURL & " " & _
                                 "--silent " & _
                                 "--location " & _
                                 "--output " & Quote(strDstPath), _
                               0, _
                               True)

  DownloadFile = (intStatusCode = 0)
End Function


'===============================================================================
' Unzip a plugin ZIP file
'===============================================================================

Sub UnzipFile(ByRef strZipFilePath, ByRef strDstFolder)
  Dim objFSO, objShell

  Set objFSO   = CreateObject("Scripting.FileSystemObject")
  Set objShell = CreateObject("WScript.Shell")

  If Not objFSO.FolderExists(strDstFolder) Then
    Call objFSO.CreateFolder(strDstFolder)
  End If

  objShell.Run Quote(str7ZipPath) & " " & _
                 "x " & Quote(strZipFilePath) & " " & _
                 "-o" & Quote(strDstFolder) & " " & _
                 "-r" & " " & _
                 "-aoa", _
               0, _
               True
End Sub


'===============================================================================
' Create a nested directory structure
'===============================================================================

Function ForceDirectories(ByRef strPath)
  Dim objFSO, strPartPath, strAbsPath, arrAbsPath, intCnt

  Set objFSO  = CreateObject("Scripting.FileSystemObject")

  strAbsPath  = objFSO.GetAbsolutePathName(strPath)
  arrAbsPath  = Split(strAbsPath, "\")
  strPartPath = objFSO.BuildPath(arrAbsPath(0), "\")

  If Not objFSO.DriveExists(strPartPath) Then
    ForceDirectories = False
  Else
    For intCnt = 1 To UBound(arrAbsPath)
      strPartPath = objFSO.BuildPath(strPartPath, arrAbsPath(intCnt))

      If Not objFSO.FolderExists(strPartPath) Then
        objFSO.CreateFolder(strPartPath)
      End If
    Next
  End If

  ForceDirectories = True
End Function


'===============================================================================
' Cleanup and terminate script
'===============================================================================

Sub CleanupAndQuit
  If objFSO.FileExists(strPluginListDownloadPath) Then
    Call objFSO.DeleteFile(strPluginListDownloadPath)
  End If

  WScript.Echo
  WScript.Quit
End Sub


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
  Dim objWshShell, objFSO, objFileStream, strScriptFilePath, strAbsFilePath, strCode

  Set objWshShell   = CreateObject("WScript.Shell")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  strScriptFilePath = objWshShell.CurrentDirectory
  strAbsFilePath    = objFSO.GetAbsolutePathName(objFSO.BuildPath(strScriptFilePath, strFilePath))

  Set objFileStream = objFSO.OpenTextFile(strAbsFilePath, 1, False, 0)
  strCode           = objFileStream.ReadAll
  objFileStream.Close

  ExecuteGlobal strCode
End Sub
