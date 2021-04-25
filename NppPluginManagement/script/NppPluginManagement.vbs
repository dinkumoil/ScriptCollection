Option Explicit


'===============================================================================
' Base configuration
'===============================================================================

'Path to directory with supporting applications
Const BIN_PATH = ".\bin"

'Path to the directory where plugin files should be stored after unpacking
'them from the ZIP package
Const UNZIP_DIR_PATH = ".\plugins"

Const NPP_ARCHITECTURE_X86 = "x86"
Const NPP_ARCHITECTURE_X64 = "x64"

Const PROGRAM_FILES_PATH_X86_X86 = "%ProgramFiles%"
Const PROGRAM_FILES_PATH_X86_X64 = "%ProgramFiles(x86)%"
Const PROGRAM_FILES_PATH_X64     = "%ProgramW6432%"

'URLs of plugin list JSON files
Const PLUGIN_LIST_X86_URL = "https://github.com/notepad-plus-plus/nppPluginList/raw/master/src/pl.x86.json"
Const PLUGIN_LIST_X64_URL = "https://github.com/notepad-plus-plus/nppPluginList/raw/master/src/pl.x64.json"

'===============================================================================


'Result code constants
Const RES_NO_ERROR                      = 0
Const RES_TASK_COMPLETED                = 1
Const RES_USER_ABORTED                  = 2
Const RES_PLUGINLIST_DOWNLOAD_FAILED    = 3
Const RES_PLUGINLIST_PARSE_FAILED       = 4
Const RES_NPP_PLUGINDIRECTORY_NOT_FOUND = 5
Const RES_ALL_PLUGINS_UPTODATE          = 6

'Command constants
Const CMD_INVALID  = -1
Const CMD_DOWNLOAD = 0
Const CMD_INSTALL  = 1
Const CMD_UPDATE   = 2
Const CMD_LIST     = 3


'Include external code
Include ".\include\IO.vbs"
Include ".\include\ClassFileVersionInfo.vbs"
Include ".\include\ClassJsonFile.vbs"
Include ".\include\ClassNppPlugin.vbs"


'Variables declaration
Dim objFSO, objWshShell, objJsonFile
Dim strBinDirPath, str7ZipPath, strCurlPath, strWAEPath, strNppPluginDirPath
Dim strNppArchitecture, strPluginListDownloadURL, strProxyURL, intHTTPStatusCode
Dim strPluginListDownloadPath, strPluginDownloadDirPath, strUnzipDirPath
Dim arrPlugins, arrPluginsData, intSelection, bolUpdateAll, bolPreSelTask, intResult


'Variables initalization
Set objFSO                = CreateObject("Scripting.FileSystemObject")
Set objWshShell           = CreateObject("WScript.Shell")
Set objJsonFile           = New clsJsonFile

strBinDirPath             = objFSO.GetAbsolutePathName(BIN_PATH)
strUnzipDirPath           = objFSO.GetAbsolutePathName(UNZIP_DIR_PATH)

str7ZipPath               = objFSO.BuildPath(strBinDirPath, "7za.exe")
strCurlPath               = objFSO.BuildPath(strBinDirPath, "curl.exe")
strWAEPath                = objFSO.BuildPath(strBinDirPath, "winapiexec.exe")
strPluginListDownloadPath = objFSO.BuildPath(objWshShell.ExpandEnvironmentStrings("%TEMP%"), "nppPluginList.json")
strPluginDownloadDirPath  = objWshShell.ExpandEnvironmentStrings("%TEMP%")


'-------------------------------------------------------------------------------
' Retrieve and check commandline arguments
'-------------------------------------------------------------------------------
Call ParseCommandline()
Call CheckCommandlineArguments()


'-------------------------------------------------------------------------------
' Read user input to select task
'-------------------------------------------------------------------------------
WScript.Echo "            ******************************************************"
WScript.Echo "            *                                                    *"
WScript.Echo "            *          Plugin management for Notepad++           *"
WScript.Echo "            *                                                    *"
WScript.Echo "            * Released to the public domain 2020 by Andreas Heim *"
WScript.Echo "            *                                                    *"
WScript.Echo "            ******************************************************"
WScript.Echo
WScript.Echo

If objFSO.FolderExists(strNppPluginDirPath) Then
  WScript.Echo "Processing Notepad++ installation: " & strNppPluginDirPath
  WScript.Echo "Version                          : " & GetVersionNumber(objFSO.BuildPath(strNppPluginDirPath, "..\notepad++.exe"))
  WScript.Echo "Architecture                     : " & GetImageArchitecture(objFSO.BuildPath(strNppPluginDirPath, "..\notepad++.exe"))
  WScript.Echo
  WScript.Echo

  Do While intSelection = CMD_INVALID
    WScript.Echo "Please select a task."
    WScript.Echo
    WScript.Echo "  (d) List and download available Notepad++ plugins"
    WScript.Echo "  (i) List and install available Notepad++ plugins"
    WScript.Echo "  (l) List already installed Notepad++ plugins"
    WScript.Echo "  (u) List and update already installed Notepad++ plugins"
    WScript.Echo "  (a) List and update all already installed Notepad++ plugins"
    WScript.Echo "  (x) Exit"
    WScript.Echo

    WScript.StdOut.Write "Your selection: "

    Select Case LCase(WScript.StdIn.ReadLine)
      Case "d"  intSelection = CMD_DOWNLOAD
      Case "i"  intSelection = CMD_INSTALL
      Case "u"  intSelection = CMD_UPDATE
      Case "a"  intSelection = CMD_UPDATE
                bolUpdateAll = True
      Case "l"  intSelection = CMD_LIST
      Case "x"  Call CleanupAndQuit(RES_USER_ABORTED)
      Case Else WScript.Echo "Please select a valid item."
    End Select

    WScript.Echo
    WScript.Echo
  Loop
Else
  WScript.Echo "Processing Notepad++ architecture: " & strNppArchitecture
  WScript.Echo
  WScript.Echo

  Do While intSelection = CMD_INVALID
    WScript.Echo "Please select a task."
    WScript.Echo
    WScript.Echo "  (d) List and download available Notepad++ plugins"
    WScript.Echo "  (x) Exit"
    WScript.Echo

    WScript.StdOut.Write "Your selection: "

    Select Case LCase(WScript.StdIn.ReadLine)
      Case "d"  intSelection = CMD_DOWNLOAD
      Case "x"  Call CleanupAndQuit(RES_USER_ABORTED)
      Case Else WScript.Echo "Please select a valid item."
    End Select

    WScript.Echo
    WScript.Echo
  Loop
End If


'-------------------------------------------------------------------------------
' Download plugin list JSON file and parse it
'-------------------------------------------------------------------------------
If Not DownloadFile(strPluginListDownloadURL, strPluginListDownloadPath, intHTTPStatusCode) Then
  WScript.Echo "Downloading plugin list failed. Return code: " & intHTTPStatusCode
  Call CleanupAndQuit(RES_PLUGINLIST_DOWNLOAD_FAILED)
End If

'Ensure Windows EOL style of JSON file
Call ConvertUTF8EOLFormat(strPluginListDownloadPath, vbCrLf)

'Parse JSON data
If Not objJsonFile.LoadFromString(ReadUTF8File(strPluginListDownloadPath)) Then
  WScript.Echo "Failed to parse JSON file with plugin list."
  Call CleanupAndQuit(RES_PLUGINLIST_PARSE_FAILED)
End If


'-------------------------------------------------------------------------------
' Retrieve data for all plugins in list
'-------------------------------------------------------------------------------
arrPlugins     = objJsonFile.Content.Item("npp-plugins")
arrPluginsData = GetPluginsData(arrPlugins)

If objFSO.FolderExists(strNppPluginDirPath) Then
  arrPluginsData = GetPluginsExtData(strNppPluginDirPath, arrPluginsData)
End If


'-------------------------------------------------------------------------------
' Perform selected action
'-------------------------------------------------------------------------------
Select Case intSelection
  Case CMD_DOWNLOAD  intResult = ListAndDownloadPlugin(strPluginDownloadDirPath, strUnzipDirPath, arrPluginsData)
  Case CMD_INSTALL   intResult = ListAndInstallPlugin(strPluginDownloadDirPath, strUnzipDirPath, strNppPluginDirPath, arrPluginsData)
  Case CMD_UPDATE    intResult = ListAndUpdatePlugin(strPluginDownloadDirPath, strUnzipDirPath, strNppPluginDirPath, arrPluginsData)
  Case CMD_LIST      intResult = ListInstalledPlugins(arrPluginsData)
End Select


'-------------------------------------------------------------------------------
' Cleanup
'-------------------------------------------------------------------------------
If bolPreSelTask Then
  Call CleanupAndQuit(intResult)
Else
  Call CleanupAndQuit(RES_NO_ERROR)
End If




'===============================================================================
' Retrieve commandline arguments
'===============================================================================

Sub ParseCommandline
  Dim strProgramFilesPath, strNppBasePath

  strNppArchitecture       = NPP_ARCHITECTURE_X86
  strPluginListDownloadURL = PLUGIN_LIST_X86_URL
  strProgramFilesPath      = GetProgramFilesPath32Bit()
  intSelection             = CMD_INVALID
  bolUpdateAll             = False
  bolPreSelTask            = False
  strProxyURL              = ""

  If WScript.Arguments.Named.Exists("T") Then
    bolPreSelTask = True

    Select Case LCase(WScript.Arguments.Named("T"))
      Case "download"  intSelection  = CMD_DOWNLOAD
      Case "install"   intSelection  = CMD_INSTALL
      Case "update"    intSelection  = CMD_UPDATE
      Case "updateall" intSelection  = CMD_UPDATE
                       bolUpdateAll  = True
      Case "list"      intSelection  = CMD_LIST
      Case Else        bolPreSelTask = False
    End Select
  End If

  If WScript.Arguments.Named.Exists("A") Then
    If StrComp(WScript.Arguments.Named("A"), "x86", vbTextCompare) = 0 Then
      strNppArchitecture       = NPP_ARCHITECTURE_X86
      strPluginListDownloadURL = PLUGIN_LIST_X86_URL
      strProgramFilesPath      = GetProgramFilesPath32Bit()

    ElseIf StrComp(WScript.Arguments.Named("A"), "x64", vbTextCompare) = 0 Then
      strNppArchitecture       = NPP_ARCHITECTURE_X64
      strPluginListDownloadURL = PLUGIN_LIST_X64_URL
      strProgramFilesPath      = PROGRAM_FILES_PATH_X64
    End If
  End If

  If WScript.Arguments.Named.Exists("N") Then
    strNppBasePath = WScript.Arguments.Named("N")
  Else
    strProgramFilesPath = objWshShell.ExpandEnvironmentStrings(strProgramFilesPath)
    strNppBasePath      = objFSO.BuildPath(strProgramFilesPath, "Notepad++")
  End If

  If WScript.Arguments.Named.Exists("P") Then
    strProxyURL = WScript.Arguments.Named("P")
  End If

  strNppPluginDirPath = objFSO.BuildPath(strNppBasePath, "plugins")
  strUnzipDirPath     = objFSO.BuildPath(strUnzipDirPath, strNppArchitecture)
End Sub


'===============================================================================
' Check commandline arguments
'===============================================================================

Sub CheckCommandlineArguments
  If intSelection <> CMD_INVALID And intSelection <> CMD_DOWNLOAD And _
     Not objFSO.FolderExists(strNppPluginDirPath)                 Then
    WScript.Echo "Notepad++ plugin directory not found: "
    WScript.Echo
    WScript.Echo "  " & strNppPluginDirPath

    Call CleanupAndQuit(RES_NPP_PLUGINDIRECTORY_NOT_FOUND)
  End If
End Sub


'===============================================================================
' List plugins and download one of them
'===============================================================================

Function ListAndDownloadPlugin(ByRef strPluginDownloadDirPath, ByRef strUnzipDirPath, ByRef arrPluginsData)
  Dim intPluginNameLength, intPluginVersionLength
  Dim intCnt, intInput, intHTTPStatusCode
  Dim strPluginName, strPluginVersion, strPluginURL
  Dim strPluginDownloadPath, strUnzipPath

  ListAndDownloadPlugin = RES_NO_ERROR

  '-----------------------------------------------------------------------------
  ' Print list of plugins and prompt user to select one
  '-----------------------------------------------------------------------------
  'Retrieve length of longest column content
  intPluginNameLength    = 0
  intPluginVersionLength = 0

  For intCnt = 0 To UBound(arrPluginsData)
    strPluginName    = arrPluginsData(intCnt).DisplayName
    strPluginVersion = arrPluginsData(intCnt).Version

    If Len(strPluginName) > intPluginNameLength Then
      intPluginNameLength = Len(strPluginName)
    End If

    If Len(strPluginVersion) > intPluginVersionLength Then
      intPluginVersionLength = Len(strPluginVersion)
    End If
  Next

  'Print plugin list (number, plugin name, version, plugin URL)
  For intCnt = 0 To UBound(arrPluginsData)
    strPluginURL     = arrPluginsData(intCnt).Repository
    strPluginName    = arrPluginsData(intCnt).DisplayName
    strPluginVersion = arrPluginsData(intCnt).Version

    WScript.Echo "[ " & String(Len(CStr(UBound(arrPluginsData)))  - Len(CStr(intCnt)),         " ") & intCnt & " ]  " & _
                 strPluginName    & String(intPluginNameLength    - Len(strPluginName)    + 2, " ") & _
                 strPluginVersion & String(intPluginVersionLength - Len(strPluginVersion) + 2, " ") & _
                 strPluginURL
  Next

  intInput = ReadUserInput("Select plugin to download by its number", UBound(arrPluginsData))

  If intInput = -1 Then
    ListAndDownloadPlugin = RES_USER_ABORTED
    Exit Function
  End If

  WScript.Echo vbNewLine
  WScript.Echo "Attempt to download selected plugin, please wait..."
  WScript.Echo

  '-----------------------------------------------------------------------------
  ' Download and unzip selected plugin package
  '-----------------------------------------------------------------------------
  strPluginURL  = arrPluginsData(intInput).Repository
  strPluginName = arrPluginsData(intInput).FolderName

  strPluginDownloadPath = objFSO.BuildPath(strPluginDownloadDirPath, strPluginName & ".zip")

  If Not DownloadFile(strPluginURL, strPluginDownloadPath, intHTTPStatusCode) Then
    WScript.Echo "Downloading plugin " & strPluginName & " failed. Return code: " & intHTTPStatusCode
    Exit Function
  Else
    strUnzipPath = objFSO.BuildPath(strUnzipDirPath, strPluginName)
    If objFSO.FolderExists(strUnzipPath) Then Call DeleteDirTree(strUnzipPath)

    Call UnzipFile(strPluginDownloadPath, strUnzipPath)
    Call objFSO.DeleteFile(strPluginDownloadPath)
  End If

  WScript.Echo "Plugin has been successfully downloaded."
  WScript.Echo "See " & strUnzipPath
End Function


'===============================================================================
' List plugins and install one of them
'===============================================================================

Function ListAndInstallPlugin(ByRef strPluginDownloadDirPath, ByRef strUnzipDirPath, ByRef strNppPluginDirPath, ByRef arrPluginsData)
  Dim intPluginNameLength, intPluginVersionLength
  Dim strPluginName, strPluginVersion, strPluginURL
  Dim intCnt, intInput, intHTTPStatusCode
  Dim strPluginDownloadPath, strUnzipPath
  Dim objShell, strApplication, strArguments

  ListAndInstallPlugin = RES_NO_ERROR

  '-----------------------------------------------------------------------------
  ' Print list of plugins and prompt user to select one
  '-----------------------------------------------------------------------------
  'Retrieve length of longest column content
  intPluginNameLength    = 0
  intPluginVersionLength = 0

  For intCnt = 0 To UBound(arrPluginsData)
    strPluginName    = arrPluginsData(intCnt).DisplayName
    strPluginVersion = arrPluginsData(intCnt).Version

    If Len(strPluginName) > intPluginNameLength Then
      intPluginNameLength = Len(strPluginName)
    End If

    If Len(strPluginVersion) > intPluginVersionLength Then
      intPluginVersionLength = Len(strPluginVersion)
    End If
  Next

  'Print plugin list (number, plugin name, version, plugin URL)
  For intCnt = 0 To UBound(arrPluginsData)
    strPluginURL     = arrPluginsData(intCnt).Repository
    strPluginName    = arrPluginsData(intCnt).DisplayName
    strPluginVersion = arrPluginsData(intCnt).Version

    WScript.Echo "[ " & String(Len(CStr(UBound(arrPluginsData)))  - Len(CStr(intCnt)),         " ") & intCnt & " ]  " & _
                 strPluginName    & String(intPluginNameLength    - Len(strPluginName)    + 2, " ") & _
                 strPluginVersion & String(intPluginVersionLength - Len(strPluginVersion) + 2, " ") & _
                 strPluginURL
  Next

  intInput = ReadUserInput("Select plugin to install by its number", UBound(arrPluginsData))

  If intInput = -1 Then
    ListAndInstallPlugin = RES_USER_ABORTED
    Exit Function
  End If

  WScript.Echo vbNewLine
  WScript.Echo "Attempt to download selected plugin, please wait..."
  WScript.Echo

  '-----------------------------------------------------------------------------
  ' Download and unzip selected plugin package
  '-----------------------------------------------------------------------------
  strPluginURL  = arrPluginsData(intInput).Repository
  strPluginName = arrPluginsData(intInput).FolderName

  strPluginDownloadPath = objFSO.BuildPath(strPluginDownloadDirPath, strPluginName & ".zip")

  If Not DownloadFile(strPluginURL, strPluginDownloadPath, intHTTPStatusCode) Then
    WScript.Echo "Downloading plugin " & strPluginName & " failed. Return code: " & intHTTPStatusCode
    Exit Function
  Else
    strUnzipPath = objFSO.BuildPath(strUnzipDirPath, strPluginName)
    If objFSO.FolderExists(strUnzipPath) Then Call DeleteDirTree(strUnzipPath)

    Call UnzipFile(strPluginDownloadPath, strUnzipPath)
    Call objFSO.DeleteFile(strPluginDownloadPath)
  End If

  WScript.Echo "Plugin has been successfully downloaded."
  WScript.Echo "See " & strUnzipPath
  WScript.Echo

  '-----------------------------------------------------------------------------
  ' Install selected plugin package
  '-----------------------------------------------------------------------------
  Set objShell = CreateObject("Shell.Application")

  strApplication = "cscript.exe"

  If IsModernNpp(objFSO.BuildPath(strNppPluginDirPath, "..\notepad++.exe")) Then
    strArguments   = "/nologo " & Quote(objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), "PluginInstaller.vbs")) & " " & _
                     "/p:" & Quote(strUnzipDirPath) & " " & _
                     "/n:" & Quote(strNppPluginDirPath)
  Else
    strArguments   = "/nologo " & Quote(objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), "PluginInstaller.vbs")) & " " & _
                     "/p:" & Quote(strUnzipPath) & " " & _
                     "/n:" & Quote(strNppPluginDirPath)
  End If

  objShell.ShellExecute strApplication, strArguments, "", "runas", 0

  WScript.Echo "Plugin has been moved to " & Quote(strNppPluginDirPath)
  WScript.Echo
End Function


'===============================================================================
' List plugins and update one of them
'===============================================================================

Function ListAndUpdatePlugin(ByRef strPluginDownloadDirPath, ByRef strUnzipDirPath, ByRef strNppPluginDirPath, ByRef arrPluginsData)
  Dim intPluginNameLength, intPluginInstVersionLength, intPluginVersionLength
  Dim strPluginName, strPluginURL, strPluginInstVersion, strPluginVersion
  Dim arrPlugins, intCnt, intPluginCnt, intInput, intHTTPStatusCode
  Dim strPluginDownloadPath, strUnzipPath
  Dim objShell, strApplication, strArguments
  Dim bolContinue

  ListAndUpdatePlugin = RES_NO_ERROR

  '-----------------------------------------------------------------------------
  ' Print list of plugins and prompt user to select one
  '-----------------------------------------------------------------------------
  'Retrieve length of longest column content
  intPluginNameLength        = 0
  intPluginInstVersionLength = 0
  intPluginVersionLength     = 0
  intPluginCnt               = 0
  arrPlugins                 = Array()

  For intCnt = 0 To UBound(arrPluginsData)
    If arrPluginsData(intCnt).IsUpdateAvailable Then
      strPluginName        = arrPluginsData(intCnt).DisplayName
      strPluginInstVersion = arrPluginsData(intCnt).InstalledVersion
      strPluginVersion     = arrPluginsData(intCnt).Version

      If Len(strPluginName) > intPluginNameLength Then
        intPluginNameLength = Len(strPluginName)
      End If

      If Len(strPluginInstVersion) > intPluginInstVersionLength Then
        intPluginInstVersionLength = Len(strPluginInstVersion)
      End If

      If Len(strPluginVersion) > intPluginVersionLength Then
        intPluginVersionLength = Len(strPluginVersion)
      End If

      ReDim Preserve arrPlugins(intPluginCnt)
      Set arrPlugins(intPluginCnt) = arrPluginsData(intCnt)
      intPluginCnt = intPluginCnt + 1
    End If
  Next

  If UBound(arrPlugins) = -1 Then
    WScript.Echo
    WScript.Echo "All plugins are up to date."
    ListAndUpdatePlugin = RES_ALL_PLUGINS_UPTODATE
    Exit Function
  End If

  'Print plugin list (number, plugin name, installed version, repo version, plugin URL)
  For intCnt = 0 To UBound(arrPlugins)
    strPluginURL         = arrPlugins(intCnt).Repository
    strPluginName        = arrPlugins(intCnt).DisplayName
    strPluginInstVersion = arrPlugins(intCnt).InstalledVersion
    strPluginVersion     = arrPlugins(intCnt).Version

    WScript.Echo "[ " & String(Len(CStr(UBound(arrPlugins)))               - Len(CStr(intCnt)),             " ") & intCnt & " ]  " & _
                 strPluginName         & String(intPluginNameLength        - Len(strPluginName)        + 2, " ") & _
                 strPluginInstVersion  & String(intPluginInstVersionLength - Len(strPluginInstVersion) + 2, " ") & _
                 strPluginVersion      & String(intPluginVersionLength     - Len(strPluginVersion)     + 2, " ") & _
                 strPluginURL
  Next

  If not bolUpdateAll Then
    intInput = ReadUserInput("Select plugin to update by its number", UBound(arrPlugins))

    If intInput = -1 Then
      ListAndUpdatePlugin = RES_USER_ABORTED
      Exit Function
    End If

    WScript.Echo vbNewLine
    WScript.Echo "Attempt to download selected plugin, please wait..."
    WScript.Echo
  Else
    WScript.Echo vbNewLine
    WScript.Echo "Start downloading and updating plugins, please wait..."
    WScript.Echo vbNewLine
  End If

  '-----------------------------------------------------------------------------
  ' Download and unzip all or only selected plugin package
  '-----------------------------------------------------------------------------
  For intCnt = 0 To UBound(arrPlugins)
    bolContinue = False

    If bolUpdateAll Or intCnt = intInput Then
      '-------------------------------------------------------------------------
      ' Download and unzip plugin package
      '-------------------------------------------------------------------------
      strPluginURL  = arrPlugins(intCnt).Repository
      strPluginName = arrPlugins(intCnt).FolderName

      strPluginDownloadPath = objFSO.BuildPath(strPluginDownloadDirPath, strPluginName & ".zip")

      If Not DownloadFile(strPluginURL, strPluginDownloadPath, intHTTPStatusCode) Then
        WScript.Echo "Downloading plugin " & strPluginName & " failed. Return code: " & intHTTPStatusCode
        bolContinue = True
      Else
        strUnzipPath = objFSO.BuildPath(strUnzipDirPath, strPluginName)
        If objFSO.FolderExists(strUnzipPath) Then Call DeleteDirTree(strUnzipPath)

        Call UnzipFile(strPluginDownloadPath, strUnzipPath)
        Call objFSO.DeleteFile(strPluginDownloadPath)
      End If

      If Not bolContinue Then
        WScript.Echo "Plugin has been successfully downloaded."
        WScript.Echo "See " & strUnzipPath
        WScript.Echo

        '-----------------------------------------------------------------------
        ' Install plugin package
        '-----------------------------------------------------------------------
        Set objShell = CreateObject("Shell.Application")

        strApplication = "cscript.exe"

        If IsModernNpp(objFSO.BuildPath(strNppPluginDirPath, "..\notepad++.exe")) Then
          strArguments   = "/nologo " & Quote(objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), "PluginInstaller.vbs")) & " " & _
                           "/p:" & Quote(strUnzipDirPath) & " " & _
                           "/n:" & Quote(strNppPluginDirPath)
        Else
          strArguments   = "/nologo " & Quote(objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), "PluginInstaller.vbs")) & " " & _
                           "/p:" & Quote(strUnzipPath) & " " & _
                           "/n:" & Quote(strNppPluginDirPath)
        End If

        objShell.ShellExecute strApplication, strArguments, "", "runas", 0

        WScript.Echo "Plugin has been moved to " & Quote(strNppPluginDirPath)
        WScript.Echo
      End If

      If bolUpdateAll Then WScript.Echo "" Else Exit For
    End If
  Next

  If bolUpdateAll Then
    ListAndUpdatePlugin = RES_TASK_COMPLETED
  End If
End Function


'===============================================================================
' List installed plugins
'===============================================================================

Function ListInstalledPlugins(ByRef arrPluginsData)
  Dim intPluginNameLength, intPluginInstVersionLength
  Dim strPluginName, strPluginURL, strPluginInstVersion
  Dim arrPlugins, intCnt, intPluginCnt

  ListInstalledPlugins = RES_USER_ABORTED

  '-----------------------------------------------------------------------------
  ' Print list of plugins
  '-----------------------------------------------------------------------------
  'Retrieve length of longest column content
  intPluginNameLength        = 0
  intPluginInstVersionLength = 0
  intPluginCnt               = 0
  arrPlugins                 = Array()

  For intCnt = 0 To UBound(arrPluginsData)
    If arrPluginsData(intCnt).IsInstalled Then
      strPluginName        = arrPluginsData(intCnt).DisplayName
      strPluginInstVersion = arrPluginsData(intCnt).InstalledVersion

      If Len(strPluginName) > intPluginNameLength Then
        intPluginNameLength = Len(strPluginName)
      End If

      If Len(strPluginInstVersion) > intPluginInstVersionLength Then
        intPluginInstVersionLength = Len(strPluginInstVersion)
      End If

      ReDim Preserve arrPlugins(intPluginCnt)
      Set arrPlugins(intPluginCnt) = arrPluginsData(intCnt)
      intPluginCnt = intPluginCnt + 1
    End If
  Next

  'Print plugin list (number, plugin name, installed version, plugin URL)
  For intCnt = 0 To UBound(arrPlugins)
    strPluginURL         = arrPlugins(intCnt).Repository
    strPluginName        = arrPlugins(intCnt).DisplayName
    strPluginInstVersion = arrPlugins(intCnt).InstalledVersion

    WScript.Echo "[ " & String(Len(CStr(UBound(arrPlugins)))               - Len(CStr(intCnt)),             " ") & intCnt & " ]  " & _
                 strPluginName         & String(intPluginNameLength        - Len(strPluginName)        + 2, " ") & _
                 strPluginInstVersion  & String(intPluginInstVersionLength - Len(strPluginInstVersion) + 2, " ") & _
                 strPluginURL
  Next
End Function


'===============================================================================
' Read user input
'===============================================================================

Function ReadUserInput(ByRef strMessage, intMaxValue)
  Dim strInput, intInput

  ReadUserInput = -1

  WScript.Echo
  WScript.Echo

  Do While True
    WScript.StdOut.Write strMessage & " (ENTER to cancel): "
    strInput = WScript.StdIn.ReadLine

    If strInput = "" Then Exit Function

    If IsNumeric(strInput) Then
      intInput = Fix(Abs(strInput))
      If intInput <= intMaxValue Then Exit Do
    End If

    WScript.Echo "Please enter a valid number."
    WScript.Echo
  Loop

  ReadUserInput = intInput
End Function


'===============================================================================
' Store basic plugin data into array
'===============================================================================

Function GetPluginsData(ByRef arrPlugins)
  Dim intCnt, arrResult()

  ReDim arrResult(UBound(arrPlugins))

  For intCnt = 0 To UBound(arrPlugins)
    Set arrResult(intCnt) = New clsNppPlugin

    With arrResult(intCnt)
      .DisplayName = arrPlugins(intCnt).Item("display-name")
      .FolderName  = arrPlugins(intCnt).Item("folder-name")
      .Repository  = arrPlugins(intCnt).Item("repository")
      .Version     = arrPlugins(intCnt).Item("version")
    End With
  Next

  GetPluginsData = arrResult
End Function


'===============================================================================
' Store extended plugin data into array
'===============================================================================

Function GetPluginsExtData(ByRef strPluginDirPath, ByRef arrPluginsData)
  Dim bolIsModernNpp, intCnt

  bolIsModernNpp = IsModernNpp(objFSO.BuildPath(strPluginDirPath, "..\notepad++.exe"))

  For intCnt = 0 To UBound(arrPluginsData)
    With arrPluginsData(intCnt)
      If bolIsModernNpp Then
        .InstallPath = objFSO.BuildPath(strPluginDirPath, .FolderName)
      Else
        .InstallPath = strPluginDirPath
      End If

      .InstallPath = objFSO.BuildPath(.InstallPath, .FolderName & ".dll")
      .IsInstalled = objFSO.FileExists(.InstallPath)

      If Not .IsInstalled Then
        .InstalledVersion  = "0.0"
        .IsUpdateAvailable = False
      Else
        .InstalledVersion  = GetVersionNumber(.InstallPath)
        .IsUpdateAvailable = (CompareVersion(.InstalledVersion, .Version) = -1)
      End If
    End With
  Next

  GetPluginsExtData = arrPluginsData
End Function


'===============================================================================
' Check if Notepad++ installation has version number 7.6 or higher
'===============================================================================

Function IsModernNpp(ByRef strNppPath)
  IsModernNpp = (CompareVersion(GetVersionNumber(strNppPath), "7.6") <> -1)
End Function


'===============================================================================
' Get version number of EXE or DLL file
'===============================================================================

Function GetVersionNumber(ByRef strFilePath)
  Dim objVersionInfo

  Set objVersionInfo      = New clsFileVersionInfo
  objVersionInfo.FilePath = objFSO.GetAbsolutePathName(strFilePath)

  GetVersionNumber = objVersionInfo.FileVersionInfoTag(FVIFileVersion)
End Function


'===============================================================================
' Compare version info
'===============================================================================

Function CompareVersion(ByRef strLeft, ByRef strRight)
  Dim arrLeft, arrRight, intMinDimCnt, intMaxDimCnt, intCnt, intResult

  intResult = 0

  arrLeft  = Split(strLeft,  ".")
  arrRight = Split(strRight, ".")

  If UBound(arrLeft) > UBound(arrRight) Then
    intMinDimCnt = UBound(arrRight)
    intMaxDimCnt = UBound(arrLeft)

    ReDim Preserve arrRight(intMaxDimCnt)

    For intCnt = intMinDimCnt + 1 To intMaxDimCnt
      arrRight(intCnt) = "0"
    Next

  ElseIf UBound(arrRight) > UBound(arrLeft) Then
    intMinDimCnt = UBound(arrLeft)
    intMaxDimCnt = UBound(arrRight)

    ReDim Preserve arrLeft(intMaxDimCnt)

    For intCnt = intMinDimCnt + 1 To intMaxDimCnt
      arrLeft(intCnt) = "0"
    Next
  End If

  For intCnt = 0 To UBound(arrLeft)
    If Fix(arrLeft(intCnt)) > Fix(arrRight(intCnt)) Then
      intResult = 1
      Exit For

    ElseIf Fix(arrLeft(intCnt)) < Fix(arrRight(intCnt)) Then
      intResult = -1
      Exit For
    End If
  Next

  CompareVersion = intResult
End Function


'===============================================================================
' Get image architecture
'===============================================================================

Function GetImageArchitecture(ByRef strImagePath)
  Dim objFSO, objShell, strPath, intResult

  Set objFSO   = CreateObject("Scripting.FileSystemObject")
  Set objShell = CreateObject("WScript.Shell")

  strPath = objFSO.GetAbsolutePathName(strImagePath)

  intResult = objShell.Run(Quote(strWAEPath) & " " & _
                           "k@GetBinaryTypeW " & Quote(strPath) & " $b:4 , " & _
                           "k@InterlockedExchange $$:3 0", _
                           0, _
                           True)

  Select Case intResult
    Case 0    GetImageArchitecture = "x86"
    Case 6    GetImageArchitecture = "x64"
    Case Else GetImageArchitecture = ""
  End Select
End Function


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
    Call ForceDirectories(strDstFolder)
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
  Dim objShell

  Set objShell = CreateObject("WScript.Shell")

  objShell.Run "cmd.exe " & _
                 "/c ""md " & _
                 Quote(strPath) & """", _
               0, _
               True
End Function


'===============================================================================
' Delete whole directory tree
'===============================================================================

Sub DeleteDirTree(ByRef strRootDir)
  Dim objShell

  Set objShell = CreateObject("WScript.Shell")

  objShell.Run "cmd.exe " & _
                 "/c ""rd /s /q " & _
                 Quote(strRootDir) & """", _
               0, _
               True
End Sub


'===============================================================================
' Get OS architecture specific "Program Files" folder path
'===============================================================================

Function GetProgramFilesPath32Bit
  Dim objWMI

  Set objWMI = GetObject("winmgmts:root\cimv2:Win32_OperatingSystem=@")

  If objWMI.OSArchitecture = "64-Bit" Then
    GetProgramFilesPath32Bit = PROGRAM_FILES_PATH_X86_X64
  Else
    GetProgramFilesPath32Bit = PROGRAM_FILES_PATH_X86_X86
  End If
End Function


'===============================================================================
' Surround a string with double quotes
'===============================================================================

Function Quote(ByRef strString)
  Quote = """" & strString & """"
End Function


'===============================================================================
' Cleanup and terminate script
'===============================================================================

Sub CleanupAndQuit(intResultCode)
  If objFSO.FileExists(strPluginListDownloadPath) Then
    Call objFSO.DeleteFile(strPluginListDownloadPath)
  End If

  WScript.Echo
  WScript.Quit intResultCode
End Sub


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
