Const FVICompanyName      = 0
Const FVIFileDescription  = 1
Const FVIComments         = 2
Const FVIProductName      = 3
Const FVIInternalName     = 4
Const FVIOriginalFilename = 5
Const FVIFileVersion      = 6
Const FVIProductVersion   = 7
Const FVILegalCopyright   = 8
Const FVILegalTrademarks  = 9
Const FVIPrivateBuild     = 10
Const FVISpecialBuild     = 11



Class clsFileVersionInfo
  Private objFSO
  Private objShell
  Private objNameSpace
  Private objFolderItem
  Private strFilePath
  Private arrVersionInfoTags


  '----------------------------------------------------------------------------
  'Constructor
  '----------------------------------------------------------------------------
  Private Sub Class_Initialize()
    Set objFSO   = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("Shell.Application")

    arrVersionInfoTags = Array ( _
      Array("CompanyName"     , "Company name"     ), _
      Array("FileDescription" , "Description"      ), _
      Array("Comments"        , "Comments"         ), _
      Array("ProductName"     , "Product name"     ), _
      Array("InternalName"    , "Internal name"    ), _
      Array("OriginalFilename", "Original filename"), _
      Array("FileVersion"     , "File version"     ), _
      Array("ProductVersion"  , "Product version"  ), _
      Array("Copyright"       , "Copyright"        ), _
      Array("LegalTrademarks" , "Trademarks"       ), _
      Array("PrivateBuild"    , "Private build"    ), _
      Array("SpecialBuild"    , "Special build"    )  _
    )

    Clear()
  End Sub


  '----------------------------------------------------------------------------
  'Destructor
  '----------------------------------------------------------------------------
  Private Sub Class_Terminate()
    Clear()

    Set objFSO   = Nothing
    Set objShell = Nothing
  End Sub


  '----------------------------------------------------------------------------
  'Interne Variablen zurücksetzen
  '----------------------------------------------------------------------------
  Private Sub Clear
    strFilePath       = ""
    Set objNameSpace  = Nothing
    Set objFolderItem = Nothing
  End Sub


  '----------------------------------------------------------------------------
  'Liest/Setzt den Pfad der Datei zum Extrahieren von Versionsinformationen
  '----------------------------------------------------------------------------
  Public Property Let FilePath(ByRef strValue)
    strFilePath = objFSO.GetAbsolutePathName(strValue)

    If StrComp(objFSO.GetExtensionName(strFilePath), "exe", vbTextCompare) = 0 Or _
       StrComp(objFSO.GetExtensionName(strFilePath), "dll", vbTextCompare) = 0 Then
      Set objNameSpace  = objShell.Namespace(objFSO.GetParentFolderName(strFilePath))
      Set objFolderItem = objNameSpace.ParseName(objFSO.GetFileName(strFilePath))
    Else
      Clear()
    End If
  End Property


  Public Property Get FilePath
    FilePath = strFilePath
  End Property


  '----------------------------------------------------------------------------
  'Liest bestimmte Versionsinformationen aus einer Datei
  '----------------------------------------------------------------------------
  Public Property Get FileVersionInfoTag(intValue)
    FileVersionInfoTag = ""

    If strFilePath = ""                         Then Exit Property
    If intValue    < 0                          Then Exit Property
    if intValue    > UBound(arrVersionInfoTags) Then Exit Property

    FileVersionInfoTag = CStr(objFolderItem.ExtendedProperty(arrVersionInfoTags(intValue)(0)))
  End Property


  '----------------------------------------------------------------------------
  'Liest die Namen bestimmter Versionsinformationen
  '----------------------------------------------------------------------------
  Public Property Get FileVersionInfoTagName(intValue)
    FileVersionInfoTagName = ""

    If strFilePath = ""                         Then Exit Property
    If intValue    < 0                          Then Exit Property
    if intValue    > UBound(arrVersionInfoTags) Then Exit Property

    FileVersionInfoTagName = arrVersionInfoTags(intValue)(0)
  End Property


  '----------------------------------------------------------------------------
  'Liest die Bezeichnung bestimmter Versionsinformationen
  '----------------------------------------------------------------------------
  Public Property Get FileVersionInfoTagFriendlyName(intValue)
    FileVersionInfoTagFriendlyName = ""

    If strFilePath = ""                         Then Exit Property
    If intValue    < 0                          Then Exit Property
    if intValue    > UBound(arrVersionInfoTags) Then Exit Property

    FileVersionInfoTagFriendlyName = arrVersionInfoTags(intValue)(1)
  End Property
End Class
