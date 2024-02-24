'///////////////////////////////////////////////////////////////////////////////
'
' Header file for utility functions
'
' Author: Andreas Heim
'
' Required header files (to be included before): None
'
'///////////////////////////////////////////////////////////////////////////////



'===============================================================================
' Character encoding names for conversion function ConvertEncoding
'===============================================================================

Const Enc_ASMO_708     = "ASMO-708"
Const Enc_BIG5         = "big5"
Const Enc_DOS_437      = "cp437"
Const Enc_DOS_708      = "ASMO-708"
Const Enc_DOS_720      = "DOS-720"
Const Enc_DOS_737      = "ibm737"
Const Enc_DOS_775      = "ibm775"
Const Enc_DOS_850      = "cp850"
Const Enc_DOS_852      = "cp852"
Const Enc_DOS_855      = "cp855"
Const Enc_DOS_857      = "cp857"
Const Enc_DOS_858      = "cp858"
Const Enc_DOS_860      = "cp860"
Const Enc_DOS_861      = "cp861"
Const Enc_DOS_862      = "DOS-862"
Const Enc_DOS_863      = "cp863"
Const Enc_DOS_864      = "cp864"
Const Enc_DOS_865      = "cp865"
Const Enc_DOS_866      = "cp866"
Const Enc_DOS_869      = "cp869"
Const Enc_DOS_874      = "DOS-874"
Const Enc_EUC_JP       = "euc-jp"
Const Enc_EUC_KR       = "euc-kr"
Const Enc_GB2312       = "GB2312"
Const Enc_HZ_GB2312    = "HZ-GB-2312"
Const Enc_GB18030      = "GB18030"
Const Enc_IBM367       = "ibm367"
Const Enc_IBM775       = "ibm775"
Const Enc_IBM819       = "ibm819"
Const Enc_IBM852       = "ibm852"
Const Enc_IBM866       = "ibm866"
Const Enc_ISO_2022_JP  = "iso-2022-jp"
Const Enc_ISO_2022_KR  = "iso-2022-kr"
Const Enc_ISO_8859_1   = "iso-8859-1"
Const Enc_ISO_8859_2   = "iso-8859-2"
Const Enc_ISO_8859_3   = "iso-8859-3"
Const Enc_ISO_8859_4   = "iso-8859-4"
Const Enc_ISO_8859_5   = "iso-8859-5"
Const Enc_ISO_8859_6   = "iso-8859-6"
Const Enc_ISO_8859_7   = "iso-8859-7"
Const Enc_ISO_8859_8   = "iso-8859-8"
Const Enc_ISO_8859_8_i = "iso-8859-8-i"
Const Enc_ISO_8859_9   = "iso-8859-9"
Const Enc_ISO_8859_15  = "iso-8859-15"
Const Enc_KOI8_R       = "koi8-r"
Const Enc_KOI8_U       = "koi8-u"
Const Enc_KOI8_RU      = "koi8-ru"
Const Enc_KSC_5601     = "ks_c_5601-1987"
Const Enc_Shift_JIS    = "shift-jis"
Const Enc_US_ASCII     = "us-ascii"
Const Enc_UTF_7        = "utf-7"
Const Enc_UTF_8        = "utf-8"
Const Enc_UTF_16       = "unicode"
Const Enc_UTF_16LE     = "unicode"
Const Enc_UTF_16BE     = "unicodeFFFE"
Const Enc_Windows_874  = "windows-874"
Const Enc_Windows_1250 = "windows-1250"
Const Enc_Windows_1251 = "windows-1251"
Const Enc_Windows_1252 = "windows-1252"
Const Enc_Windows_1253 = "windows-1253"
Const Enc_Windows_1254 = "windows-1254"
Const Enc_Windows_1255 = "windows-1255"
Const Enc_Windows_1256 = "windows-1256"
Const Enc_Windows_1257 = "windows-1257"
Const Enc_Windows_1258 = "windows-1258"



'===============================================================================
' Create a nested directory structure on local drives and network shares
'===============================================================================

Function ForceDirectories(ByRef strPath)
  Dim objFSO, arrAbsPath, strAbsPath, strPartPath, intCnt

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  'Retrieve absolute path of input path
  strAbsPath = objFSO.GetAbsolutePathName(strPath)
  arrAbsPath = Array()

  'Split path into its parts and store them in an array.
  'Last part of path is stored in lowest order array element, the path's root
  'is stored in highest order array element
  Do
    'Enlarge array for path parts
    ReDim Preserve arrAbsPath(UBound(arrAbsPath) + 1)

    'Check if last part of the path is a drive's root dir or if the remaining
    'path is the path of a network share
    If objFSO.GetFileName(strAbsPath)  = ""         Or _
       objFSO.GetDriveName(strAbsPath) = strAbsPath Then
      'Store path of drive's root dir or path of network share and exit loop
      arrAbsPath(UBound(arrAbsPath)) = strAbsPath
      Exit Do
    Else
      'Store directory name
      arrAbsPath(UBound(arrAbsPath)) = objFSO.GetFileName(strAbsPath)
    End If

    'Discard last part of path
    strAbsPath = objFSO.GetParentFolderName(strAbsPath)
  Loop

  'Init with path's root (root dir of drive or path of network share)
  strPartPath = arrAbsPath(UBound(arrAbsPath))

  'Return failure if the path's root doesn't exist
  If Not objFSO.DriveExists(strPartPath) Then
    ForceDirectories = False
  Else
    'Create directories level by level
    For intCnt = UBound(arrAbsPath) - 1 To 0 Step -1
      strPartPath = objFSO.BuildPath(strPartPath, arrAbsPath(intCnt))

      If Not objFSO.FolderExists(strPartPath) Then
        objFSO.CreateFolder(strPartPath)
      End If
    Next

    'Return success
    ForceDirectories = True
  End If
End Function



'===============================================================================
' Generates a unique filename
' Parameter strNamePattern must contain a template for the filename including
' a place holder for a random number
'===============================================================================

Dim bolPRNGInitialized

Function GetUniqueFileName(ByRef strNamePattern)
  If Not bolPRNGInitialized Then
    Randomize
    bolPRNGInitialized = True
  End If

  GetUniqueFileName = FormatString(strNamePattern, Array(Int(Abs((1 + Rnd) * 1000000))))
End Function



'===============================================================================
' Sort an array between the provided indices using Quicksort
'===============================================================================

Sub Quicksort(ByRef arrValues(), ByVal intMin, ByVal intMax)
  Dim varMediumValue, intHigh, intLow, intIdx

  'Break if length of range to sort is only one element
  If intMin >= intMax Then Exit Sub

  'Get random index for range's division
  'and select element at that index as pivot
  intIdx = intMin + Int(Rnd(intMax - intMin + 1))
  varMediumValue = arrValues(intIdx)

  'Replace pivot element with element from lower range border
  arrValues(intIdx) = arrValues(intMin)

  'Set subranges border indices
  intLow  = intMin
  intHigh = intMax

  'Repeat until the range is sorted
  Do
    'Search for elements < pivot, starting at upper subrange's end
    Do While arrValues(intHigh) >= varMediumValue
      intHigh = intHigh - 1
      If intHigh <= intLow Then Exit Do
    Loop

    If intHigh <= intLow Then
      'Range is sorted
      arrValues(intLow) = varMediumValue

      Exit Do
    End If

    'Replace first element of subrange with its last element
    arrValues(intLow) = arrValues(intHigh)

    'Search for elements >= pivot, starting at lower subrange's beginning
    intLow = intLow + 1

    Do While arrValues(intLow) < varMediumValue
      intLow = intLow + 1
      If intLow >= intHigh Then Exit Do
    Loop

    If intLow >= intHigh Then
      'Range is sorted
      intLow = intHigh
      arrValues(intHigh) = varMediumValue

      Exit Do
    End If

    'Replace last element of subrange with its first element
    arrValues(intHigh) = arrValues(intLow)
  Loop

  'Call function recursive with changed range borders
  Call Quicksort(arrValues, intMin, intLow - 1)
  Call Quicksort(arrValues, intLow + 1, intMax)
End Sub



'===============================================================================
' Convert character encoding of input string
'===============================================================================

Function ConvertEncoding(ByRef strInput, ByRef strEncFrom, ByRef strEncTo)
  Dim objInStream, objOutStream

  Set objInStream  = CreateObject("ADODB.Stream")
  Set objOutStream = CreateObject("ADODB.Stream")

  With objOutStream
    .Mode    = 3  'adModeReadWrite
    .Type    = 2  'adTypeText
    .Charset = strEncTo
    .Open
  End With

  With objInStream
    .Mode = 3  'adModeReadWrite
    .Type = 2  'adTypeText

    Select Case UCase(strEncFrom)
      Case UCase(Enc_UTF_7)
        .Charset = "utf-7"
      Case UCase(Enc_UTF_8)
        .Charset = "utf-8"
      Case UCase(Enc_UTF_16), _
           UCase(Enc_UTF_16LE)
        .Charset = "unicode"
      Case UCase(Enc_UTF_16BE)
        .Charset = "unicodeFFFE"
      Case Else
        .Charset = "Windows-1252"
    End Select

    .Open
    .WriteText strInput

    .Position = 0
    .Charset  = strEncFrom
    .CopyTo objOutStream
    .Close
  End With

  With objOutStream
    .Position = 0
    ConvertEncoding = .ReadText(-1)  'adReadAll
    .Close
  End With
End Function



'===============================================================================
' Retrieve a subrange of an array
' Start and end index can be negativ, in this case they are relative to the
' array's opposite end, thus the order of elements in the array can be inverted
'===============================================================================

Function Slice(ByRef arrInput, ByVal intStart, ByVal intEnd)
  Dim intIdx, intStep, arrOutput

  If IsArray(arrInput) Then
    If intStart < 0 Then
      intStart = intStart + UBound(arrInput) + 1
    End If

    If intEnd < 0 Then
      intEnd = intEnd + UBound(arrInput) + 1
    End If

    ReDim arrOutput(Abs(intStart - intEnd))

    If intStart > intEnd Then
      intStep = -1
    Else
      intStep = 1
    End If

    For intIdx = intStart To intEnd Step intStep
      If IsObject(arrInput(intIdx)) Then
        Set arrOutput(Abs(intIdx - intStart)) = arrInput(intIdx)
      Else
        arrOutput(Abs(intIdx - intStart)) = arrInput(intIdx)
      End If
    Next

    Slice = arrOutput
  Else
    Slice = Null
  End If
End Function



'===============================================================================
' Merges the contents of two arrays by adding the elements of the first array at
' the end of the second array
'===============================================================================

Function Merge(ByRef arrLeft, ByRef arrRight)
  Dim intIdx, arrOutput

  If IsArray(arrLeft) And IsArray(arrRight) Then
    ReDim arrOutput(UBound(arrLeft) + UBound(arrRight) + 1)

    For intIdx = 0 To UBound(arrLeft)
      If IsObject(arrLeft(intIdx)) Then
        Set arrOutput(intIdx) = arrLeft(intIdx)
      Else
        arrOutput(intIdx) = arrLeft(intIdx)
      End If
    Next

    For intIdx = 0 To UBound(arrRight)
      If IsObject(arrRight(intIdx)) Then
        Set arrOutput(intIdx + UBound(arrLeft) + 1) = arrRight(intIdx)
      Else
        arrOutput(intIdx + UBound(arrLeft) + 1) = arrRight(intIdx)
      End If
    Next

    Merge = arrOutput
  Else
    Merge = Null
  End If
End Function



'===============================================================================
' Creates a dictionary from two arrays by using the elements of the first array
' as keys and the elements of the second array as values
'===============================================================================

Function Combine(ByRef arrKeys, ByRef arrValues)
  Dim intMaxIdx, intIdx, dicOutput

  If IsArray(arrKeys) And IsArray(arrValues) Then
    intMaxIdx = UBound(arrKeys)

    If intMaxIdx > UBound(arrValues) Then
      intMaxIdx = UBound(arrValues)
    End If

    Set dicOutput = CreateObject("Scripting.Dictionary")

    For intIdx = 0 To intMaxIdx
      If Not dicOutput.Exists(arrKeys(intIdx)) Then
        Call dicOutput.Add(arrKeys(intIdx), arrValues(intIdx))
      End If
    Next

    Set Combine = dicOutput
  Else
    Set Combine = Nothing
  End If
End Function



'===============================================================================
' Remove consecutive duplicate elements from an input array (within the provided
' range) and write the remaining elements to an output array
'===============================================================================

Sub RemoveDuplicates(ByRef arrValues(), ByVal intMin, ByVal intMax, ByRef arrNewValues())
  Dim intValuesIdx, intNewValuesIdx
  ReDim arrNewValues(0)

  intNewValuesIdx = 0

  For intValuesIdx = intMin To intMax-1
    If Not arrValues(intValuesIdx) = arrValues(intValuesIdx+1) Then
      arrNewValues(intNewValuesIdx) = arrValues(intValuesIdx)
      ReDim Preserve arrNewValues(intNewValuesIdx + 1)
      intNewValuesIdx = intNewValuesIdx + 1
    End If
  Next

  arrNewValues(intNewValuesIdx) = arrValues(intMax)
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
      strString = Replace(strString, strVar, arrItems(intCnt), 1, -1, vbTextCompare)
      intStart  = intPos + Len(arrItems(intCnt))
    End If
  Next

  FormatString = strString
End Function



'===============================================================================
' Checks if a value is part of a set of values
'===============================================================================

Function IsOf(ByRef varValue, ByRef arrValues)
  Dim intCnt

  IsOf = False

  For intCnt = 0 To UBound(arrValues)
    If CStr(arrValues(intCnt)) = CStr(varValue) Then
      IsOf = True
      Exit For
    End If
  Next
End Function



'===============================================================================
' Checks if a path contains only a bare file name
'===============================================================================

Function IsBareFileName(ByRef strPath)
  Dim objFSO

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  IsBareFileName = (objFSO.GetParentFolderName(strPath) = "" And _
                    UBound(Filter(Array("\", "/"), Left(strPath, 1))) < 0)
End Function



'===============================================================================
' Convert DateTime value to time stamp in ISO-8601 format
'===============================================================================

Function DateTimeToISO8601(ByRef datDateTime)
  Dim objXmlDoc, objNode, objDateTime
  Dim strSign, strTZBias

  Set objXmlDoc              = CreateObject("Microsoft.XMLDOM")
  objXmlDoc.async            = False
  objXmlDoc.validateOnParse  = False
  objXmlDoc.resolveExternals = False

  Set objNode                = objXmlDoc.createElement("TimeStamp")
  objXmlDoc.appendChild objNode

  objNode.dataType           = "datetime"
  objNode.nodeTypedValue     = datDateTime
  objNode.dataType           = ""

  Set objDateTime            = CreateObject("WbemScripting.SWbemDateTime")
  objDateTime.SetVarDate datDateTime, True

  If objDateTime.UTC >= 0 Then
    strSign = "+"
  Else
    strSign = "-"
  End If

  strTZBias = Right("0" & Abs(Int(objDateTime.UTC / 60)), 2) & ":" &  Right("0" & Abs(objDateTime.UTC mod 60), 2)

  DateTimeToISO8601 = objNode.text & strSign & strTZBias
End function



'===============================================================================
' Convert time stamp in ISO-8601 format to DateTime value
'===============================================================================

Function ISO8601ToDateTime(ByVal strDateTime, bolAsUTC)
  Dim objXmlDoc, objNode, strUTC, strUTCSign, strTZBias

  Set objXmlDoc              = CreateObject("Microsoft.XMLDOM")
  objXmlDoc.async            = False
  objXmlDoc.validateOnParse  = False
  objXmlDoc.resolveExternals = False

  Set objNode                = objXmlDoc.createElement("TimeStamp")
  objXmlDoc.appendChild objNode

  strUTC                     = Right(strDateTime, 5)
  strUTCSign                 = Mid(strDateTime, Len(strDateTime) - 6, 1)
  intTZBias                  = CInt(Left(strUTC, 2)) * 60 + CInt(Right(strUTC, 2))
  strDateTime                = Left(strDateTime, Len(strDateTime) - 6)

  objNode.text               = strDateTime
  objNode.dataType           = "datetime"

  If Not bolAsUTC Then
    ISO8601ToDateTime = objNode.nodeTypedValue
  Else
    ISO8601ToDateTime = DateAdd("n", CInt(strUTCSign & intTZBias) * -1, objNode.nodeTypedValue)
  End If
End Function



'===============================================================================
' Decode a string containing HTML entities
'===============================================================================

Function HtmlDecode(ByRef strInString)
  Dim objHtmlDoc

  Set objHtmlDoc = CreateObject("htmlfile")

  objHtmlDoc.Open
  objHtmlDoc.Write strInString
  objHtmlDoc.Close

  HtmlDecode = objHtmlDoc.body.innerText
End Function



'===============================================================================
' Surround a string with double quotes
'===============================================================================

Function Quote(ByRef strString)
  Quote = """" & strString & """"
End Function



'===============================================================================
' Surround a string with a pair of characters
'===============================================================================

Function Enclose(ByRef strString, ByRef strLeftString, ByRef strRightStr)
  Enclose = strLeftString & strString & strRightStr
End Function



'===============================================================================
' Escape special chars of string for use with WMI
'===============================================================================

Function EscapeForWMI(ByRef strAString)
  EscapeForWMI = Replace(strAString, "\", "\\")
End Function
