'///////////////////////////////////////////////////////////////////////////////
'
' Header file for file IO functions
'
' Author: Andreas Heim
'
' Required header files (to be included before): None
'
'///////////////////////////////////////////////////////////////////////////////



'Constants for file system routines
Const ForReading      = 1
Const ForWriting      = 2
Const ForAppending    = 8

Const AsSystemDefault = -2
Const AsUnicode       = -1
Const AsAnsi          = 0



'Routines for reading and writing UTF-8 files with BOM
Function ReadUTF8BOMFile(ByRef strFile)
  With CreateObject("ADODB.Stream")
    .Type    = 2  'adTypeText
    .Charset = "UTF-8"
    .Open
    .LoadFromFile(strFile)
    ReadUTF8BOMFile = .ReadText
    .Close
  End With
End Function


Sub WriteUTF8BOMFile(ByRef strFile, ByRef strContent)
  With CreateObject("ADODB.Stream")
    .Type    = 2  'adTypeText
    .Charset = "UTF-8"
    .Open
    .WriteText strContent
    .SaveToFile strFile, 2  'adSaveCreateOverWrite
    .Close
  End With
End Sub


'Routines for reading and writing UTF-8 files without BOM
Function ReadUTF8File(ByRef strFile)
  ReadUTF8File = ReadUTF8BOMFile(strFile)
End Function


Sub WriteUTF8File(ByRef strFile, ByRef strContent)
  Dim objUTF8Stream, objBinaryStream

  Set objUTF8Stream   = CreateObject("ADODB.Stream")
  Set objBinaryStream = CreateObject("ADODB.Stream")

  With objBinaryStream
    .Type = 1  'adTypeBinary
    .Mode = 3  'adModeReadWrite
    .Open
  End With

  With objUTF8Stream
    .Type    = 2  'adTypeText
    .Mode    = 3  'adModeReadWrite
    .Charset = "UTF-8"
    .Open
    .WriteText strContent

    .Position = 3  'Skip BOM
    .CopyTo objBinaryStream
    .Close
  End With

  With objBinaryStream
    .SaveToFile strFile, 2  'adSaveCreateOverWrite
    .Close
  End With
End Sub


'Routine to convert the EOL format of a UTF-8 file with BOM
'to the provided format
Sub ConvertUTF8BOMEOLFormat(ByRef strFile, ByRef strEOLFormat)
  Dim objInStream, objOutStream, intLineSeparator, strChar, strLine
  Dim bolReadEOL, strLastEOL

  Set objInStream  = CreateObject("ADODB.Stream")
  Set objOutStream = CreateObject("ADODB.Stream")

  Select Case strEOLFormat
    Case vbCr  : intLineSeparator = 13  'adCR'
    Case vbLf  : intLineSeparator = 10  'adLF'
    Case vbCrLf: intLineSeparator = -1  'adCRLF'
    Case Else  : intLineSeparator = -1  'adCRLF'
  End Select

  With objInStream
    .Type    = 2  'adTypeText
    .Mode    = 3  'adModeReadWrite
    .Charset = "UTF-8"
    .Open
    .LoadFromFile(strFile)
  End With

  With objOutStream
    .Type          = 2  'adTypeText
    .Mode          = 3  'adModeReadWrite
    .Charset       = "UTF-8"
    .LineSeparator = intLineSeparator
    .Open
  End With

  strLine    = ""
  strLastEOL = ""
  bolReadEOL = False

  Do While Not objInStream.EOS
    strChar = objInStream.ReadText(1)  'Read 1 character

    If strChar <> vbCr And strChar <> vbLf Then
      strLine    = strLine & strChar
      bolReadEOL = False
    ElseIf Not bolReadEOL Or strLastEOL = strChar Then
      objOutStream.WriteText strLine, 1  'adWriteLine
      strLine    = ""
      strLastEOL = strChar
      bolReadEOL = (strChar = vbCr)
    Else
      bolReadEOL = False
    End If
  Loop

  If strLine <> "" Then
    objOutStream.WriteText strLine, 1  'adWriteLine
  End If

  With objInStream
    .Close
  End With

  With objOutStream
    .SaveToFile strFile, 2  'adSaveCreateOverWrite
    .Close
  End With
End Sub


'Routine to convert the EOL format of a UTF-8 file without BOM
'to the provided format
Sub ConvertUTF8EOLFormat(ByRef strFile, ByRef strEOLFormat)
  Dim strContent

  Call ConvertUTF8BOMEOLFormat(strFile, strEOLFormat)

  strContent = ReadUTF8BOMFile(strFile)
  Call WriteUTF8File(strFile, strContent)
End Sub
