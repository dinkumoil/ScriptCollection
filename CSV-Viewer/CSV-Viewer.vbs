Option Explicit


Dim objFSO, colArgs
Dim strInFileName


Set colArgs = WScript.Arguments

If colArgs.Count > 0 Then
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  strInFileName = objFSO.GetAbsolutePathName(colArgs(0))

  If Not OpenCSVFile(strInFileName, objFSO.GetBaseName(strInFileName)) Then
    MsgBox "Failed to load file" & vbCrLf & _
           vbCrLf & _
           strInFileName & vbCrLf & _
           vbCrLf & _
           Err.Description, _
           vbCritical+vbOKOnly, _
           objFSO.GetBaseName(WScript.ScriptName)
  End If
End If



Function OpenCSVFile(strFileName, strSheetName)
  On Error Resume Next
  
  Dim objExcel, objActiveWorkBook, objActiveSheet
  Dim arrColumnFormats(100), intCnt, bolResult
  
  bolResult = False
  
  'Limt length of worksheet title to 31 characters
  'and remove problematic ones
  strSheetName = Left(strSheetName, 31)
  strSheetName = Replace(strSheetName, "[", "(")
  strSheetName = Replace(strSheetName, "]", ")")

  'Create array with format codes for the first 100 columns
  'The format is set to Text
  For intCnt = 0 to UBound(arrColumnFormats)
    arrColumnFormats(intCnt) = 2  'xlTextFormat
  Next
  
  'Start Excel
  Set objExcel           = CreateObject("Excel.Application")
  objExcel.Visible       = False                         'Set Excel invisible
  objExcel.DisplayAlerts = False                         'Supress save confirmation when closing Excel

  'Add workbook and retrieve active worksheet
  Set objActiveWorkBook = objExcel.WorkBooks.Add(-4167)  'xlWBATWorksheet
  Set objActiveSheet    = objActiveWorkBook.WorkSheets(1)

  'Read CSV file
  With objActiveSheet.QueryTables.Add("TEXT;" & strFileName, objActiveSheet.Range("A1"))
    .Name                         = "Vx80 Report"
    .FieldNames                   = False                'First line doesn't contain column headers
    .RowNumbers                   = False                'First column doesn't contain line numbers
    .FillAdjacentFormulas         = False                'Don't update formulars
    .PreserveFormatting           = False                'Don't adopt cell formats of first 5 lines to new lines
    .RefreshStyle                 = 0                    'xlOverwriteCells, Overwrite cell contents
    .AdjustColumnWidth            = True                 'Fit column widths to content
    .RefreshPeriod                = 0                    'Turn off automatic data update
    .TextFilePlatform             = 1252                 'Code page of input file (65001 = UTF-8)
    .TextFileStartRow             = 1                    'Data starts in line 1
    .TextFileParseType            = 1                    'xlDelimited, Columns are separated by delimiters
    .TextFileTabDelimiter         = False                'Column delimiter is TAB
    .TextFileSemicolonDelimiter   = True                 'Column delimiter is semicolon
    .TextFileCommaDelimiter       = True                 'Column delimiter is comma
    .TextFileSpaceDelimiter       = False                'Column delimiter is space
    .TextFileOtherDelimiter       = ""                   'No alternative delimiter
    .TextFileConsecutiveDelimiter = False                'Don't interpret consecutive delimiters as one delimiter
    .TextFileTextQualifier        = 1                    'xlTextQualifierDoubleQuote, Double quotes are delimiters for text content
    .TextFileColumnDataTypes      = arrColumnFormats     'Column format codes
'    .TextFileDecimalSeparator     = "."                  'Decimal separator
'    .TextFileThousandsSeparator   = ","                  'Thousands separator
    .TextFileTrailingMinusNumbers = False                'Negative numbers are text too

    'Load file and wait until it is loaded completely
    If Err.Number = 0 Then
      bolResult                   = .Refresh(False)
    End If
  End With
  
  'Check if file has been loaded successfully
  If Err.Number = 0 And bolResult Then
    objActiveSheet.Name           = strSheetName         'Set worksheet title
    objExcel.Visible              = True                 'Set Excel visible
  Else
    objExcel.Quit                                        'Terminate Excel in case of failure
    objExcel                      = Nothing              'and free ActiveX object
  End If

  OpenCSVFile = bolResult  
End Function
