'///////////////////////////////////////////////////////////////////////////////
'
' Header file for storing data that is about to be displayed in a tableized form
'
' Author: Andreas Heim
' Date: 18.02.2024
'
' Required header files (to be included before): None
'
'///////////////////////////////////////////////////////////////////////////////



Const ALIGNMENT_LEFT  = 0
Const ALIGNMENT_RIGHT = 1



'===============================================================================
' Data model of a table column
'===============================================================================

Class clsTableColumn
  Private intAlignment
  Private intMaxWidth
  Private arrRows


  Private Sub Class_Initialize()
    intAlignment = ALIGNMENT_LEFT
    intMaxWidth  = 0
    arrRows      = Array()
  End Sub


  Public Sub AddRow(ByRef strRow)
    Row(UBound(arrRows) + 1) = strRow
  End Sub


  Public Function DeleteRow(intIdx)
    Dim intRowIdx

    If Len(arrRows(intIdx)) <> intMaxWidth Then
      For intRowIdx = intIdx To UBound(arrRows) - 1
        arrRows(intRowIdx) = arrRows(intRowIdx + 1)
      Next
    Else
      intMaxWidth = 0

      For intRowIdx = intIdx To UBound(arrRows) - 1
        If intRowIdx < intIdx Then
          Rows(intRowIdx) = arrRows(intRowIdx)
        Else
          Rows(intRowIdx) = arrRows(intRowIdx + 1)
        End If
      Next
    End If

    ReDim Preserve arrRows(UBound(arrRows) - 1)
  End Function


  Public Function GetRows
    GetRows = arrRows
  End Function


  Public Sub SetRows(ByRef arrARows)
    arrRows = arrARows
  End Sub


  Public Property Get RowCount
    RowCount = UBound(arrRows) + 1
  End Property


  Public Property Let RowCount(intValue)
    Dim intIdx

    ReDim Preserve arrRows(intValue - 1)

    For intIdx = 0 To UBound(arrRows)
      If IsNull(arrRows(intIdx)) Or IsEmpty(arrRows(intIdx)) Then
        arrRows(intIdx) = ""
      End If
    Next
  End Property


  Public Property Get MaxWidth
    MaxWidth = intMaxWidth
  End Property


  Public Property Get Alignment
    Alignment = intAlignment
  End Property


  Public Property Let Alignment(intValue)
    intAlignment = intValue
  End Property


  Public Default Property Get Row(intIdx)
    Row = arrRows(intIdx)
  End Property


  Public Property Let Row(intIdx, ByRef strARow)
    If intIdx > UBound(arrRows) Then
      RowCount = intIdx + 1
    End If

    arrRows(intIdx) = strARow

    If Len(strARow) > intMaxWidth Then
      intMaxWidth = Len(strARow)
    End If
  End Property
End Class



'===============================================================================
' Data model of an entire table consisting of columns
'===============================================================================

Class clsTable
  Private arrColumns


  Private Sub Class_Initialize()
    arrColumns = Array()
  End Sub


  Public Sub Print(ByRef strColDelim)
    Dim intColIdx, intRowIdx, intMaxRows
    Dim strLine, strPadding

    intMaxRows = 0

    For intColIdx = 0 To UBound(arrColumns)
      If arrColumns(intColIdx).RowCount > intMaxRows Then
        intMaxRows = arrColumns(intColIdx).RowCount
      End If
    Next

    For intRowIdx = 0 To intMaxRows - 1
      strLine = ""

      For intColIdx = 0 To UBound(arrColumns)
        strPadding = String(arrColumns(intColIdx).MaxWidth, " ")

        If intRowIdx < arrColumns(intColIdx).RowCount Then
          If strLine <> "" Then
            strLine = strLine & strColDelim
          End If

          If arrColumns(intColIdx).Alignment = ALIGNMENT_RIGHT Then
            strLine = strLine & Right(strPadding & arrColumns(intColIdx)(intRowIdx), arrColumns(intColIdx).MaxWidth)
          Else
            strLine = strLine & Left(arrColumns(intColIdx)(intRowIdx) & strPadding, arrColumns(intColIdx).MaxWidth)
          End If
        Else
          strLine = strLine & strPadding & strColDelim
        End If
      Next

      WScript.Echo strLine
    Next
  End Sub


  Public Function AddColumn(ByRef objColumn)
    Column(UBound(arrColumns) + 1) = objColumn
  End Function


  Public Function DeleteColumn(intIdx)
    Dim intColIdx

    For intColIdx = intIdx To UBound(arrColumns) - 1
      Set arrColumns(intColIdx) = arrColumns(intColIdx + 1)
    Next

    ReDim Preserve arrColumns(UBound(arrColumns) - 1)
  End Function


  Public Property Get ColCount
    ColCount = UBound(arrColumns) + 1
  End Property


  Public Property Let ColCount(intValue)
    Dim intIdx

    ReDim Preserve arrColumns(intValue - 1)

    For intIdx = 0 To UBound(arrColumns)
      If Not IsObject(arrColumns(intIdx)) Then
        Set arrColumns(intIdx) = New clsTableColumn
      End If
    Next
  End Property


  Public Property Get Columns
    Columns = arrColumns
  End Property


  Public Property Set Columns(ByRef arrAColumns)
    arrColumns = arrAColumns
  End Property


  Public Default Property Get Column(intIdx)
    Set Column = arrColumns(intIdx)
  End Property


  Public Property Set Column(intIdx, ByRef objColumn)
    If intIdx > UBound(arrColumns) Then
      ColCount = intIdx + 1
    End If

    arrColumns(intIdx) = objColumn
  End Property
End Class
