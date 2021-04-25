'///////////////////////////////////////////////////////////////////////////////
'
' Header file for fault tolerant comparisons (phonetic comparison using Soundex
' or Kölner Phonetic (for comparing german phrases), Levenshtein distance and
' combined methods)
'
' Author: Andreas Heim
'
' Required header files (to be included before): None
'
'///////////////////////////////////////////////////////////////////////////////



Class clsSimilarity
  '-----------------------------------------------------------------------------
  ' Private Variablen
  '-----------------------------------------------------------------------------

  Private strAlphabet
  Private strUmlauts
  Private strDelims


  '-----------------------------------------------------------------------------
  ' Private Methoden
  '-----------------------------------------------------------------------------

  Private Sub Class_Initialize()
    strUmlauts  = "ÄÖÜß"
    strDelims   = " " & vbTab
    strAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & strUmlauts & strDelims
  End Sub


  Private Function IsInSet(ByRef strChar, ByRef strSet)
    If strChar <> "" Then
      IsInSet = (InStr(1, strSet, strChar, vbBinaryCompare) > 0)
    Else
      IsInSet = False
    End If
  End Function


  Private Function IsDelimiter(ByRef strChar)
    IsDelimiter = IsInSet(strChar, strDelims)
  End Function


  Private Function ConvertToUpperCase(ByRef strChar)
    ConvertToUpperCase = UCase(strChar)
  End Function


  Private Function FilterNonAlphabetMembers(ByRef strChar)
    If IsInSet(strChar, strAlphabet) Then
      FilterNonAlphabetMembers = strChar
    Else
      FilterNonAlphabetMembers = ""
    End If
  End Function


  Private Function ReplaceUmlaut(ByRef strChar)
    Select Case strChar
      Case "Ä"  ReplaceUmlaut = "A"
      Case "Ö"  ReplaceUmlaut = "O"
      Case "Ü"  ReplaceUmlaut = "U"
      Case "ß"  ReplaceUmlaut = "S"
      Case Else ReplaceUmlaut = strChar
    End Select
  End Function


  Private Function AddToColognePhoneticResult(ByRef strResult, ByRef strChar)
    AddToColognePhoneticResult = strResult

    If strChar = ""                         Then Exit Function
    If strChar = Right(strResult, 1)        Then Exit Function
    If strChar = "0" And Len(strResult) > 0 Then Exit Function

    AddToColognePhoneticResult = AddToColognePhoneticResult & strChar
  End Function


  Private Function Min(intNum1, intNum2)
    If intNum1 < intNum2 Then
      Min = intNum1
    ElseIf intNum2 < intNum1 Then
      Min = intNum2
    Else
      Min = intNum1
    End If
  End Function


  '-----------------------------------------------------------------------------
  ' Public Methoden
  '-----------------------------------------------------------------------------

  Public Function IsColognePhoneticEqual(ByRef strString1, ByRef strString2)
    IsColognePhoneticEqual = (ColognePhonetic(strString1) = ColognePhonetic(strString2))
  End Function


  Public Function ColognePhoneticLevenshteinDistance(ByRef strString1, ByRef strString2)
    ColognePhoneticLevenshteinDistance = LevenshteinDistance(ColognePhonetic(strString1), ColognePhonetic(strString2), True)
  End Function


  Public Function IsSoundexEqual(ByRef strString1, ByRef strString2)
    IsSoundexEqual = (Soundex(strString1) = Soundex(strString2))
  End Function


  Public Function SoundexLevenshteinDistance(ByRef strString1, ByRef strString2)
    SoundexLevenshteinDistance = LevenshteinDistance(Soundex(strString1), Soundex(strString2), True)
  End Function


  ' ~~~~~~~~~~~~~~~~
  ' Kölner Phonetik
  ' ~~~~~~~~~~~~~~~~
  Public Function ColognePhonetic(ByRef strString)
    Dim intLen, intIdx, bolIsOnset
    Dim strChar, strPrevChar, strCurChar, strNextChar
    Dim strResult

    intLen      = Len(strString)
    bolIsOnset  = True
    strPrevChar = ""
    strCurChar  = ""
    strNextChar = ""
    strResult   = ""

    For intIdx = 1 To intLen + 1
      If intIdx <= intLen Then
        strChar = Mid(strString, intIdx, 1)
        strChar = ConvertToUpperCase(strChar)
        strChar = FilterNonAlphabetMembers(strChar)
      Else
        strChar = ""
      End If

      If strChar <> "" Or intIdx > intLen Then
        strPrevChar = strCurChar
        strCurChar  = strNextChar
        strNextChar = strChar

        If strCurChar <> "" Then
          If IsDelimiter(strCurChar) Then
            bolIsOnset  = True
            strPrevChar = ""
          Else
            strCurChar = ReplaceUmlaut(strCurChar)

            Select Case strCurChar
              Case "A", "E", "I", "J", "O", "U", "Y"
                strResult = AddToColognePhoneticResult(strResult, "0")

              Case "B"
                strResult = AddToColognePhoneticResult(strResult, "1")

              Case "F", "V", "W"
                strResult = AddToColognePhoneticResult(strResult, "3")

              Case "G", "K", "Q"
                strResult = AddToColognePhoneticResult(strResult, "4")

              Case "L"
                strResult = AddToColognePhoneticResult(strResult, "5")

              Case "M", "N"
                strResult = AddToColognePhoneticResult(strResult, "6")

              Case "R"
                strResult = AddToColognePhoneticResult(strResult, "7")

              Case "S", "Z"
                strResult = AddToColognePhoneticResult(strResult, "8")

              Case "H"
                strResult = strResult

              Case "P"
                If IsInSet(strNextChar, "H") Then
                  strResult = AddToColognePhoneticResult(strResult, "3")
                Else
                  strResult = AddToColognePhoneticResult(strResult, "1")
                End If

              Case "D", "T"
                If IsInSet(strNextChar, "CSZ") Then
                  strResult = AddToColognePhoneticResult(strResult, "8")
                Else
                  strResult = AddToColognePhoneticResult(strResult, "2")
                End If

              Case "X"
                If IsInSet(strPrevChar, "CKQ") Then
                  strResult = AddToColognePhoneticResult(strResult, "8")
                Else
                  strResult = AddToColognePhoneticResult(strResult, "48")
                End If

              Case "C"
                If bolIsOnset Then
                  If IsInSet(strNextChar, "AHKLOQRUX") Then
                    strResult = AddToColognePhoneticResult(strResult, "4")
                  Else
                    strResult = AddToColognePhoneticResult(strResult, "8")
                  End If

                ElseIf IsInSet(strPrevChar, "SZ") Then
                  strResult = AddToColognePhoneticResult(strResult, "8")

                ElseIf IsInSet(strNextChar, "AHKOQUX") Then
                  strResult = AddToColognePhoneticResult(strResult, "4")

                Else
                  strResult = AddToColognePhoneticResult(strResult, "8")
                End If
            End Select

            bolIsOnset  = false
          End If
        End If
      End If
    Next

    ColognePhonetic = strResult
  End Function


  ' ~~~~~~~~
  ' Soundex
  ' ~~~~~~~~
  Public Function SoundEx(ByRef strString)
    Dim intLen, intIdx, strChar
    Dim bolIgnore, bolLookBack, intGrpLen
    Dim strCode, strResult

    intLen      = Len(strString)
    bolLookBack = True
    intGrpLen   = 0
    strResult   = ""

    For intIdx = 1 To intLen
      strChar = Mid(strString, intIdx, 1)
      strChar = ConvertToUpperCase(strChar)
      strChar = FilterNonAlphabetMembers(strChar)
      strCode = ""

      If strChar <> "" Then
        If IsDelimiter(strChar) Then
          If intGrpLen > 0 Then
            strResult   = strResult & String(4-intGrpLen, "0")
            bolLookBack = True
            intGrpLen   = 0
          End If

        ElseIf intGrpLen < 4 Then
          strChar   = ReplaceUmlaut(strChar)
          bolIgnore = False

          Select Case strChar
            Case "B", "F", "P", "V"
              strCode = "1"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "BFPV")
              End If

            Case "C", "G", "J", "K", "Q", "S", "X", "Z"
              strCode = "2"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "CGJKQSXZ")
              End If

            Case "D", "T"
              strCode = "3"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "DT")
              End If

            Case "L"
              strCode = "4"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "L")
              End If

            Case "M", "N"
              strCode = "5"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "MN")
              End If

            Case "R"
              strCode = "6"

              If bolLookBack And intGrpLen = 1 Then
                bolIgnore = IsInSet(Right(strResult, 1), "R")
              End If

            Case Else
              bolIgnore = (intGrpLen > 0)
          End Select

          If bolIgnore Then
            bolLookBack = False

          ElseIf intGrpLen = 0 Then
            strResult   = strResult & strChar
            intGrpLen   = intGrpLen + 1
            bolLookBack = True

          ElseIf Not bolLookBack Or Right(strResult, 1) <> strCode Then
            strResult   = strResult & strCode
            intGrpLen   = intGrpLen + 1
            bolLookBack = True
          End If
        End If
      End If
    Next

    If intGrpLen > 0 Then
      strResult = strResult & String(4-intGrpLen, "0")
    End If

    SoundEx = strResult
  End Function


  ' ~~~~~~~~~~~~~~~~~~~
  ' Levenshtein Distanz
  ' ~~~~~~~~~~~~~~~~~~~
  Public Function LevenshteinDistance(ByVal strString1, ByVal strString2, bolIgnoreCase)
    Dim intCost, intL1, intL2, intTMin
    Dim arrDistance, intRow, intCol

    intCost = 0
    intL1   = Len(strString1)
    intL2   = Len(strString2)

    If intL1 = 0 Then
      LevenshteinDistance = intL2
      Exit Function
    End If

    If intL2 = 0 Then
      LevenshteinDistance = intL1
      Exit Function
    End If

    If bolIgnoreCase Then
      strString1 = LCase(strString1)
      strString2 = LCase(strString2)
    End If

    ReDim arrDistance(intL1+1, intL2+1)

    For intRow = 0 To intL1+1
      arrDistance(intRow, 0) = intRow
    Next

    For intCol = 0 To intL2+1
      arrDistance(0, intCol) = intCol
    Next

    For intRow = 1 To intL1
      For intCol = 1 To intL2
        intCost = Abs(StrComp(Mid(strString2, intCol, 1), Mid(strString1, intRow, 1), vbBinaryCompare))
        intTMin = Min(arrDistance(intRow-1, intCol) + 1, arrDistance(intRow, intCol-1) + 1)
        arrDistance(intRow, intCol) = Min(intTMin, arrDistance(intRow-1, intCol-1) + intCost)
      Next
    Next

    LevenshteinDistance = arrDistance(intL1, intL2)
  End Function
End Class
