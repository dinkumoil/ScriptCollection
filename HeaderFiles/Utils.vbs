'===============================================================================
' Sortiert ein Array zwischen den angegebenen Indizes mit Quicksort
'===============================================================================

Sub Quicksort(ByRef arrValues(), ByVal intMin, ByVal intMax)
  Dim varMediumValue, intHigh, intLow, intIdx

  'Wenn der Bereich nur ein Element lang ist, ist er sortiert
  If intMin >= intMax Then Exit Sub

  'Index zur Teilung des Bereichs zufällig bestimmen
  'und das Element an diesem Index als Pivot-Element auswählen
  intIdx = intMin + Int(Rnd(intMax - intMin + 1))
  varMediumValue = arrValues(intIdx)

  'Element am Teilungsindex an den Anfang des Bereichs verschieben
  arrValues(intIdx) = arrValues(intMin)

  'Bereich in zwei Teilbereiche aufteilen
  intLow  = intMin
  intHigh = intMax

  'Wiederholen bis der Bereich sortiert ist
  Do
    'Vom Bereichsende her nach einem Element < Pivot-Element suchen
    Do While arrValues(intHigh) >= varMediumValue
      intHigh = intHigh - 1
      If intHigh <= intLow Then Exit Do
    Loop

    If intHigh <= intLow Then
      'Der Bereich ist sortiert
      arrValues(intLow) = varMediumValue

      Exit Do
    End If

    'Erstes und letztes Element des Bereichs vertauschen
    arrValues(intLow) = arrValues(intHigh)

    'Vom Bereichsanfang her nach einem Element >= Pivot-Element suchen
    intLow = intLow + 1

    Do While arrValues(intLow) < varMediumValue
      intLow = intLow + 1
      If intLow >= intHigh Then Exit Do
    Loop

    If intLow >= intHigh Then
      'Der Bereich ist sortiert
      intLow = intHigh
      arrValues(intHigh) = varMediumValue

      Exit Do
    End If

    'Letztes und erstes Element des Bereichs vertauschen
    arrValues(intHigh) = arrValues(intLow)
  Loop

  'Rekursiver Funktionsaufruf mit geänderten Bereichsgrenzen
  Call Quicksort(arrValues, intMin, intLow - 1)
  Call Quicksort(arrValues, intLow + 1, intMax)
End Sub



'===============================================================================
' Entfernt aufeinanderfolgende doppelte Elemente aus dem Array arrValues (im
' angegebenen Bereich) und schreibt das Ergebnis in das Array arrNewValues
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
