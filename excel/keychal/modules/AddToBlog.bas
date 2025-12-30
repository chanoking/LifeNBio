Attribute VB_Name = "AddToBlog"
Sub AddToBlog()
    Dim wsBlog As Worksheet, wsMain As Worksheet
    Dim lastRow As Long, startRow As Long
    Dim volume As Long
    Dim arr() As Variant
    Dim i As Long, idx As Long

    Set wsBlog = ActiveSheet
    Set wsMain = ThisWorkbook.Sheets("원고기입")

    startRow = ActiveCell.Row + 1
    lastRow = wsMain.Cells(wsMain.Rows.Count, "R").End(xlUp).Row

    ' A열
    wsBlog.Range("A" & startRow & ":A" & lastRow).value = _
        wsMain.Range("A" & startRow & ":A" & lastRow).value

    ' B~G ← C~H
    wsBlog.Range("B" & startRow & ":G" & lastRow).value = _
        wsMain.Range("C" & startRow & ":H" & lastRow).value

    ' 날짜 분해
    volume = lastRow - startRow + 1
    ReDim arr(1 To volume, 1 To 3)

    idx = 1
    For i = startRow To lastRow
        arr(idx, 1) = Right(Year(wsMain.Range("B" & i).value), 2)
        arr(idx, 2) = Month(wsMain.Range("B" & i).value)
        arr(idx, 3) = Day(wsMain.Range("B" & i).value)
        idx = idx + 1
    Next i

    wsBlog.Range("H" & startRow & ":J" & lastRow).value = arr

    ' K~O ← J~N
    wsBlog.Range("K" & startRow & ":O" & lastRow).value = _
        wsMain.Range("J" & startRow & ":N" & lastRow).value

    ' P ← R
    wsBlog.Range("P" & startRow & ":P" & lastRow).value = _
        wsMain.Range("R" & startRow & ":R" & lastRow).value
End Sub


