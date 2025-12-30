Attribute VB_Name = "copySheet"
Sub copySheet()
    Dim wsSrc As Worksheet
    Dim wsTgt As Worksheet
    Dim url As String
    Dim lastRow As Long
    Dim startingRow As Long
    Dim i As Long
    Dim foundCell As Range
    
    Set wsSrc = ActiveSheet
    Set wsTgt = ThisWorkbook.Sheets("원고기입")
    
    ' URL from last used row in column Q
    url = wsSrc.Cells(wsSrc.Cells(wsSrc.Rows.Count, "P").End(xlUp).Row, "Q").value
    
    ' Find URL in target sheet column S
    Set foundCell = wsTgt.Columns("S").Find(What:=url, LookIn:=xlValues, LookAt:=xlPart)
    If foundCell Is Nothing Then
        MsgBox "URL not found in target sheet!", vbExclamation
        Exit Sub
    End If
    
    startingRow = foundCell.Row + 1
    lastRow = wsTgt.Cells(wsTgt.Rows.Count, "S").End(xlUp).Row
    
    ' Copy values directly
    wsSrc.Range("A" & startingRow & ":A" & lastRow).value = wsTgt.Range("A" & startingRow & ":A" & lastRow).value
    wsSrc.Range("B" & startingRow & ":G" & lastRow).value = wsTgt.Range("C" & startingRow & ":H" & lastRow).value
    
    ' Convert dates
    For i = startingRow To lastRow
        If IsDate(wsTgt.Cells(i, "B").value) Then
            wsSrc.Cells(i, "H").value = Right(year(wsTgt.Cells(i, "B").value), 2)
            wsSrc.Cells(i, "I").value = Format(month(wsTgt.Cells(i, "B").value), "00")
            wsSrc.Cells(i, "J").value = Format(Day(wsTgt.Cells(i, "B").value), "00")
        Else
            wsSrc.Cells(i, "H").value = ""
            wsSrc.Cells(i, "I").value = ""
            wsSrc.Cells(i, "J").value = ""
        End If
    Next i
    
    wsSrc.Range("K" & startingRow & ":N" & lastRow).value = wsTgt.Range("J" & startingRow & ":M" & lastRow).value
    wsSrc.Range("O" & startingRow & ":O" & lastRow).value = wsTgt.Range("R" & startingRow & ":R" & lastRow).value
    wsSrc.Range("P" & startingRow & ":P" & lastRow).value = wsTgt.Range("N" & startingRow & ":N" & lastRow).value
    wsSrc.Range("Q" & startingRow & ":R" & lastRow).value = wsTgt.Range("S" & startingRow & ":T" & lastRow).value
    
End Sub

