Attribute VB_Name = "WeeklyWith"
Sub WeeklyWith()
    Dim ws As Worksheet, wsMain As Worksheet
    Dim standardDate As Date
    Dim sheetName As String
    
    Set ws = ActiveSheet
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    sheetName = ActiveCell.value
    
    standardDate = Date - Weekday(Date, 2) + 1
    
    Debug.Print sheetName
    Debug.Print standardDate
    With wsMain.Range("A1")
        .AutoFilter field:=22, Criteria1:="=" & standardDate
        .AutoFilter field:=18, Criteria1:=sheetName
        .AutoFilter field:=21, Criteria1:=">0"
    End With
    
    Dim lastRow As Long
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    
    wsMain.Range("G2:H" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    ws.Range("P2").PasteSpecial xlPasteValues
    
    wsMain.Range("U2:U" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    ws.Range("M2").PasteSpecial xlPasteValues
    
    wsMain.Range("V2:V" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    ws.Range("H2").PasteSpecial xlPasteValues
    
    wsMain.Range("S2:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    ws.Range("T2").PasteSpecial xlPasteValues
    
    Dim arr As Variant
    Dim comb As Variant
    comb = Right(year(Date), 2) & "년 " & month(Date) & "월"
    
    arr = Array("라이프앤바이오", "3.판관비", "2.광고선전비", "1.바이럴마케팅", "바이럴_블로그건바이", comb, Date, "블로그 건바이", sheetName, "마케팅1팀")
    arrb = Array("A", "B", "C", "D", "F", "G", "H", "I", "K", "R")
    
    Dim i As Long, j As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row

    For i = LBound(arrb) To UBound(arrb)
        ws.Range(arrb(i) & "2:" & arrb(i) & lastRow).value = arr(i)
    Next i
    
    ws.Range("S2:S" & lastRow).value = ws.Evaluate("M2:M" & lastRow & "*1.1")
    
    wsMain.ShowAllData
    wsMain.AutoFilterMode = False
    
End Sub
