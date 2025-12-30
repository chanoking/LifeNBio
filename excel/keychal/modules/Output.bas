Attribute VB_Name = "Output"
Sub Output()
    Dim wsCur As Worksheet, wsMain As Worksheet
    Dim lastRow As Long, firstDate As Long
    
    Set wsCur = ThisWorkbook.Sheets("송출내역")
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).row
    
    firstDate = DateSerial(2025, 11, 0) + 1
    
    On Error Resume Next
    wsMain.Range("A1").AutoFilter field:=2, Criteria1:=">=" & firstDate
    On Error GoTo 0
    
    wsMain.Range("B2:D" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("B2").PasteSpecial xlPasteValues
    
    wsMain.Range("G2:H" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("F2").PasteSpecial xlPasteValues
    
    wsMain.Range("N2:N" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("H2").PasteSpecial xlPasteValues
    
    wsMain.Range("Q2:Q" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("I2").PasteSpecial xlPasteValues
    
    wsMain.Range("R2:R" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("J2").PasteSpecial xlPasteValues
    
    wsMain.ShowAllData
    wsMain.AutoFilterMode = False
    
    lastRow = wsCur.Cells(wsCur.Rows.Count, "C").End(xlUp).row
    
    wsCur.Range("A2:A" & lastRow).value = "진찬호"
    wsCur.Range("E2:E" & lastRow).value = "광고"
End Sub
