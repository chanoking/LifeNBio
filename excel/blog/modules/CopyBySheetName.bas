Attribute VB_Name = "CopyBySheetName"
Sub CopyBySheetName()
    Dim wsCur As Worksheet, wsMain As Worksheet
    Dim targetDate As Date
    Dim lastRow As Long, lastRowB As Long
    
    Set wsCur = ActiveSheet
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    
    currentName = wsCur.name
    targetDate = ActiveCell.value
    
    lastRow = wsCur.Cells(wsCur.Rows.Count, "A").End(xlUp).Row
    lastRowB = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    
    With wsMain.Range("B1")
        .AutoFilter field:=2, Criteria1:="=" & targetDate
        .AutoFilter field:=18, Criteria1:=currentName
    End With
    
    wsMain.Range("B2:B" & lastRowB).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("A" & lastRow + 1).PasteSpecial xlPasteValues
    
    wsMain.Range("R2:R" & lastRowB).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("B" & lastRow + 1).PasteSpecial xlPasteValues
    
    wsMain.Range("H2:H" & lastRowB).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("C" & lastRow + 1).PasteSpecial xlPasteValues
    
    wsMain.Range("K2:N" & lastRowB).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("D" & lastRow + 1).PasteSpecial xlPasteValues

    Application.CutCopyMode = False
    
    wsMain.ShowAllData
    wsMain.AutoFilterMode = False
    
End Sub


