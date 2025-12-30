Attribute VB_Name = "Weekly"
Sub Weekly()
Attribute Weekly.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wsThis As Worksheet, wsOther As Worksheet
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim fDate As Date, lDate As Date
    
    Set wsThis = ThisWorkbook.Sheets("정산관리(보장주간)")
    Set wsOther = ThisWorkbook.Sheets("정산관리")
    
    lastRow = wsOther.Cells(wsOther.Rows.Count, "A").End(xlUp).row
    lastCol = wsOther.Cells(1, wsOther.Columns.Count).End(xlToLeft).Column
    fDate = Date - Weekday(Date, 2) - 6
    lDate = Date - Weekday(Date, 2)
    
    On Error Resume Next
    With wsOther.Range("A1")
        .AutoFilter field:=3, Criteria1:="메인"
        .AutoFilter field:=4, Criteria1:="월보장"
    End With
    On Error GoTo 0
    
    wsOther.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("A2").PasteSpecial xlPasteValues
    
    wsOther.Range("E2:N" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("B2").PasteSpecial xlPasteValues
    
    wsOther.Range("P2:Q" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("L2").PasteSpecial xlPasteValues

    Dim destCol As Long
    destCol = 17
    For c = wsOther.Cells(1, "V").Column To lastCol
        If wsOther.Cells(1, c).value >= fDate And wsOther.Cells(1, c).value <= lDate Then
            wsThis.Cells(1, destCol).value = wsOther.Cells(1, c).value
            wsOther.Range(wsOther.Cells(2, c), wsOther.Cells(lastRow, c)).SpecialCells(xlCellTypeVisible).Copy
            wsThis.Cells(2, destCol).PasteSpecial xlPasteValues
            destCol = destCol + 1
        End If
    Next c
    
    Application.CutCopyMode = False
    
    wsOther.ShowAllData
    wsOther.AutoFilterMode = False
    
    lastRow = wsThis.Cells(wsThis.Rows.Count, "A").End(xlUp).row
    Dim sum As Long
    For r = 2 To lastRow
        sum = 0
        For c = 19 To lastCol
            If wsThis.Cells(r, c) > 0 Then
                sum = sum + 1
            End If
        Next c
        wsThis.Cells(r, "N").value = sum
        wsThis.Cells(r, "O").value = wsThis.Cells(r, "M") * sum
        If wsThis.Cells(r, "C").value = "세금" Then
            wsThis.Cells(r, "P").value = wsThis.Cells(r, "O").value * 1.1
        Else
            wsThis.Cells(r, "P").value = wsThis.Cells(r, "O").value * 0.967
        End If
    Next r
    
    
End Sub
