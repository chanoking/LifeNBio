Attribute VB_Name = "deleteRelatedSheets"
Sub deleteOtherSheet()
Attribute deleteOtherSheet.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim wsMain As Worksheet, wsOther As Worksheet, wsAnother As Worksheet
    Dim targetUrl As String
    
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    Set wsOther = ThisWorkbook.Sheets("블로그순위")
    Set wsAnother = ThisWorkbook.Sheets("붙이기용")
    
    targetUrls = Selection
    
    For Each cell In targetUrls
        Call DeleteRowsInSheet(wsAnother, "U", cell)
        Call DeleteRowsInSheet(wsOther, "P", cell)
        Call DeleteRowsInSheet(wsMain, "R", cell)
    Next cell

    
    MsgBox "Rows with '" & targetValue & "' deleted in '" & wsMain.name & "','" & wsOther.name & "', and '" & wsAnother.name & "'.", vbInformation
End Sub

Private Sub DeleteRowsInSheet(wsTarget As Worksheet, colLetter As String, targetValue As Variant)
    Dim targetCol As Long
    Dim lastRow As Long
    Dim rng As Range
    
    targetCol = wsTarget.Range(colLetter & 1).Column
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, colLetter).End(xlUp).row
    
    If lastRow < 2 Then Exit Sub
    
    Set rng = wsTarget.Range("A1", wsTarget.Cells(lastRow, wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column))
    
    If wsTarget.AutoFilterMode Then wsTarget.AutoFilterMode = False
    rng.AutoFilter field:=targetCol, Criteria1:=targetValue
    
    On Error Resume Next
    rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0
    
    wsTarget.AutoFilterMode = False
End Sub
