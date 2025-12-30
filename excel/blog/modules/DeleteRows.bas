Attribute VB_Name = "DeleteRows"
Sub DeleteRowsAndRecalculate()
Attribute DeleteRowsAndRecalculate.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim wsSource As Worksheet
    Dim wsRank As Worksheet
    Dim wsTarget As Worksheet
    Dim targetValue As Variant
    Dim targetSheetName As String
    Dim targetRow As Long
    
    On Error GoTo ErrHandler
    
    '--- Source sheet setup ---
    Set wsSource = ActiveSheet
    targetValue = ActiveCell.value
    targetSheetName = wsSource.Cells(ActiveCell.Row, "R").value
    targetRow = ActiveCell.Row
    
    '--- Validation ---
    If targetValue = "" Then
        MsgBox "No target value selected.", vbExclamation
        Exit Sub
    End If
    
    If targetSheetName = "" Then
        MsgBox "No sheet name found in column R.", vbExclamation
        Exit Sub
    End If
    
    '--- Target sheets ---
    Set wsRank = ThisWorkbook.Sheets("블로그순위")
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    
    '--- Delete in both referencing sheets ---
    Call DeleteRowsInSheet(wsRank, "R", targetValue)
    Call DeleteRowsInSheet(wsTarget, "M", targetValue)
    
    '--- Delete the row in the original sheet (after the two deletions) ---
    Call DeleteRowsInSheet(wsSource, "T", targetValue)
    
    '--- Recalculate only the referencing sheets ---
    wsRank.Calculate
    wsTarget.Calculate
    
    MsgBox "Rows with '" & targetValue & "' deleted and recalculated in '" & wsRank.name & "' and '" & wsTarget.name & "'.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub DeleteRowsInSheet(wsTarget As Worksheet, colLetter As String, targetValue As Variant)
    Dim targetCol As Long
    Dim lastRow As Long
    Dim rng As Range
    
    targetCol = wsTarget.Range(colLetter & "1").Column
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, targetCol).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub  ' nothing to delete
    
    Set rng = wsTarget.Range("A1", wsTarget.Cells(lastRow, wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column))
    
    If wsTarget.AutoFilterMode Then wsTarget.AutoFilterMode = False
    rng.AutoFilter field:=targetCol, Criteria1:=targetValue
    
    On Error Resume Next
    rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0
    
    wsTarget.AutoFilterMode = False
End Sub

