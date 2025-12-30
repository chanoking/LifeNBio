Attribute VB_Name = "FetchURLTitle"
Sub fetchURLTitleQuoteDate()
Attribute fetchURLTitleQuoteDate.VB_ProcData.VB_Invoke_Func = "Q\n14"
    Dim wsTarget As Worksheet, wsMain As Worksheet
    Dim targetSheetName As String
    Dim targetDate As Date
    Dim r As Long, lastRow As Long
    
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    
    ' Get target sheet name and row info
    targetSheetName = ActiveCell.value
    r = ActiveCell.Row
    targetDate = wsMain.Cells(r, "B").value
    'Debug.Print row
    'Debug.Print targetDate
    
    Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
    
    ' Find last row in target sheet
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    
    Debug.Print lastRow
    
    ' Apply filter and copy visible rows
    With wsTarget.Range("A1:N" & lastRow)
        ' Clear any existing filters
        If wsTarget.AutoFilterMode Then .AutoFilter
        
        ' Filter by date in column A
        .AutoFilter field:=1, Criteria1:="=" & targetDate
        
        ' Copy visible rows excluding header
        On Error Resume Next
        .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Columns("K:N").Copy
        On Error GoTo 0
    End With
    
    ' Paste values in column S of current row
    wsMain.Range("S" & r).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Clear filter in target sheet
    If wsTarget.FilterMode Then wsTarget.ShowAllData
    wsTarget.AutoFilterMode = False
End Sub

