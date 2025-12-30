Attribute VB_Name = "FindUnexposedKeywords"
Sub FindUnexposedKeywords()
Attribute FindUnexposedKeywords.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim wsCurrent As Worksheet, wsManagement As Worksheet
    Dim lastRow As Long
    
    Set wsCurrent = ThisWorkbook.Sheets("변환용")
    Set wsManagement = ThisWorkbook.Sheets("정산관리")
    
    lastRow = wsManagement.Cells(wsManagement.Rows.Count, "A").End(xlUp).row
    
    On Error Resume Next
    With wsManagement.Range("A1")
        .AutoFilter field:=2, Criteria1:="메인"
        .AutoFilter field:=21, Criteria1:=0
        .AutoFilter field:=3, Criteria1:="월보장"
    End With
    On Error GoTo 0
    
    wsManagement.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCurrent.Range("B2").PasteSpecial xlPasteValues
    wsManagement.Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCurrent.Range("A2").PasteSpecial xlPasteValues
    
    
    Application.CutCopyMode = False
    wsManagement.ShowAllData
    wsManagement.AutoFilterMode = False
End Sub

