Attribute VB_Name = "Transition"
Sub UseForTransition()
Attribute UseForTransition.VB_ProcData.VB_Invoke_Func = "t\n14"
    Dim wsMain As Worksheet, wsManagement As Worksheet
    Dim lastRow As Long
    
    Set wsMain = ThisWorkbook.Sheets("변환용")
    Set wsManagement = ThisWorkbook.Sheets("정산관리")
    
    lastRow = wsManagement.Cells(wsManagement.Rows.Count, "A").End(xlUp).row
    
 
    lastRowB = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).row
    
    wsMain.Range("A2:D" & lastRowB).ClearContents
    
    wsManagement.Range("A2:A" & lastRow).Copy
    wsMain.Range("A2").PasteSpecial xlPasteValues
    
    wsManagement.Range("I2:I" & lastRow).Copy
    wsMain.Range("B2").PasteSpecial xlPasteValues
    
    wsManagement.Range("U2:U" & lastRow).Copy
    wsMain.Range("C2").PasteSpecial xlPasteValues
    
    wsManagement.Range("N2:N" & lastRow).Copy
    wsMain.Range("D2").PasteSpecial xlPasteValues
End Sub
