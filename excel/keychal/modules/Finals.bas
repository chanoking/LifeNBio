Attribute VB_Name = "Finals"
Sub Finals()
Attribute Finals.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim wsSource As Worksheet, wsThis As Worksheet
    Dim sh As Common
    Dim lastRow As Long
    
    Set wsSource = ThisWorkbook.Sheets("정산관리")
    Set wsThis = ActiveSheet
    Set sh = New Common
    
    sh.init "정산관리"
    
    Dim targetInfl As String
    targetInfl = ActiveCell.value
    
    With wsSource.Range("A1")
        .AutoFilter field:=5, Criteria1:=targetInfl
        .AutoFilter field:=3, Criteria1:="메인"
        .AutoFilter field:=4, Criteria1:="월보장"
    End With
    
    lastRow = wsThis.Cells(wsThis.Rows.Count, "E").End(xlUp).row
    
    wsThis.Range("D2:" & "AS" & lastRow).ClearContents
    
    wsSource.Range("E2:E" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("D2").PasteSpecial xlPasteValues
    
    wsSource.Range("G2:G" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("E2").PasteSpecial xlPasteValues
    
    wsSource.Range("A2:A" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("F2").PasteSpecial xlPasteValues
    
    wsSource.Range("L2:M" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("G2").PasteSpecial xlPasteValues

    wsSource.Range("J2:J" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("I2").PasteSpecial xlPasteValues
    
    wsSource.Range("P2:T" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("J2").PasteSpecial xlPasteValues
    
    wsSource.Range("V2:AZ" & sh.lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("O2").PasteSpecial xlPasteValues
    
    wsSource.Range("V1:AZ1").SpecialCells(xlCellTypeVisible).Copy
    wsThis.Range("O1").PasteSpecial xlPasteValues
    
    wsSource.ShowAllData
    wsSource.AutoFilterMode = False
    
End Sub
