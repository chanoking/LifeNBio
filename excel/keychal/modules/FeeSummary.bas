Attribute VB_Name = "FeeSummary"
Sub FeeSummaryA()
    Dim wsCur As Worksheet, wsFee As Worksheet
    Dim lastRow As Long, r As Long, keyCnt As Long
    Dim sheetName As String
    
    Set wsCur = ActiveSheet
    Set wsFee = ThisWorkbook.Sheets("Old_정산관리")
    
    lastRow = wsFee.Cells(wsFee.Rows.Count, "A").End(xlUp).row
    keyCnt = 0
    sheetName = wsCur.name
    
    With wsFee.Range("A1")
        .AutoFilter field:=6, Criteria1:=sheetName
        .AutoFilter field:=19, Criteria1:=">0"
    End With
    
    wsCur.Range("A2:V" & lastRow).ClearContents
    
    wsFee.Range("S2:S" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("M2").PasteSpecial xlPasteValues
    
    wsFee.Range("G2:H" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("P2").PasteSpecial xlPasteValues
    
    wsFee.Range("L2:M" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("T2").PasteSpecial xlPasteValues
    
    wsFee.Range("T2:T" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("S2").PasteSpecial xlPasteValues
    
    wsFee.Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("V2").PasteSpecial xlPasteValues
    
    wsFee.Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsCur.Range("K2").PasteSpecial xlPasteValues
    
    wsFee.ShowAllData
    wsFee.AutoFilterMode = False
    
    lastRow = wsCur.Cells(wsCur.Rows.Count, "K").End(xlUp).row
    
    Dim fillValues(1 To 1, 1 To 9) As Variant
    Dim lastDate As Date

    lastDate = DateSerial(2025, 12, 0)
    
    fillValues(1, 1) = "라이프앤바이오"
    fillValues(1, 2) = "3.판관비"
    fillValues(1, 3) = "2.광고선전비"
    fillValues(1, 4) = "1.바이럴마케팅"
    fillValues(1, 6) = "바이럴_키챌월보장"
    fillValues(1, 7) = lastDate
    fillValues(1, 8) = Format(lastDate, "m") & "월 " & Format(lastDate, "d") & "일"
    fillValues(1, 9) = "키워드 챌린지 월보장"
    
    wsCur.Range("A2:I" & lastRow).value = fillValues
    Dim name As String
    Dim cell As Range
    
    For r = 2 To lastRow
        name = wsCur.Cells(r, "K").value
        If name = "모모둥이" Then
            wsCur.Cells(r, "K").value = "(주)모모컴퍼니"
            GoTo nextiteration
        End If
        If name = "민들레" Then
            wsCur.Cells(r, "K").value = "(주)민들레컴퍼니"
            GoTo nextiteration
        End If
        If name = "셀럽주부" Then
            wsCur.Cells(r, "K").value = "(주)에벤에셀컴퍼니"
            GoTo nextiteration
        End If
        If name = "푸들ol" Then
            wsCur.Cells(r, "K").value = "(주)엠케이푸"
            GoTo nextiteration
        End If
        If name = "갬성주부" Then
            wsCur.Cells(r, "K").value = "갬성주부(김숙진)"
            GoTo nextiteration
        End If
nextiteration:
    Next r
    
        
End Sub

