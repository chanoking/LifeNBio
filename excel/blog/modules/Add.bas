Attribute VB_Name = "Add"
Sub Add()
Attribute Add.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim ws As Worksheet, wsBlog As Worksheet, wsView As Worksheet
    Dim lastRow As Long
    Dim sh As New common
    
    Set ws = ThisWorkbook.Sheets("원고기입")
    Set wsBlog = ThisWorkbook.Sheets("블로그순위")
    Set wsView = ThisWorkbook.Sheets("조회수")
    
    sh.init "원고기입"
    lastRow = sh.lastRow
    
    Dim values As Variant
    
    ' ------- U Column ? multiply by 1.1 -------
    values = ws.Range("U2:U" & lastRow).value
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If IsNumeric(values(i, 1)) And values(i, 1) > 0 Then
            ws.Cells(i + 1, "W").value = values(i, 1) * 1.1
        End If
    Next i
    
    ' ------- B Column ? week calculation -------
    values = ws.Range("B2:B" & lastRow).value
    
    Dim d As Date, firstDate As Date, firstMonDate As Date, week As Long
    
    For i = LBound(values) To UBound(values)
        If IsDate(values(i, 1)) Then
            d = values(i, 1)
            firstDate = DateSerial(year(d), month(d), 1)
            firstMonDate = firstDate + (8 - Weekday(firstDate, vbMonday)) Mod 7
            week = Int((d - firstMonDate) / 7) + 1
            
            ws.Cells(i + 1, "X").value = Right(year(d), 2) & "년 " & month(d) & "월 " & week & "주"
        End If
    Next i
    
    ' ------- Copy S:V to Y:AB -------
    ws.Range("Y2:AB" & lastRow).value = wsBlog.Range("S2:V" & lastRow).value
    
    ' ------- Latest week -------
    Dim latest As Variant
    latest = ws.Cells(lastRow, "X").value
    
    Dim r As Long, keyword As String
    Dim found As Range, rowIdx As Long
    
    For r = lastRow To 2 Step -1
        If ws.Cells(r, "X").value = latest Then
            keyword = Replace(ws.Cells(r, "N").value, " ", "")
            
            Set found = wsView.Range("A:A").Find(What:=keyword, _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                MatchCase:=False)
            
            If Not found Is Nothing Then
                rowIdx = found.Row
                ws.Range("AC" & r & ":AD" & r).value = wsView.Range("C" & rowIdx & ":D" & rowIdx).value
            End If
        Else
            Exit For
        End If
    Next r
    
    MsgBox "Completo!"
End Sub


