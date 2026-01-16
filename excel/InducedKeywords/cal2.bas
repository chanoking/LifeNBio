Attribute VB_Name = "cal2"
Sub cal2()
    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim wsSummary As Worksheet
    Dim dict As Object
    Dim weekObj As Object, monthObj As Object
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long, add As Long
    Dim key, view
    Dim aWeekAgo As Date, aMonthAgo As Date
    Dim d As Date
    
    Set ws = ThisWorkbook.Sheets("Rate of change")
    Set wsSource = ThisWorkbook.Sheets("rank_raw")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' -------------------------------
    ' rank_raw → Dictionary
    ' -------------------------------
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        key = wsSource.Cells(r, "A").Value
        view = wsSource.Cells(r, "B").Value
        If Not dict.Exists(key) Then
            dict.add key, view
        End If
    Next r
    
    ' -------------------------------
    ' Rate of change 초기 복사
    ' -------------------------------
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("B2:B" & lastRow).Value = ws.Range("C2:C" & lastRow).Value
    
    ' -------------------------------
    ' 현재값 / 변화율 계산
    ' -------------------------------
    For r = 8 To lastRow
        key = ws.Cells(r, "A").Value
        
        If isInclude(r, 1) Or key = "" Then
            ' skip
        Else
            If dict.Exists(key) Then
                ws.Cells(r, "C").Value = dict(key)
            Else
                ws.Cells(r, "C").Value = 10
            End If
            
            ws.Cells(r, "D").Value = _
                Round(((ws.Cells(r, "C").Value - ws.Cells(r, "B").Value) / ws.Cells(r, "C").Value) * 100, 2) & "%"
            
            With ws.Cells(r, "D")
                If ws.Cells(r, "D").Value > 0 Then
                    .Interior.Color = RGB(255, 235, 238)
                    .Font.Bold = True
                ElseIf ws.Cells(r, "D").Value < 0 Then
                    .Interior.Color = RGB(227, 242, 253)
                    .Font.Bold = True
                Else
                    .Interior.Color = RGB(245, 245, 245)
                    .Font.Bold = False
                End If
            End With
        End If
    Next r
    
    ' -------------------------------
    ' 값 복사
    ' -------------------------------
    ws.Range("G2:G" & lastRow).Value = ws.Range("C2:C" & lastRow).Value
    ws.Range("K2:K" & lastRow).Value = ws.Range("C2:C" & lastRow).Value
    
    ' -------------------------------
    ' Summary 처리
    ' -------------------------------
    Set weekObj = CreateObject("Scripting.Dictionary")
    Set monthObj = CreateObject("Scripting.Dictionary")
    
    aWeekAgo = Date - 7
    aMonthAgo = Date - 28
    
    lastCol = wsSummary.Cells(1, wsSummary.Columns.Count).End(xlToLeft).Column
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, "A").End(xlUp).Row
    
    For c = 2 To lastCol
        d = wsSummary.Cells(1, c).Value
        
        If d = aWeekAgo Then
            For r = 5 To lastRow
                key = wsSummary.Cells(r, "A").Value
                view = wsSummary.Cells(r, c).Value
                If Not isInclude(r, 2) Then
                    weekObj(key) = view
                End If
            Next r
        End If
        
        If d = aMonthAgo Then
            For r = 5 To lastRow
                key = wsSummary.Cells(r, "A").Value
                view = wsSummary.Cells(r, c).Value
                If Not isInclude(r, 2) Then
                    monthObj(key) = view
                End If
            Next r
            Exit For
        End If
    Next c
    
    ' -------------------------------
    ' 결과 반영
    ' -------------------------------
    For r = 8 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        key = ws.Cells(r, "A").Value
        
        If isInclude(r, 1) Or key = "" Then
            ' skip
        Else
            ws.Cells(r, "F").Value = weekObj(key)
            ws.Cells(r, "J").Value = monthObj(key)
            
            ws.Cells(r, "H").Value = _
                Round(((ws.Cells(r, "G").Value - ws.Cells(r, "F").Value) / ws.Cells(r, "G").Value) * 100, 2) & "%"
            
            With ws.Cells(r, "H")
                If ws.Cells(r, "H").Value > 0 Then
                    .Interior.Color = RGB(255, 235, 238)
                    .Font.Bold = True
                ElseIf ws.Cells(r, "H").Value < 0 Then
                    .Interior.Color = RGB(227, 242, 253)
                    .Font.Bold = True
                Else
                    .Interior.Color = RGB(245, 245, 245)
                    .Font.Bold = False
                End If
            End With
            
            ws.Cells(r, "L").Value = _
                Round(((ws.Cells(r, "K").Value - ws.Cells(r, "J").Value) / ws.Cells(r, "K").Value) * 100, 2) & "%"
                
            With ws.Cells(r, "L")
                If ws.Cells(r, "L").Value > 0 Then
                    .Interior.Color = RGB(255, 235, 238)
                    .Font.Bold = True
                ElseIf ws.Cells(r, "L").Value < 0 Then
                    .Interior.Color = RGB(227, 242, 253)
                    .Font.Bold = True
                Else
                    .Interior.Color = RGB(245, 245, 245)
                    .Font.Bold = False
                End If
            End With
        End If
    Next r
    
    '-----------------------------------------------------------------------------------------------------------------
    
    Dim itemRows, ranges, applyColumns, values
    Dim i As Long, j As Long
    
    itemRows = Array(7, 11, 15, 20, 27, 33, 39, 47, 52, 56, 61, 66, 71, 76, 79, 82, 86, 92, 96, 103, 116, 130, 134, 141)
    ranges = Array("A8:A9", "A12:A13", "A16:A18", "A21:A25", "A28:A31", _
                   "A34:A37", "A40:A45", "A48:A50", "A53:A54", "A57:A59", _
                    "A62:A64", "A67:A69", "A72:A74", "A77:A77", "A80:A80", _
                    "A83:A84", "A87:A90", "A93:A94", "A97:A101", "A104:A111", _
                    "A117:A128", "A131:A132", "A135:A136", "A142:A146")
    applyColumns = Array("B", "C", "F", "G", "J", "K")
    
    For i = LBound(applyColumns) To UBound(applyColumns)
        add = 0
            For j = LBound(itemRows) To UBound(itemRows)
                If j = 13 Or j = 14 Then
                    ws.Cells(itemRows(j), applyColumns(i)).Value = ws.Range(Replace(ranges(j), "A", applyColumns(i))).Value
                Else
                    values = ws.Range(Replace(ranges(j), "A", applyColumns(i))).Value
                    ws.Cells(itemRows(j), applyColumns(i)).Value = sum(values)
                End If
                
                add = add + ws.Cells(itemRows(j), applyColumns(i)).Value
                
                If itemRows(j) = 103 Then
                    ws.Cells(5, applyColumns(i)).Value = add
                    add = 0
                ElseIf itemRows(j) = 134 Then
                    ws.Cells(114, applyColumns(i)).Value = add
                    add = 0
                ElseIf itemRows(j) = 141 Then
                    ws.Cells(139, applyColumns(i)).Value = add
                    add = 0
                    ws.Cells(2, applyColumns(i)).Value = ws.Cells(139, applyColumns(i)).Value _
                                                        + ws.Cells(114, applyColumns(i)).Value _
                                                        + ws.Cells(5, applyColumns(i)).Value
                End If
                
            Next j
    Next i
    
    For i = LBound(itemRows) To UBound(itemRows)
        ws.Cells(itemRows(i), "D") = Round((ws.Cells(itemRows(i), "C") - ws.Cells(itemRows(i), "B")) _
                                                / ws.Cells(itemRows(i), "C"), 2) & "%"
        ws.Cells(itemRows(i), "H") = Round((ws.Cells(itemRows(i), "G") - ws.Cells(itemRows(i), "F")) _
                                                / ws.Cells(itemRows(i), "G"), 2) & "%"
        ws.Cells(itemRows(i), "L") = Round((ws.Cells(itemRows(i), "K") - ws.Cells(itemRows(i), "J")) _
                                                / ws.Cells(itemRows(i), "K"), 2) & "%"
                                                                                    
        With ws.Cells(itemRows(i), "D")
            If ws.Cells(itemRows(i), "D").Value > 0 Then
                .Interior.Color = RGB(255, 235, 238)
                .Font.Bold = True
            ElseIf ws.Cells(itemRows(i), "D").Value < 0 Then
                .Interior.Color = RGB(227, 242, 253)
                .Font.Bold = True
            Else
                .Interior.Color = RGB(245, 245, 245)
                .Font.Bold = False
            End If
        End With
        
        With ws.Cells(itemRows(i), "H")
            If ws.Cells(itemRows(i), "H").Value > 0 Then
                .Interior.Color = RGB(255, 235, 238)
                .Font.Bold = True
            ElseIf ws.Cells(itemRows(i), "H").Value < 0 Then
                .Interior.Color = RGB(227, 242, 253)
                .Font.Bold = True
            Else
                .Interior.Color = RGB(245, 245, 245)
                .Font.Bold = False
            End If
        End With
        
        With ws.Cells(itemRows(i), "L")
            If ws.Cells(itemRows(i), "L").Value > 0 Then
                .Interior.Color = RGB(255, 235, 238)
                .Font.Bold = True
            ElseIf ws.Cells(itemRows(i), "L").Value < 0 Then
                .Interior.Color = RGB(227, 242, 253)
                .Font.Bold = True
            Else
                .Interior.Color = RGB(245, 245, 245)
                .Font.Bold = False
            End If
        End With
    Next i
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    
    Dim branRows As Variant
    Dim rowNum As Long
    Dim a, b, t
    brandrows = Array(5, 114, 139, 2)
    combine = Array("CBD", "GFH", "KJL")
    For i = LBound(brandrows) To UBound(brandrows)
        rowNum = brandrows(i)
        For j = LBound(combine) To UBound(combine)
            a = Left(combine(j), 1)
            b = Mid(combine(j), 2, 1)
            t = Right(combine(j), 1)
            ws.Cells(rowNum, t).Value = Round(((ws.Cells(rowNum, a).Value - ws.Cells(rowNum, b).Value) _
                                                    / ws.Cells(rowNum, a).Value), 2) & "％"
                                                    
            With ws.Cells(rowNum, t)
                If ws.Cells(rowNum, t).Value > 0 Then
                    .Interior.Color = RGB(255, 235, 238)
                    .Font.Bold = True
                ElseIf ws.Cells(rowNum, t).Value < 0 Then
                    .Interior.Color = RGB(227, 242, 253)
                    .Font.Bold = True
                Else
                    .Interior.Color = RGB(245, 245, 245)
                    .Font.Bold = False
                End If
            End With
        Next j
    Next i
    
    MsgBox "Carried out what you asked!", vbInformation
End Sub


Function isInclude(target As Long, version As Long) As Boolean
    Dim i As Long
    Dim itemsRows As Variant
    
    If version = 1 Then
        itemsRows = Array(11, 15, 20, 27, 33, 39, 47, 52, 56, 61, 66, 71, 76, 79, 82, _
                            86, 92, 96, 103, 114, 116, 130, 134, 139, 141)
        
        For i = LBound(itemsRows) To UBound(itemsRows)
            If target = itemsRows(i) Then
                isInclude = True
                Exit Function
            End If
        Next i
        isInclude = False
    Else
        itemsRows = Array(7, 10, 14, 20, 25, 30, 37, 41, 44, 48, 52, 56, _
                            60, 62, 64, 67, 72, 75, 81, 90, 91, 104, 107, 110, 111)
        
        For i = LBound(itemsRows) To UBound(itemsRows)
            If target = itemsRows(i) Then
                isInclude = True
                Exit Function
            End If
        Next i
        isInclude = False
    End If
End Function

Function sum(arr As Variant) As Long
    Dim i As Long, acc As Long

    For i = LBound(arr) To UBound(arr)
        acc = acc + arr(i, 1)
    Next i
    
    sum = acc
End Function
