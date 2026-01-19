Sub cal2()

    Dim ws As Worksheet, wsSource As Worksheet, wsSummary As Worksheet
    Dim dict As Object, weekObj As Object, monthObj As Object
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long, add As Double
    Dim key, view
    Dim aWeekAgo As Date, aMonthAgo As Date
    Dim d As Date

    Set ws = ThisWorkbook.Sheets("Rate of change")
    Set wsSource = ThisWorkbook.Sheets("rank_raw")
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    Set dict = CreateObject("Scripting.Dictionary")

    '--------------------------------
    ' rank_raw → Dictionary
    '--------------------------------
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        key = wsSource.Cells(r, "A").Value
        view = wsSource.Cells(r, "B").Value
        If Not dict.exists(key) Then dict.add key, view
    Next r

    '--------------------------------
    ' 이전값 복사
    '--------------------------------
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("B2:B" & lastRow).Value = ws.Range("C2:C" & lastRow).Value

    '--------------------------------
    ' 현재값 + 변화율
    '--------------------------------
    For r = 8 To lastRow
        key = ws.Cells(r, "A").Value

        If Not isInclude(r, 1) And key <> "" Then
            If dict.exists(key) Then
                ws.Cells(r, "C").Value = dict(key)
            Else
                ws.Cells(r, "C").Value = 10
            End If

            SetRate ws.Cells(r, "D"), ws.Cells(r, "C"), ws.Cells(r, "B")
        End If
    Next r

    '--------------------------------
    ' 값 복사
    '--------------------------------
    ws.Range("G2:G" & lastRow).Value = ws.Range("C2:C" & lastRow).Value
    ws.Range("K2:K" & lastRow).Value = ws.Range("C2:C" & lastRow).Value

    '--------------------------------
    ' Summary 처리
    '--------------------------------
    Set weekObj = CreateObject("Scripting.Dictionary")
    Set monthObj = CreateObject("Scripting.Dictionary")

    aWeekAgo = Date - 7
    aMonthAgo = Date - 28

    lastCol = wsSummary.Cells(1, wsSummary.Columns.Count).End(xlToLeft).Column
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, "A").End(xlUp).Row

    For c = 2 To lastCol
        d = wsSummary.Cells(1, c).Value

        If Int(d) = Int(aWeekAgo) Then
            For r = 5 To lastRow
                key = wsSummary.Cells(r, "A").Value
                If Not isInclude(r, 2) Then weekObj(key) = wsSummary.Cells(r, c).Value
            Next r
        End If

        If Int(d) = Int(aMonthAgo) Then
            For r = 5 To lastRow
                key = wsSummary.Cells(r, "A").Value
                If Not isInclude(r, 2) Then monthObj(key) = wsSummary.Cells(r, c).Value
            Next r
            Exit For
        End If
    Next c

    '--------------------------------
    ' 주간 / 월간 변화율
    '--------------------------------
    For r = 8 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        key = ws.Cells(r, "A").Value

        If Not isInclude(r, 1) And key <> "" Then
            ws.Cells(r, "F").Value = weekObj(key)
            ws.Cells(r, "J").Value = monthObj(key)

            SetRate ws.Cells(r, "H"), ws.Cells(r, "G"), ws.Cells(r, "F")
            SetRate ws.Cells(r, "L"), ws.Cells(r, "K"), ws.Cells(r, "J")
        End If
    Next r

    '--------------------------------
    ' 합계 계산
    '--------------------------------
    Dim itemRows, ranges, applyColumns, values
    Dim i As Long, j As Long

    itemRows = Array(7, 11, 15, 20, 27, 33, 39, 47, 52, 56, 61, 66, 71, 76, _
                     79, 82, 86, 93, 97, 104, 117, 131, 135, 142)

    ranges = Array("A8:A9", "A12:A13", "A16:A18", "A21:A25", "A28:A31", _
                   "A34:A37", "A40:A45", "A48:A50", "A53:A54", "A57:A59", _
                   "A62:A64", "A67:A69", "A72:A74", "A77:A77", "A80:A80", _
                   "A83:A84", "A87:A91", "A94:A95", "A98:A102", "A105:A112", _
                    "A118:A129", "A132:A133", "A136:A137", "A143:A147")

    applyColumns = Array("B", "C", "F", "G", "J", "K")

    For i = LBound(applyColumns) To UBound(applyColumns)
        add = 0
        For j = LBound(itemRows) To UBound(itemRows)

            If j = 13 Or j = 14 Then
                ws.Cells(itemRows(j), applyColumns(i)).Value = _
                    ws.Range(Replace(ranges(j), "A", applyColumns(i))).Value
            Else
                values = ws.Range(Replace(ranges(j), "A", applyColumns(i))).Value
                ws.Cells(itemRows(j), applyColumns(i)).Value = SumArray(values)
            End If

            add = add + ws.Cells(itemRows(j), applyColumns(i)).Value

            Select Case itemRows(j)
                Case 103
                    ws.Cells(5, applyColumns(i)).Value = add: add = 0
                Case 134
                    ws.Cells(114, applyColumns(i)).Value = add: add = 0
                Case 141
                    ws.Cells(139, applyColumns(i)).Value = add
                    ws.Cells(2, applyColumns(i)).Value = _
                        ws.Cells(5, applyColumns(i)).Value + _
                        ws.Cells(114, applyColumns(i)).Value + _
                        ws.Cells(139, applyColumns(i)).Value
            End Select
        Next j
    Next i

    '--------------------------------
    ' 합계 변화율
    '--------------------------------
    For i = LBound(itemRows) To UBound(itemRows)
        SetRate ws.Cells(itemRows(i), "D"), ws.Cells(itemRows(i), "C"), ws.Cells(itemRows(i), "B")
        SetRate ws.Cells(itemRows(i), "H"), ws.Cells(itemRows(i), "G"), ws.Cells(itemRows(i), "F")
        SetRate ws.Cells(itemRows(i), "L"), ws.Cells(itemRows(i), "K"), ws.Cells(itemRows(i), "J")
    Next i

    '--------------------------------
    ' 브랜드 요약
    '--------------------------------
    Dim brandRows, combine, rowNum, a, b, t

    brandRows = Array(5, 114, 139, 2)
    combine = Array("CBD", "GFH", "KJL")

    For i = LBound(brandRows) To UBound(brandRows)
        rowNum = brandRows(i)
        For j = LBound(combine) To UBound(combine)
            a = Left(combine(j), 1)
            b = Mid(combine(j), 2, 1)
            t = Right(combine(j), 1)

            SetRate ws.Cells(rowNum, t), ws.Cells(rowNum, a), ws.Cells(rowNum, b)
        Next j
    Next i

    MsgBox "Carried out what you asked!", vbInformation
End Sub


Function CalcRate(curVal As Double, prevVal As Double) As Variant
    If curVal = 0 Then
        CalcRate = ""
    Else
        CalcRate = (curVal - prevVal) / curVal
    End If
End Function


Sub ApplyRateStyle(targetCell As Range)
    If targetCell.Value = "" Then Exit Sub

    If targetCell.Value > 0 Then
        targetCell.Interior.Color = RGB(255, 235, 238)
        targetCell.Font.Bold = True
    ElseIf targetCell.Value < 0 Then
        targetCell.Interior.Color = RGB(227, 242, 253)
        targetCell.Font.Bold = True
    Else
        targetCell.Interior.Color = RGB(245, 245, 245)
        targetCell.Font.Bold = False
    End If
End Sub


Sub SetRate(targetCell As Range, curVal As Double, prevVal As Double)
    Dim rate As Variant

    rate = CalcRate(curVal, prevVal)

    If rate = "" Then
        targetCell.Value = ""
        Exit Sub
    End If

    targetCell.Value = rate
    targetCell.NumberFormat = "0.00%"

    ApplyRateStyle targetCell
End Sub


Function SumArray(arr As Variant) As Double
    Dim i As Long, acc As Double
    For i = LBound(arr) To UBound(arr)
        acc = acc + arr(i, 1)
    Next i
    SumArray = acc
End Function
Function isInclude(target As Long, version As Long) As Boolean

    Dim i As Long
    Dim itemsRows As Variant

    If version = 1 Then

        itemsRows = Array( _
            11, 15, 20, 27, 33, 39, 47, 52, 56, 61, _
            66, 71, 76, 79, 82, 86, 93, 97, _
            104, 115, 117, 131, 135, 140, 142 _
        )

        For i = LBound(itemsRows) To UBound(itemsRows)
            If target = itemsRows(i) Then
                isInclude = True
                Exit Function
            End If
        Next i

        isInclude = False

    Else

        itemsRows = Array( _
            7, 10, 14, 20, 25, 30, 37, 41, 44, 48, _
            52, 56, 60, 62, 64, 67, 73, 76, _
            82, 91, 92, 105, 108, 111, 112 _
        )

        For i = LBound(itemsRows) To UBound(itemsRows)
            If target = itemsRows(i) Then
                isInclude = True
                Exit Function
            End If
        Next i

        isInclude = False

    End If

End Function

