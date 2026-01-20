Attribute VB_Name = "RateOfChange"
Sub RateOfChange()

    ' 설정 초기화
    If IsEmpty(Block_Range) Then InitConfig
    
    Dim ws As Worksheet, wsSource As Worksheet, wsSummary As Worksheet
    Dim dict As Object, weekObj As Object, monthObj As Object
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long, add As Double
    Dim key, view
    Dim aWeekAgo As Date, aMonthAgo As Date
    Dim d As Date

    Set ws = ThisWorkbook.Sheets("Rate of change")
    Set wsSource = ThisWorkbook.Sheets("view_raw")
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    Set dict = CreateObject("Scripting.Dictionary")

    '--------------------------------
    ' rank_raw → Dictionary
    '--------------------------------
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        key = wsSource.Cells(r, "A").Value
        view = wsSource.Cells(r, "B").Value
        If Not dict.Exists(key) Then dict.add key, view
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

        If Not Includes(r, 2) And key <> "" Then
            If dict.Exists(key) Then
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
                If Not Includes(r, 1) Then weekObj(key) = wsSummary.Cells(r, c).Value
            Next r
        End If

        If Int(d) = Int(aMonthAgo) Then
            For r = 5 To lastRow
                key = wsSummary.Cells(r, "A").Value
                If Not Includes(r, 1) Then monthObj(key) = wsSummary.Cells(r, c).Value
            Next r
            Exit For
        End If
    Next c

    '--------------------------------
    ' 주간 / 월간 변화율
    '--------------------------------
    For r = 8 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        key = ws.Cells(r, "A").Value

        If Not Includes(r, 2) And key <> "" Then
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

    applyColumns = Array("C", "F", "G", "J", "K")

    For i = LBound(applyColumns) To UBound(applyColumns)
        add = 0
        For j = LBound(Item_Rows) To UBound(Item_Rows)

            If j = 13 Or j = 14 Then
                ws.Cells(Item_Rows(j), applyColumns(i)).Value = _
                    ws.Range(Replace(Block_Range(j), "A", applyColumns(i))).Value
            Else
                values = ws.Range(Replace(Block_Range(j), "A", applyColumns(i))).Value
                ws.Cells(Item_Rows(j), applyColumns(i)).Value = SumArr(values)
            End If

            add = add + ws.Cells(Item_Rows(j), applyColumns(i)).Value

            Select Case Item_Rows(j)
                Case 104
                    ws.Cells(5, applyColumns(i)).Value = add: add = 0
                Case 135
                    ws.Cells(115, applyColumns(i)).Value = add: add = 0
                Case 142
                    ws.Cells(140, applyColumns(i)).Value = add
                    ws.Cells(2, applyColumns(i)).Value = _
                        ws.Cells(5, applyColumns(i)).Value + _
                        ws.Cells(115, applyColumns(i)).Value + _
                        ws.Cells(140, applyColumns(i)).Value
            End Select
        Next j
    Next i

    '--------------------------------
    ' 합계 변화율
    '--------------------------------
    For i = LBound(Item_Rows) To UBound(Item_Rows)
        SetRate ws.Cells(Item_Rows(i), "D"), ws.Cells(Item_Rows(i), "C"), ws.Cells(Item_Rows(i), "B")
        SetRate ws.Cells(Item_Rows(i), "H"), ws.Cells(Item_Rows(i), "G"), ws.Cells(Item_Rows(i), "F")
        SetRate ws.Cells(Item_Rows(i), "L"), ws.Cells(Item_Rows(i), "K"), ws.Cells(Item_Rows(i), "J")
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
