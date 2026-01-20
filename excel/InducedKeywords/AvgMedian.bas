Attribute VB_Name = "AvgMedian"
Sub WeekAvg()
    Dim wsSummary As Worksheet, wsWeekAvg As Worksheet
    Dim lastRow As Long, r As Long, i As Long, ii As Long
    Dim dict As Object
    Dim key, valA, valB
    Dim v1, v2
    Dim rangeArrA, rangeArrB, a, b
    
    ' 설정 초기화
    If IsEmpty(Block_Range) Then InitConfig
    
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    Set wsWeekAvg = ThisWorkbook.Sheets("WeekAvg")
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' =========================
    ' Summary → Dictionary
    ' =========================
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, "A").End(xlUp).Row
    
    rangeArrA = Array("BF", "GK")
    rangeArrB = Array("BC", "EF")
    
    For ii = LBound(rangeArrA) To UBound(rangeArrA)
        If dict Is Nothing Then
            Set dict = CreateObject("Scripting.Dictionary")
        End If
        
        Dim rng As String
        
        a = Left(rangeArrB(ii), 1)
        b = Right(rangeArrB(ii), 1)
        
        For r = 5 To lastRow
            rng = Left(rangeArrA(ii), 1) & r & ":" & Right(rangeArrA(ii), 1) & r
            If Not Includes(r, 1) Then
                key = wsSummary.Cells(r, "A").Value
            
                valA = Round(CalcAvg(wsSummary.Range(rng).Value), 0)
                
                On Error Resume Next
                valB = Application.WorksheetFunction.Median(wsSummary.Range(rng).Value)
                
                If Err.Number > 0 Then
                    valB = 0
                    Err.Clear
                End If
                
                On Error GoTo 0
            
                dict(key) = Array(valA, valB)
            End If
        Next r
        ' =========================
        ' WeekAvg 채우기
        ' =========================
        lastRow = wsWeekAvg.Cells(wsWeekAvg.Rows.Count, "A").End(xlUp).Row
    
        For r = 8 To lastRow
            key = wsWeekAvg.Cells(r, "A").Value
            rng = a & r & ":" & b & r
            If key <> "" And Not Includes(r, 2) Then
                If dict.Exists(key) Then
                    wsWeekAvg.Range(rng).Value = dict(key)
                Else
                    wsWeekAvg.Range(rng).Value = 10
                End If
            End If
        Next r
    
        ' =========================
        ' 블록 합계
        ' =========================
    
        Dim accA As Long, accB As Long
        For i = LBound(Block_Range) To UBound(Block_Range)
            v1 = wsWeekAvg.Range(Replace(Block_Range(i), "A", a)).Value
            v2 = wsWeekAvg.Range(Replace(Block_Range(i), "A", b)).Value
        
            If i = 13 Or i = 14 Then
                wsWeekAvg.Cells(Item_Rows(i), a).Value = v1
                wsWeekAvg.Cells(Item_Rows(i), b).Value = v2
            Else
                wsWeekAvg.Cells(Item_Rows(i), a).Value = SumArr(v1)
                wsWeekAvg.Cells(Item_Rows(i), b).Value = SumArr(v2)
            End If
        
            accA = accA + wsWeekAvg.Cells(Item_Rows(i), a).Value
            accB = accB + wsWeekAvg.Cells(Item_Rows(i), b).Value
        
            Select Case Item_Rows(i)
            Case 104
                wsWeekAvg.Cells(5, a).Value = accA: accA = 0
                wsWeekAvg.Cells(5, b).Value = accB: accB = 0
            Case 135
                wsWeekAvg.Cells(115, a).Value = accA: accA = 0
                wsWeekAvg.Cells(115, b).Value = accB: accB = 0
            Case 142
                wsWeekAvg.Cells(140, a).Value = accA: accA = 0
                wsWeekAvg.Cells(140, b).Value = accB: accB = 0
            End Select
        Next i

        wsWeekAvg.Cells(2, a).Value = wsWeekAvg.Cells(115, a).Value + wsWeekAvg.Cells(140, a).Value _
                                      + wsWeekAvg.Cells(5, a).Value
        wsWeekAvg.Cells(2, b).Value = wsWeekAvg.Cells(115, b).Value + wsWeekAvg.Cells(140, b).Value _
                                      + wsWeekAvg.Cells(5, b).Value
        Set dict = Nothing
        
    Next ii
    
    
    MsgBox "Carried out what you asked!"
End Sub

Function CalcAvg(arr) As Double
    Dim i As Long
    Dim sum As Long
    
    For i = LBound(arr, 2) To UBound(arr, 2)
        sum = sum + arr(1, i)
    Next i
    
    CalcAvg = sum / (UBound(arr, 2) - LBound(arr, 2) + 1)
End Function


