Attribute VB_Name = "AverageChange"
Sub AverageChange()
    If IsEmpty(Block_Range) Then InitConfig
    
    Dim wsTrend As Worksheet
    Dim wsRef As Worksheet
    Dim dict As Object
    Dim r As Long, c As Long, i As Long, lastRow As Long, acc As Double
    Dim key
    Dim cell As Range
    
    Set wsTrend = ThisWorkbook.Sheets("Trend")
    Set wsRef = ThisWorkbook.Sheets("Summary")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsRef.Cells(wsRef.Rows.Count, "A").End(xlUp).Row
    
    For c = 2 To 19
        For r = 5 To lastRow
            If Not Includes(r, 1) Then
                key = wsRef.Cells(r, "A").Value
                dif = wsRef.Cells(r, c).Value - wsRef.Cells(r, c + 1)
                If dict.exists(key) Then
                    dict(key) = dict(key) + dif
                Else
                    dict(key) = dif
                End If
            End If
        Next r
    Next c
    
    lastRow = wsTrend.Cells(wsTrend.Rows.Count, "A").End(xlUp).Row
    
    For r = 8 To lastRow
        key = wsTrend.Cells(r, "A").Value
        If Not Includes(r, 2) And key <> "" Then
            Set cell = wsTrend.Cells(r, "B")
            cell.Value = Round(dict(key) / 19, 2)
            Call FormatPainting_Common(cell, cell.Value)
        End If
    Next r
    
    For i = LBound(Block_Range) To UBound(Block_Range)
        If i = 13 Or i = 14 Then
            wsTrend.Cells(Item_Rows(i), "B").Value = wsTrend.Range(Replace(Block_Range(i), "A", "B")).Value
            Call FormatPainting_Common(wsTrend.Cells(Item_Rows(i), "B"), wsTrend.Cells(Item_Rows(i), "B").Value)
            acc = acc + wsTrend.Cells(Item_Rows(i), "B").Value
        Else
            wsTrend.Cells(Item_Rows(i), "B").Value = SumArr(wsTrend.Range(Replace(Block_Range(i), "A", "B")).Value)
            'Debug.Print SumArr(wsTrend.Range(Replace(Block_Range(i), "A", "B")).Value)
            Call FormatPainting_Common(wsTrend.Cells(Item_Rows(i), "B"), wsTrend.Cells(Item_Rows(i), "B").Value)
            acc = acc + wsTrend.Cells(Item_Rows(i), "B").Value
        End If
        
        Dim temp As Long
        temp = Item_Rows(i)
        Call Select_Case(temp, wsTrend, acc, "B")
        acc = 0
    Next i
    
    
    
    MsgBox "Carried out what you asked!"
End Sub
