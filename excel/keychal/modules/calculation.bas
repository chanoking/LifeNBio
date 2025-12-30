Attribute VB_Name = "calculation"
Sub cal()
Attribute cal.VB_ProcData.VB_Invoke_Func = "C\n14"
    Dim ws As Worksheet, otherWS As Worksheet, anotherWS As Worksheet
    Dim lastRow As Long, otherLastRow As Long, anotherLastRow As Long
    Dim r As Long, c As Long, startCol As Long, duration As Long
    Dim dict As Object, obj As Object
    Dim key As String, keyB As String, keyMain As String
    Dim firstDate As Date, firstMonDate As Date
    Dim visibleRow As Range
    Dim days As Long
    Dim quote
    
    Set ws = ThisWorkbook.Sheets("정산관리")
    Set otherWS = ThisWorkbook.Sheets("순위")
    Set anotherWS = ThisWorkbook.Sheets("원고기입")
    Set dict = CreateObject("Scripting.Dictionary")
    Set obj = CreateObject("Scripting.Dictionary")
    
    startCol = ws.Cells(1, "V").Column
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    otherLastRow = otherWS.Cells(otherWS.Rows.Count, "C").End(xlUp).row
    anotherLastRow = anotherWS.Cells(anotherWS.Rows.Count, "B").End(xlUp).row
    
    ' Calculate first date
    firstDate = DateSerial(2025, 12, 1)
    days = Day(DateSerial(2026, 1, 0))
    
    ' Fill date row (60 columns)
    For c = 0 To 30
        ws.Cells(1, startCol + c).value = firstDate + c
        ws.Cells(1, startCol + 31 + c).value = DateSerial(2025, 11, 1) + c
    Next c
    
    ' Calculate Q
    For r = 2 To lastRow
        If ws.Cells(r, "C").value = "서브" Then GoTo ContinueLoop
        quote = ws.Cells(r, "P").value
        If quote = "" Then quote = 0
        ws.Cells(r, "Q").value = quote / days
ContinueLoop:
    Next r
    
    ' Filter
    With anotherWS.Range("A1")
        .AutoFilter field:=2, Criteria1:=">=" & firstDate
        .AutoFilter field:=17, Criteria1:="메인"
    End With
    
    Dim visibleRange As Range
    On Error Resume Next
    Set visibleRange = anotherWS.Range("B2:B" & anotherLastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Count dictionary
    If Not visibleRange Is Nothing Then
        For Each visibleRow In visibleRange.Rows
            key = anotherWS.Cells(visibleRow.row, "F").value & "||" & anotherWS.Cells(visibleRow.row, "N").value
            obj(key) = obj(key) + 1
        Next visibleRow
    End If
    
    If anotherWS.FilterMode Then anotherWS.ShowAllData
    anotherWS.AutoFilterMode = False
    
    ' Build dict from "순위"
    For r = 2 To otherLastRow
        key = otherWS.Cells(r, "A").value & "||" & otherWS.Cells(r, "B").value
        dict(key) = otherWS.Cells(r, "C").value
    Next r
    
    ' Fill U column
    For r = 2 To lastRow
        key = ws.Cells(r, "A").value & "||" & ws.Cells(r, "L").value
        keyB = ws.Cells(r, "E").value & "||" & ws.Cells(r, "A").value
        ws.Cells(r, "U").value = IIf(dict.Exists(key), dict(key), 0)
        ws.Cells(r, "J").value = IIf(obj.Exists(keyB), obj(keyB), "")
    Next r
    
    ' Copy today's values once
    Dim todayCol As Long: todayCol = 0
    For c = startCol To startCol + 60
        If ws.Cells(1, c).value = Date Then
            todayCol = c
            Exit For
        End If
    Next c
    
    If todayCol > 0 Then
        For r = 2 To lastRow
            ws.Cells(r, todayCol).value = ws.Cells(r, "U").value
        Next r
    End If
    
    ' Calculate duration R and S
    For r = 2 To lastRow
        duration = 0
        For c = startCol To startCol + 30
            If ws.Cells(r, c).value > 0 Then duration = duration + 1
        Next c
        ws.Cells(r, "R").value = duration
        ws.Cells(r, "S").value = ws.Cells(r, "Q").value * duration
    Next r
    
    ' Look up latest writing date and label exposure
    For r = 2 To lastRow
        
        If ws.Cells(r, "C").value = "서브" Then GoTo NextRow
        
        ' Build main key
        keyMain = ws.Cells(r, "E").value & "||" & ws.Cells(r, "A").value
        
        ' Find last matching record
        ws.Cells(r, "K").value = ""
        For record = anotherLastRow To 2 Step -1
            keyB = anotherWS.Cells(record, "F").value & "||" & anotherWS.Cells(record, "N").value
            If keyMain = keyB Then
                ws.Cells(r, "K").value = anotherWS.Cells(record, "B").value
                Exit For
            End If
        Next record
        
        ' Exposure logic (O column)
        If ws.Cells(r, "U").value > 0 And ws.Cells(r, "K").value > ws.Cells(r, "N").value Then
            ws.Cells(r, "O").value = "노출"
        Else
            If ws.Cells(r, "N").value >= ws.Cells(r, "K").value Then
                ws.Cells(r, "O").value = True
            Else
                ws.Cells(r, "O").value = False
            End If
        End If
        
NextRow:
    Next r
    
    ' Calculate T column
    For r = 2 To lastRow
        If ws.Cells(r, "F").value = "세금" Then
            ws.Cells(r, "T").value = ws.Cells(r, "S").value * 1.1
        Else
            ws.Cells(r, "T").value = ws.Cells(r, "S").value * 0.967
        End If
    Next r
    
End Sub

