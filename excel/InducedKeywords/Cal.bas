Attribute VB_Name = "Cal"
Sub Cal()
    Dim ws As Worksheet, wsRaw As Worksheet
    Dim lastRow As Long, r As Long, lastCol As Long, i As Long
    Dim arr As Variant
    
    Set ws = ThisWorkbook.Sheets("Summary")
    Set wsRaw = ThisWorkbook.Sheets("rank_raw")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    arr = Array(7, 10, 14, 20, 25, 30, 37, 41, 44, 48, 52, 56, 60, 62, 64, 67, 72, 75, 81, 90, 91, 104, 107, 110, 111)
    
    Dim targetRow As Long
    Dim foundCell As Range
    
    ws.Columns("B").Insert shift:=xlToRight
    ws.Cells(1, "B").Value = Date
    
    For r = 5 To lastRow
        targetRow = ws.Cells(r, "B").Row
        
        If Not IsInArray(targetRow, arr) Then
            Set foundCell = wsRaw.Range("A:A").Find(What:=ws.Cells(r, "A").Value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                ws.Cells(r, "B").Value = wsRaw.Cells(foundCell.Row, "B").Value
            Else
                ws.Cells(r, "B").Value = 10
            End If
        End If
    Next r
    
    Dim sum As Long, cnt As Long
    cnt = 0
    For r = 5 To lastRow + 1
        targetRow = ws.Cells(r, "B").Row
        If ws.Cells(r, "A").Value <> "" And Not IsInArray(targetRow, arr) Then
            sum = sum + ws.Cells(r, "B").Value
            cnt = cnt + 1
        Else
            If Not (ws.Cells(r, "B").Value = "" And ws.Cells(r + 1, "B").Value) Then
                ws.Cells(r - cnt - 1, "B").Value = sum
                sum = 0
                cnt = 0
            End If
        End If
    Next r
    
    Dim aRow As Long, bRow As Long, cRow As Long
    
    arr = Array(4, 7, 10, 14, 20, 25, 30, 37, 41, 44, 48, 52, 56, 60, 62, 64, 67, 72, 75, 81)
    
    sum = 0
    For i = LBound(arr) To UBound(arr)
        sum = sum + ws.Cells(arr(i), "B")
    Next i
    
    ws.Cells(3, "B").Value = sum
    
    arr = Array(91, 104, 107, 110)
    
    sum = 0
    For i = LBound(arr) To UBound(arr)
        sum = sum + ws.Cells(arr(i), "B")
    Next i
    
    ws.Cells(90, "B").Value = sum
    
    ws.Cells(110, "B").Value = ws.Cells(111, "B")
    
    ws.Cells(2, "B").Value = ws.Cells(110, "B").Value + ws.Cells(90, "B").Value + ws.Cells(3, "B").Value
    
    ws.Columns("C").Copy
    ws.Columns("B").PasteSpecial Paste:=xlPasteFormats
    
    Application.CutCopyMode = False

    MsgBox "Carried out what you asked!"
End Sub

Function IsInArray(target As Long, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = target Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function



