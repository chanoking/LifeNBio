Attribute VB_Name = "AddDailyValue"
Sub AddDailyValue()
    Dim wsSummary As Worksheet, wsView_raw As Worksheet
    Dim itemRows, brandRows, key, view, rangeArr
    Dim dict As Object
    Dim i As Long, lastRow As Long, r As Long, target As Long, acc As Long
    
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    Set wsView_raw = ThisWorkbook.Sheets("view_raw")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsView_raw.Cells(wsView_raw.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        key = wsView_raw.Cells(r, "A").Value
        view = wsView_raw.Cells(r, "B").Value
        dict.add key, view
    Next r
    
    itemRows = Array(4, 7, 10, 14, 20, 25, 30, 37, 41, 44, 48, 52, 56, 60, 62, 64, 67, 73, 76, 82, 92, 105, 108, 112)
    rangeArr = Array("B5:B6", "B8:B9", "B11:B13", "B15:B19", "B21:B24", "B26:B29", "B31:B36", "B38:B40", "B42:B43", "B45:B47", "B49:B51", _
                        "B53:B55", "B57:B59", 61, 63, "B65:B66", "B68:B72", "B74:B75", "B77:B81", "B83:B90", "B93:B104", _
                        "B106:B107", "B109:B110", "B113:B117")
    brandRows = Array(3, 91, 111)
    
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, "A").End(xlUp).Row
    
    wsSummary.Columns("B").Insert shift:=xlToRight
    wsSummary.Cells(1, "B").Value = Date
    
    For r = 5 To lastRow
        key = wsSummary.Cells(r, "A").Value
        If Not includes(itemRows, r) And Not includes(brandRows, r) Then
            If Not dict.exists(key) Then
                wsSummary.Cells(r, "B").Value = 10
            Else
                wsSummary.Cells(r, "B").Value = dict(key)
            End If
        End If
    Next r
    
    For i = LBound(itemRows) To UBound(itemRows)
        If itemRows(i) = 60 Or itemRows(i) = 62 Then
            wsSummary.Cells(itemRows(i), "B").Value = wsSummary.Cells(rangeArr(i), "B").Value
            acc = acc + wsSummary.Cells(rangeArr(i), "B").Value
        Else
            wsSummary.Cells(itemRows(i), "B").Value = sum(wsSummary.Range(rangeArr(i)).Value)
            acc = acc + wsSummary.Cells(itemRows(i), "B").Value
        Select Case itemRows(i)
            Case 82: wsSummary.Cells(3, "B").Value = acc: acc = 0
            Case 108: wsSummary.Cells(91, "B").Value = acc: acc = 0
            Case 112: wsSummary.Cells(111, "B").Value = acc: acc = 0
        End Select
        End If
    Next i
    
    wsSummary.Cells(2, "B").Value = wsSummary.Cells(3, "B").Value + wsSummary.Cells(91, "B").Value _
                                        + wsSummary.Cells(111, "B").Value
    
    wsSummary.Columns("C").Copy
    wsSummary.Columns("B").PasteSpecial Paste:=xlPasteFormats
    
    Application.CutCopyMode = False
    
    MsgBox "Carried out what you asked!"

    
End Sub

Function sum(arr) As Long
    Dim i As Long, result As Long
    
    For i = LBound(arr) To UBound(arr)
        result = result + arr(i, 1)
    Next i
    
    sum = result
    
End Function

Function includes(arr, target) As Boolean
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If target = arr(i) Then
            includes = True
            Exit Function
        End If
    Next i
    
    includes = False
End Function
