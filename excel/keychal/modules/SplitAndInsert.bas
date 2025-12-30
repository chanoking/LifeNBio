Attribute VB_Name = "SplitAndInsert"
Option Explicit

Sub SplitAndInsert()

    Dim wsCur As Worksheet, wsMain As Worksheet
    Dim sh As New Common
    Dim lastRow As Long, lastRowB As Long, lastRowC As Long
    Dim r As Long, i As Long, j As Long
    Dim item As New Items
    Dim key As Variant
    Dim keywords As Variant, info As Variant
    Dim collect As Collection
    Dim cnt As Long
    Dim id As Variant

    ' worksheets
    Set wsCur = ThisWorkbook.Sheets("원고기입")
    Set wsMain = ThisWorkbook.Sheets("정산관리")

    ' get last rows
    sh.init "정산관리"
    lastRow = sh.lastRow

    sh.init "원고기입"
    
    lastRowB = wsCur.Cells(wsCur.Rows.Count, "Q").End(xlUp).row
    lastRowC = wsCur.Cells(wsCur.Rows.Count, "C").End(xlUp).row
    
    Dim fileNames As String
    Dim arr() As String
    For r = lastRowB + 1 To lastRowC
        arr = Split(wsCur.Cells(r, "C").value, "_")
        wsCur.Range("C" & r & ":P" & r).value = arr
        wsCur.Range("O" & r & ":P" & r).value = wsCur.Evaluate("O" & r & ":P" & r & "*1")
    Next r

    Dim x As Long, s As Long, e As Long

    ' read keywords
    If lastRowB + 1 = sh.lastRow Then
        keywords = wsCur.Cells(lastRowB + 1, "N").value
        s = 1
        e = 1
    Else
        keywords = wsCur.Range("N" & lastRowB + 1 & ":N" & sh.lastRow).value
        s = LBound(keywords)
        e = UBound(keywords)
    End If
    
    ' build item map
        
    For x = s To e
        If e = 1 Then
            If keywords <> "" Then
                For r = 2 To lastRow
                    If keywords = wsMain.Cells(r, "A").value Then
                        item.AddItem keywords, wsMain.Cells(r, "B").value
                        Exit For
                    End If
                Next r
            End If
        Else
            If keywords(x, 1) <> "" Then
                For r = 2 To lastRow
                    If keywords(x, 1) = wsMain.Cells(r, "A").value Then
                        item.AddItem keywords(x, 1), wsMain.Cells(r, "B").value
                        Exit For
                    End If
                Next r
            End If
        End If
    Next x

    ' insertion start row
    i = lastRowB + 1

    ' iterate dictionary
    For Each key In item.AllKeys
        info = item.GetItem(key)
        Set collect = New Collection

        For r = 2 To lastRow
            id = wsMain.Cells(r, "B").value
            
            ' skip unrelated keys
            If id <> info(0) Then GoTo NextRow

            If wsMain.Cells(r, "C").value = "메인" And wsMain.Cells(r + 1, "C").value = "메인" Then
                wsCur.Cells(i, "Q").value = "메인"
                i = i + 1
                Exit For
            End If
            
            If wsMain.Cells(r, "C").value <> "메인" Then
                collect.Add wsMain.Cells(r, "A").value
            Else
                Call FlushCollect(wsCur, collect, i)
            End If

NextRow:
        Next r

        ' flush remaining rows for this key
        Call FlushCollect(wsCur, collect, i)

    Next key
    
    lastRowC = wsCur.Cells(wsCur.Rows.Count, "C").End(xlUp).row
    
    wsCur.Range("B" & lastRowB + 1 & ":B" & lastRowC).value = Date
    
    wsCur.Range("B" & lastRowC & ":Q" & lastRowC).Borders(xlEdgeBottom).LineStyle = xlContinuous

    MsgBox "Completed!", vbInformation

End Sub

Private Sub FlushCollect(ws As Worksheet, collect As Collection, ByRef i As Long)

    Dim cnt As Long, j As Long

    cnt = collect.Count
    If cnt = 0 Then Exit Sub

    ws.Rows(i + 1).Resize(cnt).Insert Shift:=xlDown

    ws.Range("C" & i & ":P" & i).AutoFill _
        Destination:=ws.Range("C" & i & ":P" & i + cnt), _
        Type:=xlFillCopy

    For j = 1 To cnt
        ws.Cells(i + j, "N").value = collect(j)
        ws.Cells(i, "Q").value = "메인"
        ws.Cells(i + j, "Q").value = "서브"
    Next j

    Set collect = New Collection
    i = i + cnt + 1

End Sub


