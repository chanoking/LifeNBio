Attribute VB_Name = "Rank"
Sub Rank()

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim item As New Items          ' main dictionary: key → dict
    Dim dict As Items              ' sub dictionary: url → rank
    Dim key As Variant, url As String, rk As Variant
    Dim arr As Variant
    Dim k As Variant

    Set ws = ThisWorkbook.Sheets("순위")

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' =========================
    ' BUILD DICTIONARY
    ' =========================
    For r = 2 To lastRow

        arr = ws.Range("A" & r & ":C" & r).value

        key = arr(1, 1)
        url = arr(1, 2)
        rk = arr(1, 3)

        If Not item.Exists(key) Then
            Set dict = New Items
            dict.AddItem url, rk
            item.AddItemO key, dict
        Else
            Set dict = item.GetItemO(key)
            If Not dict.Exists(url) Then
                dict.AddItem url, rk
            End If
        End If

    Next r

    ' =========================
    ' CLEAR OUTPUT
    ' =========================
    ws.Range("A2:C" & lastRow).ClearContents

    ' =========================
    ' WRITE BACK RESULT
    ' =========================
    r = 2
    For Each key In item.AllKeys

        Set dict = item.GetItemO(key)

        For Each k In dict.AllKeys
            rk = dict.GetItem(k)
            If IsNumeric(rk(0)) Then
                ws.Cells(r, "A").value = key
                ws.Cells(r, "B").value = k
                ws.Cells(r, "C").value = rk(0)
                r = r + 1
            End If
        Next k

    Next key
    
    Dim targetDate As Date
    
    If Weekday(Date, 2) = 1 Then
        targetDate = Date - 4
    Else
        targetDate = Date - 1
    End If
    
    Dim sh As New common
    Dim d As Date
    
    sh.init "자사최블"
    lastRow = sh.lastRow
    
    Dim wsOwn As Worksheet
    
    Set wsOwn = ThisWorkbook.Sheets("자사최블")
    
    item.Reset
    
    For r = lastRow To lastRow - 100 Step -1
        arr = wsOwn.Range("G" & r & ":K" & r).value
        key = arr(1, 1)
        d = arr(1, 2)
        rk = arr(1, 4)
        url = arr(1, 5)
        If d >= targetDate Then
            item.AddItem key, url, rk
        Else
            Exit For
        End If
    Next r

    sh.init "순위"
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    r = lastRow + 1
    Dim info As Variant
    
    For Each key In item.AllKeys
        info = item.GetItem(key)
        url = info(0)
        rk = info(1)
        If IsNumeric(rk) Then
            ws.Cells(r, "A").value = Replace(key, " ", "")
            ws.Cells(r, "B").value = url
            ws.Cells(r, "C").value = rk
            r = r + 1
        End If
    Next key
    
    MsgBox "Completed!", vbExclamation

End Sub


