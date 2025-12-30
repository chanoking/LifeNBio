Attribute VB_Name = "Rank"
Sub rank()

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim item As New Items          ' main dictionary: key ¡æ dict
    Dim dict As Items              ' sub dictionary: url ¡æ rank
    Dim key As Variant, url As String, rk As Variant
    Dim arr As Variant
    Dim k As Variant

    Set ws = ThisWorkbook.Sheets("¼øÀ§")

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

    MsgBox "Completed!", vbExclamation

End Sub

