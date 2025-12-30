Attribute VB_Name = "Marketing_Plan"
Sub Marketing_Plan_Fee()
    Dim wsMarketing As Worksheet, wsSource As Worksheet
    Dim sh As New Common
    Dim myItems As New Items
    Dim r As Long
    
    Set wsMarketing = ThisWorkbook.Sheets("마케팅비용")
    Set wsSource = ThisWorkbook.Sheets("정산관리")
    
    sh.init "정산관리"
    
    Dim key As Variant, val As Variant, brand As Variant, _
    types As Variant, info As Variant, typeB As Variant, cnt As Variant
    
    For r = 2 To sh.lastRow
    
        key = wsSource.Cells(r, "G").value
        val = wsSource.Cells(r, "P").value
        brand = wsSource.Cells(r, "H").value
        typesB = wsSource.Cells(r, "D").value
        
        If val = "" Then val = 0
        cnt = 1
        If IsEmpty(val) Then GoTo nextiteration
        If wsSource.Cells(r, "C").value = "메인" And val > 0 And key <> "리버티엑스" Then
            If Not myItems.Exists(key) Then
                myItems.AddItem key, val, brand, typesB, cnt
            Else
                info = myItems.GetItem(key)
                info(0) = info(0) + val
                info(3) = info(3) + 1
                myItems.Update key, info(0), info(1), info(2), info(3)
            End If
        End If
nextiteration:
    Next r
    
    r = 2
    For Each key In myItems.AllKeys
        info = myItems.GetItem(key)
        wsMarketing.Cells(r, "O").value = "Plan"
        
        If info(1) = "파이토뉴트리" Then
            wsMarketing.Cells(r, "P").value = "01." & info(1)
        ElseIf info(1) = "혜인서" Then
            wsMarketing.Cells(r, "P").value = "02." & info(1)
        Else
            wsMarketing.Cells(r, "P").value = "03." & info(1)
        End If
        
        wsMarketing.Cells(r, "Q").value = key
        wsMarketing.Cells(r, "R").value = "01.바이럴_블로그"
        
        If info(2) = "월보장" Then
            wsMarketing.Cells(r, "S").value = "키챌_월보장"
        Else
            wsMarketing.Cells(r, "S").value = "키챌_건바이"
        End If
        
        wsMarketing.Cells(r, "U").value = "12월"
        
        wsMarketing.Cells(r, "V").value = info(0)
        wsMarketing.Cells(r, "X").value = info(3)
        wsMarketing.Cells(r, "Y").value = info(0) / info(3)
        wsMarketing.Cells(r, "Z").value = "1.바이럴마케팅"
        
        r = r + 1
    Next key
End Sub
