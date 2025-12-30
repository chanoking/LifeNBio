Attribute VB_Name = "MarektingFee"
Sub Marketing_FeeOfKey()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim sheet As Common
    Dim myItems As Items
    Dim lastRow As Long, r As Long, dataRows As Long
    Dim name As String, brand As String
    Dim price As Variant, info As Variant, key As Variant, types As Variant, typesB As Variant, _
        quote As Variant
    
    ' ====== Setup ======
    Set sheet = New Common
    Set wsSource = ThisWorkbook.Sheets("Old_정산관리")
    Set wsDest = ThisWorkbook.Sheets("마케팅비용")
    Set myItems = New Items
    
    sheet.init "Old_정산관리"
    lastRow = sheet.lastRow
    
    ' ====== 1. Aggregate items ======
    For r = 2 To lastRow
        price = wsSource.Cells(r, "S").value
        name = Replace(wsSource.Cells(r, "E").value, " ", "")
        brand = wsSource.Cells(r, "H").value
        types = wsSource.Cells(r, "C").value
        typesB = wsSource.Cells(r, "D").value
        quote = wsSource.Cells(r, "P").value
        If quote = "" Then quote = 0
        
        If types = "메인" And quote > 0 Then
            If Not myItems.Exists(name) Then
                myItems.AddItem name, price, brand, 1
            Else
                info = myItems.GetItem(name)
                info(0) = info(0) + price        ' accumulate price
                info(2) = info(2) + 1            ' increment count
                myItems.Update name, info(0), info(1), info(2)
            End If
        End If
    Next r
    
    ' ====== 2. Write aggregated results ======
    r = 2
    For Each key In myItems.AllKeys
        info = myItems.GetItem(key)
        wsDest.Cells(r, "B").value = info(1)
        wsDest.Cells(r, "C").value = key
        wsDest.Cells(r, "H").value = info(0)
        wsDest.Cells(r, "J").value = info(2)
        r = r + 1
    Next key
    
    sheet.init wsDest.name
    lastRow = sheet.lastRow
    dataRows = lastRow - 1
    
    ' ====== 3. Fill column A with "Actual" ======
    wsDest.Range("A2:A" & lastRow).value = "Actual"
    
    ' ====== 4. Fill columns D:G ======
    Dim arrDG() As Variant
    ReDim arrDG(1 To dataRows, 1 To 4)
    
    For r = 1 To dataRows
        arrDG(r, 1) = "01.바이럴_블로그"
        arrDG(r, 2) = "키챌_월보장"
        arrDG(r, 3) = ""
        arrDG(r, 4) = "11월"
    Next r
    wsDest.Range("D2:G" & lastRow).value = arrDG
    
    ' ====== 5. Column K = H / J ======
    Dim arrH As Variant, arrJ As Variant, arrK() As Variant
    arrH = wsDest.Range("H2:H" & lastRow).value
    arrJ = wsDest.Range("J2:J" & lastRow).value
    ReDim arrK(1 To dataRows, 1 To 1)
    
    For r = 1 To dataRows
        If arrJ(r, 1) <> 0 Then
            arrK(r, 1) = arrH(r, 1) / arrJ(r, 1)
        Else
            arrK(r, 1) = 0
        End If
    Next r
    wsDest.Range("K2:K" & lastRow).value = arrK
    
    ' ====== 6. Column L constant ======
    wsDest.Range("L2:L" & lastRow).value = "1.바이럴마케팅"
    
    ' ====== 7. Clean Names in column C ======
    Dim arrName As Variant
    arrName = wsDest.Range("B2:B" & lastRow).value
    
    For r = 1 To dataRows
        If arrName(r, 1) = "파이토뉴트리" Then arrName(r, 1) = "01.파이토뉴트리"
        If arrName(r, 1) = "혜인서" Then arrName(r, 1) = "02.혜인서"
        If arrName(r, 1) = "흑보목" Then arrName(r, 1) = "03.흑보목"
    Next r
    wsDest.Range("B2:B" & lastRow).value = arrName
    
    MsgBox "Marketing Fee of Blog Completed!", vbInformation
    
End Sub


