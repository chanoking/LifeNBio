Attribute VB_Name = "FreeStatement"
Sub FreeStatement()

    Dim ws As Worksheet
    Dim myItems As Items
    Dim lastRow As Long
    Dim r As Long
    
    Dim name As String, brand As String
    Dim price As Long, cnt As Long
    Dim info As Variant, key As Variant
    
    Set ws = ThisWorkbook.Sheets("프리내역")
    Set myItems = New Items
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' =======================
    ' 1. Aggregate items
    ' =======================
    For r = 2 To lastRow
        
        name = Replace(ws.Cells(r, "B").value, " ", "")
        If InStr(name, "조인트리션") Then name = "조인트리션"
        price = ws.Cells(r, "C").value
        brand = ws.Cells(r, "A").value
        
        If Not myItems.Exists(name) Then
            myItems.AddItem name, price, brand, 1
    
        Else
            info = myItems.GetItem(name)
            info(1) = info(1) + price  ' accumulate price
            info(2) = info(2) + 1      ' count
            myItems.AddItem name, info(1), info(0), info(2)
        End If
    
    Next r
    
    
    ' =======================
    ' 2. Write back results
    ' =======================
    r = 2
    For Each key In myItems.AllKeys
        info = myItems.GetItem(key)
        
        ws.Cells(r, "F").value = info(0)    ' brand
        ws.Cells(r, "G").value = key        ' name
        ws.Cells(r, "L").value = info(1)    ' total price
        ws.Cells(r, "N").value = info(2)    ' count
        
        r = r + 1
    Next key

    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    
    ws.Range("E2:E" & lastRow).value = "Actual"
    
    
    ' =======================
    ' 3. Fill columns H:K
    ' =======================
    Dim arrHK As Variant
    ReDim arrHK(1 To lastRow - 1, 1 To 4)
    
    For r = 1 To lastRow - 1
        arrHK(r, 1) = "01.바이럴_블로그"
        arrHK(r, 2) = "프리랜서_원고"
        arrHK(r, 3) = ""                   ' column J empty?
        arrHK(r, 4) = "11월"
    Next r
    
    ws.Range("H2:K" & lastRow).value = arrHK
    
    
    ' =======================
    ' 4. Create O column (L / N)
    ' =======================
    Dim arrC As Variant
    Dim arrA As Variant
    Dim arrB As Variant
    
    arrA = ws.Range("L2:L" & lastRow).value
    arrB = ws.Range("N2:N" & lastRow).value
    
    ReDim arrC(1 To UBound(arrA), 1 To 1)
    
    For r = 1 To UBound(arrA)
        arrC(r, 1) = arrA(r, 1) / arrB(r, 1)
    Next r
    
    ws.Range("O2:O" & lastRow).value = arrC
    
    
    ' =======================
    ' 5. Column P constant
    ' =======================
    ws.Range("P2:P" & lastRow).value = "1.바이럴마케팅"
    
    
    ' =======================
    ' 6. Clean names (replace)
    ' =======================
    Dim arrName As Variant
    arrName = ws.Range("F2:G" & lastRow).value
    
    For r = 1 To UBound(arrName)
        arrName(r, 1) = Replace(arrName(r, 1), " ", "")
        
        If arrName(r, 2) = "인-칼슘앱솔브" Then arrName(r, 2) = "인칼슘앱솔브"
        If InStr(arrName(r, 2), "조인트리션") Then arrName(r, 2) = "조인트리션"
        If arrName(r, 1) = "파이토뉴트리" Then arrName(r, 1) = "01. 파이토뉴트리"
        If arrName(r, 1) = "혜인서" Then arrName(r, 1) = "02. 혜인서"
        If arrName(r, 1) = "흑보목" Then arrName(r, 1) = "03. 흑보목"
    Next r
        
    ws.Range("F2:G" & lastRow).value = arrName

End Sub

