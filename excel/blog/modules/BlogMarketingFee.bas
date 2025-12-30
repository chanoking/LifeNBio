Attribute VB_Name = "BlogMarketingFee"
Sub Blog_Marketing_Fee()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim sheet As common
    Dim myItems As Items
    Dim lastRow As Long, r As Long, dataRows As Long
    Dim dateA As Date, dateB As Date, d As Date, dateC As Date
    Dim name As String, brand As String
    Dim price As Variant, info As Variant, key As Variant
    
    ' ====== Setup ======
    Set sheet = New common
    Set wsSource = ThisWorkbook.Sheets("원고기입")
    Set wsDest = ThisWorkbook.Sheets("마케팅비용")
    Set myItems = New Items
    
    dateA = DateSerial(2025, 11, 1)
    dateB = DateSerial(2025, 12, 0)
    dateC = DateSerial(2026, 1, 0)
    
    sheet.init "원고기입"
    lastRow = sheet.lastRow
    
    ' ====== 1. Aggregate items ======
    For r = 8979 To 2 Step -1
        d = wsSource.Cells(r, "B").value
        price = wsSource.Cells(r, "U").value
        If d >= dateA And d <= dateB And IsNumeric(price) And price > 0 Then
            name = Replace(wsSource.Cells(r, "H").value, " ", "")
            brand = wsSource.Cells(r, "G").value
            
            If Not myItems.Exists(name) Then
                myItems.AddItem name, price, brand, 1
            Else
                info = myItems.GetItem(name)
                info(0) = info(0) + price        ' accumulate price
                info(2) = info(2) + 1            ' increment count
                myItems.Update name, info(0), info(1), info(2)
            End If
        End If
nextiteration:
    Next r
    
    ' ====== 2. Write aggregated results ======
    r = 2
    For Each key In myItems.AllKeys
        info = myItems.GetItem(key)
        wsDest.Cells(r, "B").value = info(1)   ' brand
        wsDest.Cells(r, "C").value = key       ' name
        wsDest.Cells(r, "H").value = info(0)   ' total price
        wsDest.Cells(r, "J").value = info(2)   ' count
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
        arrDG(r, 2) = "블로그_건바이"
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
    arrName = wsDest.Range("B2:C" & lastRow).value
    
    For r = 1 To dataRows
        arrName(r, 2) = Replace(arrName(r, 2), " ", "")
        If arrName(r, 2) = "인-칼슘앱솔브" Then arrName(r, 2) = "인칼슘앱솔브"
        If InStr(arrName(r, 2), "조인트리션") Then arrName(r, 2) = "조인트리션"
        If arrName(r, 1) = "파이토뉴트리" Then arrName(r, 1) = "01.파이토뉴트리"
        If arrName(r, 1) = "혜인서" Then arrName(r, 1) = "02.혜인서"
        If arrName(r, 1) = "흑보목" Then arrName(r, 1) = "03.흑보목"
    Next r
    wsDest.Range("B2:C" & lastRow).value = arrName
    
    wsDest.Range("O2:R" & lastRow).value = wsDest.Range("B2:E" & lastRow).value
    wsDest.Range("N2:N" & lastRow).value = "Plan"
    wsDest.Range("T2:T" & lastRow).value = "12월"
    wsDest.Range("Y2:Y" & lastRow).value = "1.바이럴마케팅"
    
    Dim cntArr() As Variant
    
    Dim totalA As Long, totalB As Long
    
    For r = 2 To lastRow
        totalA = totalA + wsDest.Cells(r, "J").value
    Next r
    
    Debug.Print dateC
    Debug.Print GetDays(dateB + 1, dateC)
    
    totalB = GetDays(dateB + 1, dateC) * 2
    Debug.Print totalB
    cntArr = wsDest.Range("J2:J" & lastRow).value
    
    Dim i As Long
    For i = 1 To UBound(cntArr, 1)   ' start at 1, not 0
        cntArr(i, 1) = cntArr(i, 1) / totalA
        Debug.Print cntArr(i, 1)
        wsDest.Cells(i + 1, "W").value = Int(totalB * cntArr(i, 1))
        wsDest.Cells(i + 1, "U").value = wsDest.Cells(i + 1, "W").value * 70000
        wsDest.Cells(i + 1, "X").value = wsDest.Cells(i + 1, "U").value / _
                                            wsDest.Cells(i + 1, "W").value
    Next i
    
    MsgBox "Completed!", vbInformation
    
End Sub


Function GetDays(dateA As Date, dateB As Date) As Long
    Dim d As Date
    Dim days As Long
    
    days = Day(dateB)
    For d = dateA To dateB
        If Weekday(d, 2) = 6 Or Weekday(d, 2) = 7 Then
            days = days - 1
        End If
    Next d
    
    GetDays = days
End Function

