Attribute VB_Name = "InputOfBlog"
Sub InputOfBlog()
Attribute InputOfBlog.VB_ProcData.VB_Invoke_Func = "B\n14"
    Dim wsInput As Worksheet, wsMain As Worksheet, wsKeywords As Worksheet, wsURLs As Worksheet
    Dim myItems As Items
    Dim mainData As Variant, keywordData As Variant, urlData As Variant
    Dim lastRowMain As Long, lastRowKeyword As Long, lastRowURLs As Long
    Dim r As Long
    Dim key As Variant, val As Variant, valB As Variant, info As Variant, url As Variant
    Dim header As Variant
    Dim rOut As Long
    
    Set wsInput = ActiveSheet
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    Set wsKeywords = ThisWorkbook.Sheets("Keywords")
    Set wsURLs = ThisWorkbook.Sheets("URLs")
    Set myItems = New Items
    
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    mainData = wsMain.Range("H2:S" & lastRowMain).value
    
    lastRowKeyword = wsKeywords.Cells(wsKeywords.Rows.Count, "C").End(xlUp).Row
    keywordData = wsKeywords.Range("B2:C" & lastRowKeyword).value

    lastRowURLs = wsURLs.Cells(wsURLs.Rows.Count, "C").End(xlUp).Row
    urlData = wsURLs.Range("A2:D" & lastRowURLs).value
    
    wsInput.Cells.ClearContents
    
    If wsInput.name = "key" Then
        ' ====== Load Main raw into array ======
        If lastRowMain < 2 Then Exit Sub
        
        For r = 1 To UBound(mainData, 1)
            key = Replace(mainData(r, 7), " ", "")
            val = Replace(mainData(r, 1), " ", "")
            If val <> "혼합" Then
                myItems.AddItem key, val
            End If
        Next r
        
        For r = 1 To UBound(keywordData, 1)
            key = keywordData(r, 2)
            myItems.AddItem key, keywordData(r, 1)
        Next r
        
        rOut = 2
        
        For Each key In myItems.AllKeys
            info = myItems.GetItem(key)
            wsInput.Cells(rOut, "C").value = key
            wsInput.Cells(rOut, "B").value = info(0)
            wsInput.Cells(rOut, "A").value = "블로그"
            wsInput.Cells(rOut, "D").value = "null"
            rOut = rOut + 1
        Next key
        
        header = Array("구분", "제품", "키워드", "우선순위")
        
        Dim c As Long
        For c = 1 To 4
            wsInput.Cells(1, c).value = header(c - 1)
        Next c
    ElseIf wsInput.name = "url" Then
        header = Array("제품", "키워드", "파트", "url")
        For c = 1 To 4
            wsInput.Cells(1, c).value = header(c - 1)
        Next c
        
        If lastRowMain < 2 Then Exit Sub

        For r = 1 To UBound(mainData, 1)
            key = Replace(mainData(r, 7), " ", "") & "||" & mainData(r, 12)
            val = Replace(mainData(r, 1), " ", "")
            valB = Replace(mainData(r, 7), " ", "")
            url = mainData(r, 12)
            If url = "" Then url = "blank"
            If val <> "혼합" Then
                myItems.AddItem key, val, valB, url
            End If
        Next r
        
        For r = 1 To UBound(urlData, 1)
            key = urlData(r, 2) & "||" & urlData(r, 4)
            myItems.AddItem key, urlData(r, 1), urlData(r, 2), urlData(r, 4)
        Next r
        
        rOut = 2
        
        For Each key In myItems.AllKeys
            info = myItems.GetItem(key)
            wsInput.Cells(rOut, "A").value = info(0)
            wsInput.Cells(rOut, "B").value = info(1)
            wsInput.Cells(rOut, "D").value = info(2)
            wsInput.Cells(rOut, "C").value = "블로그"
            rOut = rOut + 1
        Next key
    Else
        wsInput.Cells(1, "A").value = "키워드"
        wsInput.Cells(1, "B").value = "URL"
        If lastRowMain < 2 Then Exit Sub
        
        For r = 1 To UBound(mainData, 1)
            key = Replace(mainData(r, 7), " ", "") & "||" & Replace(mainData(r, 12), " ", "")
            val = Replace(mainData(r, 7), " ", "")
            url = mainData(r, 12)
            If url = "" Then url = "blank"
            myItems.AddItem key, val, url
        Next r
         
        rOut = 2
        
        For Each key In myItems.AllKeys
            info = myItems.GetItem(key)
            wsInput.Cells(rOut, "A").value = info(0)
            wsInput.Cells(rOut, "B").value = info(1)
            rOut = rOut + 1
        Next key
    End If
            
    MsgBox "Completo!", vbExclamation
End Sub

