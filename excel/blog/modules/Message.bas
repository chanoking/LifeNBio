Attribute VB_Name = "Message"
Sub Message()
    Dim wsSub As Worksheet, wsMain As Worksheet
    Dim lastRow As Long, r As Long
    Dim temp As String, result As String
    
    Set wsSub = ThisWorkbook.Sheets("메시지")
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    result = Format(Date, "mm/dd") & vbLf & "최적"
    temp = ""
    For r = lastRow To 2 Step -1
        If wsMain.Cells(r, "B").value = Date _
            And wsMain.Cells(r, "Q").value = "메인" _
            And wsMain.Cells(r, "R").value = "위드플래닝" Then
                If Left(wsMain.Cells(r, "M").value, 1) <> temp Then
                    result = result & vbLf & vbLf & Left(wsMain.Cells(r, "M").value, 1) _
                                & "형" & vbLf & wsMain.Cells(r, "N").value
                    temp = Left(wsMain.Cells(r, "M").value, 1)
                Else
                    result = result & vbLf & wsMain.Cells(r, "N").value
                End If
        End If
    Next r
    
    wsSub.Cells(5, "A").value = result & vbLf & vbLf & "키워드 확인 부탁드립니다!"
    
End Sub
