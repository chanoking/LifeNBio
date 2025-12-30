Attribute VB_Name = "MessageTool"
Sub transaction()
Attribute transaction.VB_ProcData.VB_Invoke_Func = "T\n14"
    Dim wsMain As Worksheet, wsMsg As Worksheet
    Dim lastRow As Long, r As Long, cnt As Long, cntSpecial As Long, A As Long, B As Long, Aa As Long, _
        Bb As Long, valueSum As Long
    Dim result As String, formatDate As String, startDate As String, lastDate As String
    Dim standardDate As Date
    
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    Set wsMsg = ThisWorkbook.Sheets("메시지")
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, "V").End(xlUp).Row
    standardDate = wsMain.Cells(lastRow, "V").value
    formatDate = Format(standardDate, "mm/dd")
    sDate = Format(standardDate - 7, "mm/dd")
    lDate = Format(standardDate - 3, "mm/dd")
    
    
    result = "결재일" & vbLf & formatDate & vbLf & vbLf & "업체" & vbLf & _
                    "위드플래닝" & vbLf & vbLf
    
    For r = lastRow To 2 Step -1
        If wsMain.Cells(r, "V").value <> standardDate And IsDate(wsMain.Cells(r, "V").value) Then
            Exit For
        End If
        If wsMain.Cells(r, "V").value = standardDate And wsMain.Cells(r, "R").value = "위드플래닝" Then
            cnt = cnt + 1
            If Left(wsMain.Cells(r, "M").value, 1) = "A" Then
                A = A + 1
            Else
                B = B + 1
            End If
            If IsNumeric(wsMain.Cells(r, "U").value) And wsMain.Cells(r, "U").value > 0 Then
                If Left(wsMain.Cells(r, "M").value, 1) = "A" Then
                    Aa = Aa + 1
                Else
                    Bb = Bb + 1
                End If
                cntSpecial = cntSpecial + 1
                valueSum = valueSum + wsMain.Cells(r, "U").value
            End If
        End If
    Next r
    
    result = result & "진행 건수" & vbLf & cnt & vbLf & vbLf & "노출건수" & vbLf & cntSpecial _
                & vbLf & vbLf & "A형" & vbLf & Aa & " of " & A & vbLf _
                & vbLf & "B형" & vbLf & Bb & " of " & B & vbLf & vbLf & "비용" & vbLf & Format(valueSum, "#,##0") & " (VAT Included: " _
                & Format(valueSum * 1.1, "#,##0") & ")" & vbLf & vbLf & sDate & " ~ " & lDate & " 내역 전달드립니다!" & vbLf & vbLf & "확인 후 이상없을 경우 세금계산서 발행 부탁드립니다" & vbLf _
                & "발송 메일: chano94@lifenbio.com"
                
    wsMsg.Cells(5, "B").value = result
                
End Sub
