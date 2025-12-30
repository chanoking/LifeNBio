Attribute VB_Name = "Add"
Sub Add()
Attribute Add.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim ws As Worksheet, wsSource As Worksheet, wsView As Worksheet, wsBlog As Worksheet
    Dim lastRow As Long, r As Long, lastRowB As Long, i As Long
    Dim sh As New Common
    
    Set ws = ThisWorkbook.Sheets("원고기입")
    Set wsSource = ThisWorkbook.Sheets("정산관리")
    Set wsView = ThisWorkbook.Sheets("조회수")
    Set wsBlog = ThisWorkbook.Sheets("블로그순위")
    
    sh.init "정산관리"
    lastRow = sh.lastRow
    
    sh.init "원고기입"
    lastRowB = sh.lastRow
    
    ws.Range("T2:X" & lastRowB).ClearContents
    
    Dim keyword, types, quote, val
    For r = 2 To lastRow
        keyword = wsSource.Cells(r, "A").value
        types = wsSource.Cells(r, "C").value
        quote = wsSource.Cells(r, "P").value
        val = wsSource.Cells(r, "S").value
        
        If quote = "" Then quote = 0
        
        If types = "메인" And quote > 0 Then
            For i = lastRowB To 2 Step -1
                If ws.Cells(i, "N").value = keyword Then
                    ws.Cells(i, "T").value = quote
                    ws.Cells(i, "U").value = val
                    Exit For
                End If
            Next i
        End If
    Next r
    
    Dim firstD As Date, firstM As Date, d As Date
    Dim w As Long
    Dim sDate As Date, eDate As Date
    Dim foundCell As Range
    
    sDate = DateSerial(2025, 12, 1)
    eDate = DateSerial(2026, 1, 0)
    
    For r = 2 To lastRowB
        d = ws.Cells(r, "B").value
        firstD = DateSerial(Year(d), month(d), 0) + 1
        firstM = firstD + (8 - Weekday(firstD, 2)) Mod 7
        w = Int((d - firstM) / 7) + 1
        ws.Cells(r, "V").value = Right(Year(d), 2) & "년 " & month(d) & "월 " & w & "주"
    Next r
    

    For r = lastRowB To 2 Step -1
        d = ws.Cells(r, "B").value
        If d >= sDate And d <= eDate Then
            key = ws.Cells(r, "N").value
            Set foundCell = wsView.Range("A:A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not foundCell Is Nothing Then
                ws.Range("W" & r & ":X" & r).value = wsView.Range("C" & foundCell.row & ":D" & foundCell.row).value
            End If
        Else
            Exit For
        End If
    Next r
    
    ws.Range("Y2:AB" & lastRowB).value = wsBlog.Range("Q2:T" & lastRowB).value
    
    MsgBox "Completo!"
End Sub
