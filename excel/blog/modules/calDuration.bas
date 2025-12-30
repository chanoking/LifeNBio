Attribute VB_Name = "calDuration"
Sub calDuration()
    Dim ws As Worksheet
    Dim sheet As New common
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long, cnt As Long
    Dim f As Range
    
    Set ws = ThisWorkbook.Sheets("블로그순위")
    
    sheet.init "블로그순위"
    lastRow = sheet.lastRow
    
    ' 날짜가 있는 열 찾기 (1행 기준)
    Set f = ws.Rows(1).Find(What:=Date, LookIn:=xlFormulas, LookAt:=xlWhole)
    
    If f Is Nothing Then
        MsgBox "오늘 날짜를 찾을 수 없습니다."
        Exit Sub
    End If
    
    lastCol = f.Column
    
    For r = 2 To lastRow
        cnt = 0
        For c = 23 To lastCol
            If ws.Cells(r, c).Value > 0 And ws.Cells(r, c).Value <> "" Then
                cnt = cnt + 1
            End If
        Next c
        ws.Cells(r, "V").Value = cnt
    Next r
    
    MsgBox "Completed!", vbExclamation
End Sub

