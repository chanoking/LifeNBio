Attribute VB_Name = "greeting"
Sub greeting()
    Dim ws As Worksheet
    Dim infls As Variant
    
    Set ws = ThisWorkbook.Sheets("message")
    
    infls = Array("모모둥이", "민들레", "봉봉댁", "셀럽주부", "푸들")
    
    Dim i As Long, r As Long
    r = 1
    
    Dim mon As Integer
    mon = month(DateSerial(2025, 11, 1))
    Dim lastDate As Variant
    lastDate = Format(DateSerial(2025, 12, 0), "mm/dd")
    For i = LBound(infls) To UBound(infls)
        ws.Cells(r, "B").value = _
            infls(i) & "님, 안녕하세요:)" & vbLf & mon & "월 고생 많으셨고, 정산내역 전달드립니다!" _
            & vbLf & "확인 후 이상 없을 경우 아래 메일로 세금계산서 발행 부탁드립니다!" _
            & vbLf & "chano94@lifenbio.com (" & lastDate & ")" & vbLf & vbLf & "감사합니다!"
        r = r + 1
    Next i
    
    infls = Array("갬성주부", "수미지")
    For i = LBound(infls) To UBound(infls)
        ws.Cells(r, "B").value = _
            infls(i) & "님, 안녕하세요:)" & vbLf & mon & "월 고생 많으셨고, 정산내역 전달드립니다!" _
            & vbLf & "확인 부탁드립니다!" & vbLf & vbLf & "감사합니다!"
        r = r + 1
    Next i
    
    
End Sub
