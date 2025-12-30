Attribute VB_Name = "Attach"
Sub Attach()
    Dim ws As Worksheet, wsSource As Worksheet
    Dim sh As common
    Dim start As Long, last As Long
    
    Set sh = New common
    sh.init "붙이기용"
    start = sh.lastRow + 1
    
    sh.init "원고기입"
    last = sh.lastRow
    
    Set ws = ThisWorkbook.Sheets("붙이기용")
    Set wsSource = ThisWorkbook.Sheets("원고기입")
    
    wsSource.Range("A" & start & ":A" & last).Copy
    ws.Range("A" & start & ":A" & last).PasteSpecial xlPasteValues
    
    wsSource.Range("C" & start & ":H" & last).Copy
    ws.Range("B" & start & ":G" & last).PasteSpecial xlPasteValues
    
    '---- Get the dates into an array ----
    Dim dates As Variant
    dates = wsSource.Range("B" & start & ":B" & last).value
    
    Dim i As Long
    Dim years As New Collection
    Dim months As New Collection
    Dim days As New Collection
    
    For i = 1 To UBound(dates)
        years.Add Right(year(dates(i, 1)), 2)
        months.Add month(dates(i, 1))
        days.Add Day(dates(i, 1))
    Next i
    
    '---- Combine into a 2D array ----
    Dim comb As New Collection
    comb.Add years
    comb.Add months
    comb.Add days
    
    Dim arr() As Variant
    ReDim arr(1 To years.Count, 1 To comb.Count)
    
    Dim r As Long, c As Long
    
    For c = 1 To comb.Count
        For r = 1 To comb(c).Count
            arr(r, c) = comb(c)(r)
        Next r
    Next c
    
    '---- Output to sheet ----
    ws.Range("H" & start).Resize(UBound(arr, 1), UBound(arr, 2)).value = arr
    
    MsgBox "Completo!"
End Sub

