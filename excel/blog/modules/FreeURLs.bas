Attribute VB_Name = "FreeURLs"
Sub FreeURLs()
    Dim ws As Worksheet, wsSource As Worksheet
    Dim sh As common
    Dim keywords As Variant, keyword As Variant
    Dim dict As Object
    Dim queue As Collection
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("FREE")
    Set wsSource = ThisWorkbook.Sheets("원고기입")
    Set sh = New common
    
    sh.init "FREE"
    Dim lastRow As Long, otherLastRow As Long
    Dim key As String
    
    lastRow = sh.lastRow
    
    sh.init "원고기입"
    otherLastRow = sh.lastRow
    
    For r = 2 To lastRow
        key = ws.Cells(r, "M").value & "||" & ws.Cells(r, "O").value
        sh.init "원고기입"
        For Record = otherLastRow To 2 Step -1
            If wsSource.Cells(Record, "B").value < DateSerial(2025, 11, 1) Then
                Exit For
            End If
            If (wsSource.Cells(Record, "N").value & "||" & wsSource.Cells(Record, "P").value) = key Then
                ws.Cells(r, "P").value = wsSource.Cells(Record, "S").value
            End If
        Next Record
    Next r
    
    MsgBox "Completo!"
End Sub

