Attribute VB_Name = "CalDuAndVal"
Sub calDurationAndValue()
Attribute calDurationAndValue.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim ws As Worksheet
    Dim sh As Common
    Dim r As Long, c As Long
    Dim sumCount As Long
    Const START_COL As Long = 22
    Const END_COL As Long = 52
    
    Set sh = New Common
    sh.init "沥魂包府"
    Set ws = ThisWorkbook.Sheets("沥魂包府")
    
    With ws
        ' Calculate sum per row
        For r = 2 To sh.lastRow
            sumCount = 0
            For c = START_COL To END_COL
                If .Cells(r, c).value > 0 Then sumCount = sumCount + 1
            Next c
            .Cells(r, "R").value = sumCount
        Next r
        
        ' Calculate value and tax
        For r = 2 To sh.lastRow
            .Cells(r, "S").value = .Cells(r, "R").value * .Cells(r, "Q").value
            If .Cells(r, "E").value = "技陛" Then
                .Cells(r, "T").value = .Cells(r, "S").value * 1.1
            Else
                .Cells(r, "T").value = .Cells(r, "S").value * 0.967
            End If
        Next r
    End With
End Sub

