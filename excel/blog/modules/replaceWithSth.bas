Attribute VB_Name = "replaceWithSth"
Sub replaceWithSth()
Attribute replaceWithSth.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim ws As Worksheet
    Dim rng As Range
    Dim r As Range
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    For Each r In rng
        ' Replace "m." with ""
        ws.Cells(r.Row, r.Column).value = Replace(ws.Cells(r.Row, r.Column).value, "m.", "")
        ' Replace "https" with "http"
        ws.Cells(r.Row, r.Column).value = Replace(ws.Cells(r.Row, r.Column).value, "https", "http")
    Next r
    
End Sub

