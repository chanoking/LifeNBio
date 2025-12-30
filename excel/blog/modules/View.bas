Attribute VB_Name = "View"
Sub View()
Attribute View.VB_ProcData.VB_Invoke_Func = "G\n14"
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Variant
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    For Each cell In rng
        If IsNumeric(cell) And cell <> "" Then
            ws.Cells(cell.Row, "I").value = 1
        Else
            ws.Cells(cell.Row, "I").value = 0
        End If
    Next cell
    
End Sub
