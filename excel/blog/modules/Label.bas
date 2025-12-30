Attribute VB_Name = "Label"
Sub MakeLabel()
Attribute MakeLabel.VB_ProcData.VB_Invoke_Func = "M\n14"
    Dim ws As Worksheet
    Dim rng As Range
    Dim firstRow As Long, lastRow As Long, i As Long
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If

    Set ws = Selection.Worksheet
    Set rng = Selection
    
    firstRow = rng.Rows(1).Row
    lastRow = rng.Rows(rng.Rows.Count).Row
    
    For i = firstRow To lastRow
        ws.Cells(i, "A").value = ws.Cells(i, "C").value _
                                & "/" _
                                & Join(Application.Transpose(Application.Transpose(ws.Range("G" & i & ":H" & i).value)), "_") _
                                & "/" _
                                & Join(Application.Transpose(Application.Transpose(ws.Range("E" & i & ":F" & i).value)), "_") _
                                & "/" _
                                & Join(Application.Transpose(Application.Transpose(ws.Range("J" & i & ":L" & i).value)), "_") _
                                & "_" _
                                & ws.Cells(i, "P").value _
                                & "/" _
                                & ws.Cells(i, "N").value & "_" & ws.Cells(i, "S").value
        Next i
    
        
                                
                                
                                
End Sub






