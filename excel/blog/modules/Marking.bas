Attribute VB_Name = "Marking"
Sub MarkMainOrSub_SelectedRows()
Attribute MarkMainOrSub_SelectedRows.VB_ProcData.VB_Invoke_Func = "m\n14"
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long
    Dim firstRow As Long, lastRow As Long
    Dim primaryKey As String
    Dim dict As Object
    Dim key As Variant
    
    '--- make sure something is selected ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    Set ws = Selection.Worksheet
    Set rng = Selection
    
    firstRow = rng.Rows(1).Row
    lastRow = rng.Rows(rng.Rows.Count).Row
    
    '--- create dictionary ---
    Set dict = CreateObject("Scripting.Dictionary")
    
    '--- build primary key and record first occurrence ---
    For i = firstRow To lastRow
        primaryKey = ws.Cells(i, "B").value _
                   & Join(Application.Transpose(Application.Transpose(ws.Range("F" & i & ":H" & i).value)), "") _
                   & ws.Cells(i, "K").value _
                   & Join(Application.Transpose(Application.Transpose(ws.Range("O" & i & ":P" & i).value)), "")
        
        If Not dict.Exists(primaryKey) Then
            dict.Add primaryKey, i
        End If
    Next i
    
    '--- mark "메인" or "서브" ---
    For i = firstRow To lastRow
        primaryKey = ws.Cells(i, "B").value _
                   & Join(Application.Transpose(Application.Transpose(ws.Range("F" & i & ":H" & i).value)), "") _
                   & ws.Cells(i, "K").value _
                   & Join(Application.Transpose(Application.Transpose(ws.Range("O" & i & ":P" & i).value)), "")
        
        If dict(primaryKey) = i Then
            ws.Cells(i, "Q").value = "메인" ' first occurrence
        Else
            ws.Cells(i, "Q").value = "서브" ' later occurrence
        End If
    Next i
    
    MsgBox "완료되었습니다! (" & firstRow & " ~ " & lastRow & ")", vbInformation
End Sub


