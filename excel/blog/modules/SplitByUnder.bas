Attribute VB_Name = "SplitByUnder"
Sub SplitTextByUnderbar()
Attribute SplitTextByUnderbar.VB_ProcData.VB_Invoke_Func = "y\n14"
    Dim cell As Range
    Dim parts As Variant
    Dim i As Integer
    
    For Each cell In Selection
        parts = split(cell.value, "_")
        For i = LBound(parts) To UBound(parts)
            cell.Offset(0, i).value = Trim(parts(i))
        Next i
    Next cell
End Sub

