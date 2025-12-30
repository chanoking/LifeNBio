Attribute VB_Name = "SortingByOneColumn"
Sub SortSelectionRangeByFirstColumn()
Attribute SortSelectionRangeByFirstColumn.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim sortRange As Range
    Dim keyColumn As Range
    
    Set sortRange = Selection
    Set keyColumn = sortRange.Columns(1)
    
    sortRange.Worksheet.sort.SortFields.Clear
    sortRange.Worksheet.sort.SortFields.Add key:=keyColumn, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    With sortRange.Worksheet.sort
        .SetRange sortRange
        .header = xlNo
        .Apply
    End With
    
    MsgBox "Selected range sorted by first column."
End Sub
