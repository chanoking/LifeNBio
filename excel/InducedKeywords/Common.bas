Attribute VB_Name = "Common"
Public Block_Range As Variant
Public Item_Rows As Variant

Sub InitConfig()
    Block_Range = Array("A8:A9", "A12:A13", "A16:A18", "A21:A25", "A28:A31", _
                   "A34:A37", "A40:A45", "A48:A50", "A53:A54", "A57:A59", _
                   "A62:A64", "A67:A69", "A72:A74", "A77:A77", "A80:A80", _
                   "A83:A84", "A87:A91", "A94:A95", "A98:A102", "A105:A112", _
                    "A118:A129", "A132:A133", "A136:A137", "A143:A147")
    Item_Rows = Array(7, 11, 15, 20, 27, 33, 39, 47, 52, 56, 61, 66, 71, 76, _
                     79, 82, 86, 93, 97, 104, 117, 131, 135, 142)
End Sub

Public Function SumArr(arr As Variant) As Double
    Dim i As Long
    Dim aggregate As Double
    
    For i = LBound(arr) To UBound(arr)
        'Debug.Print aggregate
        aggregate = aggregate + arr(i, 1)
        'Debug.Print arr(i, 1), aggregate
    Next i
    
    SumArr = aggregate
End Function

Public Function Includes(target As Long, num As Long) As Boolean
    Static dictA As Object, dictB As Object
    Dim src, i As Long
    
    If dictA Is Nothing Then
        Set dictA = CreateObject("Scripting.Dictionary")
        src = Array(7, 10, 14, 20, 25, 30, 37, 41, 44, 48, 52, 56, 60, _
                    62, 64, 67, 73, 76, 82, 91, 92, 105, 108, 111, 112)
        For i = LBound(src) To UBound(src)
            dictA(src(i)) = True
        Next
    End If
    
    If dictB Is Nothing Then
        Set dictB = CreateObject("Scripting.Dictionary")
        src = Array(11, 15, 20, 27, 33, 39, 47, 52, 56, 61, 66, 71, 76, _
                    79, 82, 86, 93, 97, 104, 115, 117, 131, 135, 140, 142)
        For i = LBound(src) To UBound(src)
            dictB(src(i)) = True
        Next
    End If
    
    If num = 1 Then
        Includes = dictA.exists(target)
    Else
        Includes = dictB.exists(target)
    End If
End Function
Public Sub FormatPainting_Common(targetCell As Range, val As Double)
    If val > 0 Then
        With targetCell
            .Interior.Color = RGB(255, 200, 200)
            .Font.Bold = True
        End With
    ElseIf val < 0 Then
        With targetCell
            .Interior.Color = RGB(214, 233, 255)
            .Font.Bold = True
        End With
    Else
        With targetCell
            .Interior.Color = RGB(235, 235, 235)
            .Font.Bold = True
        End With
    End If
End Sub
Public Sub Select_Case(rowNum As Long, ws As Worksheet, acc As Double, col As String)
    Select Case rowNum
    Case 104
        ws.Cells(5, col).Value = acc
        Call FormatPainting_Common(ws.Cells(5, col), acc)
    Case 135
        ws.Cells(115, col).Value = acc
        Call FormatPainting_Common(ws.Cells(115, col), acc)
    Case 142
        ws.Cells(140, col).Value = acc
        Call FormatPainting_Common(ws.Cells(140, col), acc)
    End Select
    
    If rowNum = 142 Then
        ws.Cells(2, col).Value = ws.Cells(5, col).Value + ws.Cells(115, col).Value _
                                    + ws.Cells(140, col).Value
        Call FormatPainting_Common(ws.Cells(2, col), ws.Cells(2, col).Value)
    End If
End Sub
