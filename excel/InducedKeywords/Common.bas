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

Public Function SumArr(arr As Variant) As Long
    Dim i As Long
    Dim aggregate As Long
    
    For i = LBound(arr) To UBound(arr)
        aggregate = aggregate + arr(i, 1)
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
        Includes = dictA.Exists(target)
    Else
        Includes = dictB.Exists(target)
    End If
End Function


