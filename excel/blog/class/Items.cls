VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dict As Object

Private Sub Class_Initialize()
    Set dict = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get AllKeys() As Variant
    AllKeys = dict.Keys
End Property

' Flexible AddItem with optional accumulation
Public Sub AddItem(key As Variant, ParamArray values() As Variant)
    Dim arr() As Variant
    Dim i As Long
    
    If Not dict.Exists(key) Then
        ' Key doesn't exist: store values as-is
        dict(key) = values
    Else
        ' Key exists: accumulate numeric values
        arr = dict(key)
        For i = LBound(values) To UBound(values)
            If IsNumeric(values(i)) Then
                arr(i) = arr(i) + values(i)   ' accumulate
            Else
                arr(i) = values(i)            ' overwrite non-numeric
            End If
        Next i
        dict(key) = arr
    End If
End Sub

Public Function Exists(key As Variant) As Boolean
    Exists = dict.Exists(key)
End Function

Public Function GetItem(key As Variant) As Variant
    GetItem = dict(key)
End Function

Public Function GetItemO(key As Variant) As Object
    Set GetItemO = dict(key)
End Function
Public Sub AddItemO(key As Variant, value As Object)
    dict.Add key, value
End Sub

Public Sub Reset()
    dict.RemoveAll
End Sub

Public Sub Update(key As Variant, ParamArray values() As Variant)
    dict(key) = values()
End Sub

Public Function cnt() As Long
    cnt = dict.Count
End Function




