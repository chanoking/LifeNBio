VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pws As Worksheet

Public Sub init(sheetName As String)
    Set pws = ThisWorkbook.Sheets(sheetName)
End Sub

Public Property Get lastRow() As Long
    lastRow = pws.Cells(pws.Rows.Count, "H").End(xlUp).Row
End Property

