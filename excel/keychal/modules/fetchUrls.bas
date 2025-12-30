Attribute VB_Name = "fetchUrls"
Sub fetchUrls()
Attribute fetchUrls.VB_ProcData.VB_Invoke_Func = "i\n14"
    Dim wsThis As Worksheet, wsSource As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim r As Long, lastRow As Long
    Dim result As String, intoOne As String
    Dim dict As Object
    
    Set wsThis = ThisWorkbook.Sheets("FREE")
    Set wsSource = ThisWorkbook.Sheets("원고기입")
    Set dict = CreateObject("Scripting.Dictionary")
    Set rng = Selection
    
    arr = rng.value
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
    
    ' Build dictionary from source sheet
    For r = lastRow To 2 Step -1
        intoOne = Join(Application.Index(wsSource.Range("C" & r & ":P" & r).value, 1, 0), "")
        If Not dict.Exists(intoOne) Then
            dict.Add intoOne, wsSource.Range("R" & r).value
        End If
    Next r
    
    ' Loop through selected cells and write results
    For r = 1 To UBound(arr, 1)
        result = Join(Application.Index(arr, r, 0), "")
        If dict.Exists(result) Then
            wsThis.Cells(r + rng.row - 1, "P").value = dict(result)
        Else
            wsThis.Cells(r + rng.row - 1, "P").value = "Not Yet"
        End If
    Next r
End Sub

