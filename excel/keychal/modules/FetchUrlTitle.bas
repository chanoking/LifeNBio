Attribute VB_Name = "FetchUrlTitle"
Sub FetchUrlTitle()
    Dim wsMain As Worksheet, wsOther As Worksheet
    Dim lastRow As Long, r As Long
    Dim cell As Range
    Dim key As String, url As String, title As String, id As String
    Dim check As Variant
    Dim sh As New Common
    Dim item As New Items
    Dim itemB As New Items
    Dim info As Variant
    
    Set wsOther = ThisWorkbook.Sheets("정산관리")
    Set wsMain = ThisWorkbook.Sheets("원고기입")
     
    sh.init "정산관리"
    lastRow = sh.lastRow
    
    ' === LOAD DICTIONARY A (key → id) and B (id → url,title,check) ===
    For r = 2 To lastRow
        
        key = wsOther.Cells(r, "A").value
        url = wsOther.Cells(r, "L").value
        title = wsOther.Cells(r, "M").value
        id = wsOther.Cells(r, "B").value
        check = wsOther.Cells(r, "O").value
        
        If Not item.Exists(key) Then
            item.AddItem key, id
        End If
        
        If Not itemB.Exists(id) Then
            itemB.AddItem id, url, title, check
        End If
        
    Next r
    
    ' === PROCESS SELECTION ===
    For Each cell In Selection
        
        key = cell.value
        
        ' Safety: key must exist
        If Not item.Exists(key) Then
            GoTo ContinueLoop
        End If
        
        id = item.GetItem(key)(0)
        
        ' Safety: id must exist in ItemB
        If Not itemB.Exists(id) Then
            GoTo ContinueLoop
        End If
        
        info = itemB.GetItem(id)
        
        ' info = Array(url, title, check)
        If info(2) = True Then
            wsMain.Range("R" & cell.row & ":S" & cell.row).value = Array(info(0), info(1))
        End If
        
ContinueLoop:
    Next cell
    
End Sub


