Attribute VB_Name = "msgTool"
Sub msgTool()
    Dim wsMain As Worksheet, wsMsg As Worksheet
    Dim foundCell As Range
    Dim r As Long, lastRow As Long
    Dim key As String, keyword As String
    Dim dict As Object
    Dim infl As Variant
    Dim msg As String
    Dim ii As Long
    
    ' --- Set sheets ---
    Set wsMain = ThisWorkbook.Sheets("원고기입")
    Set wsMsg = ThisWorkbook.Sheets("message")
    
    ' --- Create dictionary ---
    Set dict = CreateObject("Scripting.Dictionary") ' key -> Collection of keywords
    
    ' --- Find today's row ---
    Set foundCell = wsMain.Range("B:B").Find(What:=Date, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "No row found for today's date."
        Exit Sub
    End If
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).row
    
    ' --- Build dictionary of key -> collection of keywords ---
    Dim col As Collection
    For r = foundCell.row To lastRow
        infl = wsMain.Cells(r, "F").value
        keyword = wsMain.Cells(r, "N").value
        ' Combine columns to form the key
        key = wsMain.Cells(r, "F").value & "||" & wsMain.Cells(r, "G").value & "||" & _
              wsMain.Cells(r, "H").value & "||" & wsMain.Cells(r, "I").value & "||" & _
              wsMain.Cells(r, "K").value & "||" & wsMain.Cells(r, "L").value & "||" & _
              wsMain.Cells(r, "M").value & "||" & wsMain.Cells(r, "O").value & "||" & _
              wsMain.Cells(r, "P").value
              
        If Not dict.Exists(key) Then
            Set col = New Collection
            col.Add keyword
            Set dict(key) = col
        Else
            dict(key).Add keyword
        End If
    Next r
    
    ' --- Prepare influencer dictionary ---
    Dim inflDict As Object
    Set inflDict = CreateObject("Scripting.Dictionary") ' influencer -> Collection of keys
    
    Dim k As Variant
    For Each k In dict.Keys
        infl = Split(k, "||")(0) ' first part is influencer
        If Not inflDict.Exists(infl) Then
            Set col = New Collection
            Set inflDict(infl) = col
        End If
        inflDict(infl).Add k
    Next k
    
    ' --- Build messages per influencer ---
    ii = 2 ' starting row in wsMsg
    Dim kw As Variant, line As String
    
    For Each infl In inflDict.Keys
        msg = "안녕하세요 " & infl & "님:)"
        
        For Each k In inflDict(infl) ' k is original key
            line = ""
            For Each kw In dict(k) ' kw is individual keyword
                If line = "" Then
                    line = kw
                Else
                    line = line & ", " & kw
                End If
            Next kw
            msg = msg & vbLf & "[" & line & "]"
        Next k
        
        msg = msg & vbLf & "전달드립니다!"
        wsMsg.Cells(ii, "A").value = msg
        ii = ii + 1
    Next infl
End Sub


