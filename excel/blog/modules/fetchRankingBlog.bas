Attribute VB_Name = "fetchRankingBlog"
Sub fetchingRankingBlog()
Attribute fetchingRankingBlog.VB_ProcData.VB_Invoke_Func = "o\n14"
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim selRng As Range
    Dim tgtLastRow As Long
    Dim dict As Object
    Dim r As Long
    Dim tgtKey As String, tgtURL As String
    Dim srcKey As String, srcURL As String
    Dim dictKey As String
    Dim tgtRank As Variant
    Dim lastCol As Long
    Dim todayCol As Long
    Dim c As Long
    Dim selRow As Long
    Dim firstRank As Variant
    Dim highestRank As Variant
    Dim duration As Long
    Dim val As Variant
    
    Set selRng = Selection
    selRow = selRng.Row
    Set wsSrc = ActiveSheet
    Set wsTgt = ThisWorkbook.Sheets("¼øÀ§")
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select two columns (keyword + url).", vbExclamation
        Exit Sub
    End If
    
    If selRng.Columns.Count <> 2 Then
        MsgBox "Please select exactly TWO columns - left = keyword, right = URL.", vbExclamation
        Exit Sub
    End If
    
    tgtLastRow = wsTgt.Cells(wsTgt.Rows.Count, "A").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    For r = 2 To tgtLastRow
        tgtKey = wsTgt.Cells(r, "A").value
        tgtURL = wsTgt.Cells(r, "B").value
        dictKey = tgtKey & "||" & tgtURL
        If Not dict.Exists(dictKey) Then
            dict.Add dictKey, wsTgt.Cells(r, "C").value
        End If
    Next r
    
    For r = 1 To selRng.Rows.Count
        srcKey = NormalizeRemoveSpaces(CStr(selRng.Cells(r, 1).value))
        srcURL = CStr(selRng.Cells(r, 2).value)
        dictKey = srcKey & "||" & srcURL
        If dict.Exists(dictKey) Then
            wsSrc.Cells(selRng.Row + r - 1, "U").value = dict(dictKey)
        Else
            wsSrc.Cells(selRng.Row + r - 1, "U").value = ""
        End If
    Next r
    
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    For c = 23 To lastCol
        If wsSrc.Cells(1, c).value = Date Then
            todayCol = c
            Exit For
        End If
    Next c
    
    For r = 1 To selRng.Rows.Count
        wsSrc.Cells(selRow + r - 1, todayCol).value = wsSrc.Cells(selRow + r - 1, "U").value
    Next r
    
    Dim dateEndCol As Long
    dateEndCol = 23
    Do While IsDate(wsSrc.Cells(1, dateEndCol).value)
        dateEndCol = dateEndCol + 1
    Loop
    dateEndCol = dateEndCol - 1

    
    For r = 1 To selRng.Rows.Count
        firstRank = ""
        highestRank = ""
        duration = 0
        
        For c = 23 To dateEndCol
            val = wsSrc.Cells(selRow + r - 1, c).value
            
            If val <> "" And IsNumeric(val) Then
                duration = duration + 1
                
                If firstRank = "" Then
                    firstRank = val
                End If
        
                If highestRank = "" Then
                    highestRank = val
                Else
                    If val < highestRank Then highestRank = val
                End If
            End If
        Next c
        
        wsSrc.Cells(selRow + r - 1, "S").value = firstRank
        wsSrc.Cells(selRow + r - 1, "T").value = highestRank
        wsSrc.Cells(selRow + r - 1, "V").value = duration
    Next r
    
                    
            
        
        
    
    MsgBox "Ranks fetched successfully.", vbInformation
    
End Sub

Private Function NormalizeRemoveSpaces(s As String) As String
    NormalizeRemoveSpaces = Replace(s, " ", "")
End Function
