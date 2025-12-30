Attribute VB_Name = "ByWeek"
Sub LetoutByWeek()
Attribute LetoutByWeek.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ws As Worksheet, wsSource As Worksheet
    Dim sh As New Common
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim d As Date
    Dim weeks(0 To 5) As Date, weeksE(0 To 5) As Date
    Dim totals(0 To 5) As Double
    
    Set ws = ThisWorkbook.Sheets("주차별")
    Set wsSource = ThisWorkbook.Sheets("정산관리")
    
    sh.init "정산관리"
    lastRow = sh.lastRow
    
    ' ====== Define weeks ======
    Dim firstD As Date
    firstD = DateSerial(2025, 12, 1)
    
    weeks(1) = firstD + (8 - Weekday(firstD, 2)) Mod 7
    weeksE(1) = weeks(1) + 6
    
    Dim i As Long
    For i = 2 To 5
        weeks(i) = weeks(i - 1) + 7
        weeksE(i) = weeksE(i - 1) + 7
    Next i
    
    weeks(0) = weeks(1) - 7
    weeksE(0) = weeksE(1) - 7
    
    ' ====== Determine last column with dates ======
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' ====== Loop through date columns ======
    For c = 22 To lastCol
        d = wsSource.Cells(1, c).value
        For i = 0 To 5
            If d >= weeks(i) And d <= weeksE(i) Then
                For r = 2 To lastRow
                    If wsSource.Cells(r, c).value > 0 Then
                        totals(i) = totals(i) + wsSource.Cells(r, "Q").value
                    End If
                Next r
            End If
        Next i
    Next c
    
    ' ====== Write results to output sheet ======
    For i = 0 To 5
        ws.Cells(i + 2, "B").value = totals(i)
    Next i
    
    MsgBox "Completo!", vbExclamation
End Sub

