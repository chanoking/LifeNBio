Attribute VB_Name = "ShowToInfls"
Sub ShowToInfl()
Attribute ShowToInfl.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim ws As Worksheet
    Dim sh As Common
    Dim r As Long
    Dim keywords As Long, valueSum As Long
    Dim taxA As Long, taxB As Long, finalValue As Long
    Dim infl As String
    Dim message As String
    
    Set ws = ActiveSheet
    Set sh = New Common
    sh.init ws.name
    
    valueSum = 0
    keywords = 0
    
    ' Calculate total value and number of rows
    For r = 2 To sh.lastRow
        keywords = keywords + 1
        valueSum = valueSum + ws.Cells(r, "M").value
    Next r
    
    infl = ws.Cells(2, "D").value
    
    If isOrNot(infl) Then
        ' For specific influencers
        taxA = Application.WorksheetFunction.RoundDown(valueSum * 0.03, -1)
        taxB = Application.WorksheetFunction.RoundDown(taxA * 0.03, -1)
        finalValue = valueSum - taxA - taxB
        
        message = infl & " 님, 안녕하세요:)" _
                    & vbLf & "11월 정산내역 전달드립니다!" _
                    & vbLf & vbLf & "진행건수: " & keywords _
                    & vbLf & "정산금액: " & valueSum _
                    & vbLf & "사업소득세: " & taxA _
                    & vbLf & "지방소득세: " & taxB _
                    & vbLf & "지급액: " & finalValue _
                    & vbLf & vbLf & "확인 부탁드립니다." _
                    & vbLf & vbLf & "감사합니다."
    Else
        ' For others
        taxA = valueSum * 0.1
        message = infl & " 님, 안녕하세요:)" _
                    & vbLf & "11월 정산내역 전달드립니다!" _
                    & vbLf & vbLf & "진행건수: " & keywords _
                    & vbLf & "정산금액: " & valueSum _
                    & vbLf & "VAT: " & taxA _
                    & vbLf & "지급액: " & (valueSum + taxA) _
                    & vbLf & vbLf & "확인 후 이상없을 경우 세금계산서 아래 메일로 전달 부탁드립니다." _
                    & vbLf & "chano94@lifenbio.com" _
                    & vbLf & vbLf & "감사합니다."
    End If
    
    ws.Cells(12, "D").value = message
End Sub


Function isOrNot(infl As String) As Boolean
    Dim arr As Variant
    arr = Array("수미지", "갬성주부")
    isOrNot = Not IsError(Application.Match(infl, arr, 0))
End Function

