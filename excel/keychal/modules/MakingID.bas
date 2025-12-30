Attribute VB_Name = "MakingID"
Sub makingIDs()

    Dim ws As Worksheet
    Dim sh As New Common
    Dim item As Object          ' Dictionary
    Dim all As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim key As Variant
    Dim info As Variant

    ' Sheet 초기화
    sh.init "정산관리"
    Set ws = ThisWorkbook.Sheets("정산관리")
    lastRow = sh.lastRow

    ' 데이터 읽기
    all = ws.Range("D2:H" & lastRow).value

    ' Dictionary 생성
    Set item = CreateObject("Scripting.Dictionary")

    ' 중복 제거 + 정보 저장
    For i = 1 To UBound(all, 1)
        key = ws.Cells(i + 1, "A").value
        
        If Not item.Exists(key) Then
            item.Add key, all(i, 1) & "|" & all(i, 2) & "|" & all(i, 3) & "|" & all(i, 4) & "|" & all(i, 5)
        End If
    Next i

    ' 결과 출력 (원본 행 기준)
    For i = 1 To UBound(all, 1)
        key = ws.Cells(i + 1, "A").value
        info = Split(item(key), "|")

        ws.Cells(i + 1, "B").value = BuildID(info)
    Next i
    
    Dim ids
    
    ids = ws.Range("B2:B" & lastRow).value
    
    Set item = Nothing
    Set item = CreateObject("Scripting.Dictionary")
    
    Dim suffix As Long
    Dim id As String
    Dim ms As String
    suffix = 1
    ms = ""
    
    Dim col As New Collection
    
    For i = 2 To lastRow
        id = ws.Cells(i, "B").value
        ms = ws.Cells(i, "C").value
        If Not item.Exists(id) And ms = "메인" Then
            suffix = 1
            item.Add id, suffix
        ElseIf item.Exists(id) And ms = "메인" Then
            suffix = suffix + 1
            item(id) = suffix
        End If
        
        ws.Cells(i, "B").value = id & item(id)
    Next i
    
End Sub

Function BuildID(info As Variant) As String

    Dim id As String
    Dim distincA As String, distincB As String
    Dim infl As String, product As String, brand As String

    ' info 구조
    distincA = info(0)
    infl = info(1)
    distincB = info(2)
    product = info(3)
    brand = info(4)

    id = ""

    ' influencer
    Select Case infl
        Case "모모둥이": id = id & "m"
        Case "민들레": id = id & "f"
        Case "봉봉댁": id = id & "b"
        Case "셀럽주부": id = id & "c"
        Case "소신있는라이프": id = id & "l"
        Case "푸들ol": id = id & "b"
        Case "갬성주부": id = id & "e"
    End Select

    ' brand
    Select Case brand
        Case "파이토뉴트리": id = id & "p"
        Case "혜인서": id = id & "h"
        Case "흑보목": id = id & "Hh"
    End Select
    
    Select Case product
        Case "블러드플로우케어": id = id & "bl"
        Case "파미로겐": id = id & "fe"
        Case "흑본전탕": id = id & "bla"
        Case "지니어스뉴": id = id & "ge"
        Case "그로우뉴": id = id & "gr"
        Case "헤모웰당": id = id & "su"
        Case "185커큐민": id = id & "nu"
        Case "맨드로포즈": id = id & "me"
        Case "조인트리션": id = id & "joi"
        Case "위이지케어": id = id & "sto"
        Case "인칼슘앱솔브": id = id & "bo"
    End Select
        
    ' 보장
    If distincA = "월보장" Then
        id = id & "g"
    Else
        id = id & "e"
    End If

    ' 세금
    If distincB = "세금" Then
        id = id & "t"
    Else
        id = id & "e"
    End If

    BuildID = id

End Function

