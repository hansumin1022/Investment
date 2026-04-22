' ============================================================
' RISE 200위클리커버드콜
' ============================================================
Function GetRise200Nav(Optional field As String = "price") As Variant
    GetRise200Nav = GetNaverFinanceData("490070", field)
End Function
' ============================================================
' ACE KRX 금현물
' ============================================================
Function GetACEGoldNav(Optional field As String = "price") As Variant
    GetACEGoldNav = GetNaverFinanceData("411060", field)
End Function

' ============================================================
' 공통 파싱 함수
' ============================================================
Function GetNaverFinanceData(itemCode As String, Optional field As String = "price") As Variant
    Dim xml As Object, html As String, re As Object, m As Object
    Dim url As String: url = "https://finance.naver.com/item/main.naver?code=" & itemCode
    
    On Error GoTo ErrHandler
    Set xml = CreateObject("MSXML2.XMLHTTP")
    xml.Open "GET", url, False
    xml.send
    
    ' 바이너리 데이터를 텍스트로 변환 (EUC-KR)
    html = ConvertEucKrToUtf8(xml.responseBody)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    Select Case LCase(Trim(field))
        Case "price", "nav"
            ' "현재가" 글자 대신 HTML 클래스 구조(no_today)를 찾아 숫자 추출
            re.Pattern = "no_today[\s\S]*?blind"">([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetNaverFinanceData = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetNaverFinanceData = 0 ' 에러 메시지 대신 0 반환 (형식 오류 방지)
            End If
            
        Case "change"
            ' 전일대비 금액 추출 (no_exday 클래스 내부 blind 태그)
            re.Pattern = "no_exday[\s\S]*?blind"">([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                Dim val As Double
                val = CDbl(Replace(m(0).SubMatches(0), ",", ""))
                ' 하락/상승 화살표 판단 (ico down이 있으면 마이너스 처리)
                If InStr(html, "ico down") > 0 Then val = val * -1
                GetNaverFinanceData = val
            Else
                GetNaverFinanceData = 0
            End If
            
        Case "change_pct"
            ' 등락률 추출
            re.Pattern = "n_chg[\s\S]*?blind"">([+-]?[\d\.]+)%"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetNaverFinanceData = CDbl(m(0).SubMatches(0)) / 100 ' 엑셀 백분율 형식 대응
            Else
                GetNaverFinanceData = 0
            End If
            
        Case Else
            GetNaverFinanceData = "Error: Invalid Field"
    End Select

    Exit Function
ErrHandler:
    GetNaverFinanceData = "Error: " & Err.Description
End Function

' ============================================================
' 인코딩 변환 함수 (안쓰면 계속 오류 떠잉)
' ============================================================
Private Function ConvertEucKrToUtf8(ByVal binaryBody As Variant) As String
    Dim adoStream As Object
    On Error Resume Next
    Set adoStream = CreateObject("ADODB.Stream")
    If adoStream Is Nothing Then
        ConvertEucKrToUtf8 = ""
        Exit Function
    End If
    
    adoStream.Type = 1 ' adTypeBinary
    adoStream.Open
    adoStream.Write binaryBody
    adoStream.Position = 0
    adoStream.Type = 2 ' adTypeText
    adoStream.Charset = "euc-kr"
    
    ConvertEucKrToUtf8 = adoStream.ReadText
    adoStream.Close
End Function
