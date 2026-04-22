Function GetRISE200Nav(Optional field As String = "nav") As Variant
    ' RISE 200 위클리 커버드콜 (종목코드: 475720)
    Dim url As String
    ' 네이버페이 증권 ETF 실시간 API 사용 (디자인 변경에 영향받지 않음)
    url = "https://finance.naver.com/api/sise/etfItemList.nhn"
    
    Dim xml As Object
    Set xml = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrHandler
    
    xml.Open "GET", url, False
    xml.setRequestHeader "User-Agent", "Mozilla/5.0"
    xml.send
    
    Dim html As String
    html = xml.responseText
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    
    ' 종목코드 475720에 해당하는 JSON 데이터 블록만 정확히 추출
    re.Pattern = "{""itemcode"":""475720""[^}]+}"
    Dim m As Object
    Set m = re.Execute(html)
    
    If m.Count = 0 Then
        GetRISE200Nav = "데이터 검색 실패"
        Exit Function
    End If
    
    Dim jsonBlock As String
    jsonBlock = m(0).Value
    
    Select Case LCase(Trim(field))
        Case "nav", ""
            ' 순자산가치(NAV) 추출
            re.Pattern = """nav"":([-\d\.]+)"
            Set m = re.Execute(jsonBlock)
            If m.Count > 0 Then GetRISE200Nav = CDbl(m(0).SubMatches(0)) Else GetRISE200Nav = "NAV 파싱 오류"
            
        Case "change"
            ' 전일대비 등락 추출
            re.Pattern = """changeVal"":([-\d\.]+)"
            Set m = re.Execute(jsonBlock)
            If m.Count > 0 Then GetRISE200Nav = CDbl(m(0).SubMatches(0)) Else GetRISE200Nav = "등락 파싱 오류"
            
        Case "change_pct"
            ' 등락률(%) 추출
            re.Pattern = """changeRate"":([-\d\.]+)"
            Set m = re.Execute(jsonBlock)
            If m.Count > 0 Then GetRISE200Nav = CDbl(m(0).SubMatches(0)) Else GetRISE200Nav = "등락% 파싱 오류"
            
        Case "date"
            ' API는 당일 실시간 데이터를 제공하므로 현재 날짜 반환
            GetRISE200Nav = Format(Date, "yy.mm.dd")
            
        Case Else
            GetRISE200Nav = "nav / change / change_pct / date 중 하나를 입력하세요."
    End Select
    
    Exit Function
ErrHandler:
    GetRISE200Nav = "오류:" & Err.Description
End Function
