' ============================================================
' RISE 200위클리커버드콜
' ============================================================
Function GetRISE200Nav(Optional field As String = "price") As Variant
    GetRISE200Nav = GetNaverFinanceData("475720", field)
End Function
' ============================================================
' ACE KRX 금현물
' ============================================================
Function GetACEGoldNav(Optional field As String = "price") As Variant
    GetACEGoldNav = GetNaverFinanceData("411060", field)
End Function

' ============================================================
' 공통 파싱 함수 (네이버 증권 기반)
' ============================================================
Function GetNaverFinanceData(itemCode As String, Optional field As String = "price") As Variant
    Dim xml As Object, html As String, re As Object, m As Object
    Dim url As String: url = "https://finance.naver.com/item/main.naver?code=" & itemCode
    
    On Error GoTo ErrHandler
    Set xml = CreateObject("MSXML2.XMLHTTP")
    xml.Open "GET", url, False
    xml.send
    
    ' 바이너리 데이터 to 텍스트 변환 (EUC-KR)
    html = ConvertEucKrToUtf8(xml.responseBody)

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    Select Case LCase(Trim(field))
        Case "price", "nav"
            re.Pattern = "no_today[\s\S]*?blind"">([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetNaverFinanceData = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetNaverFinanceData = 0
            End If
            
        Case "change"
            re.Pattern = "no_exday[\s\S]*?blind"">([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                Dim val As Double
                val = CDbl(Replace(m(0).SubMatches(0), ",", ""))
                If InStr(html, "ico down") > 0 Then val = val * -1
                GetNaverFinanceData = val
            Else
                GetNaverFinanceData = 0
            End If
            
        Case "change_pct"
            re.Pattern = "n_chg[\s\S]*?blind"">([+-]?[\d\.]+)%"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetNaverFinanceData = CDbl(m(0).SubMatches(0)) / 100
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
