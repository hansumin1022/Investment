Attribute VB_Name = "acekrxgold"
Function GetACEGoldNav(Optional field As String = "nav") As Variant

    ' ACE KRX 금현물 ETF 종목코드: 411060
    Dim url As String
    url = "https://finance.naver.com/item/sise.naver?code=411060"

    Dim xml As Object
    Set xml = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrHandler

    xml.Open "GET", url, False
    xml.setRequestHeader "User-Agent", "Mozilla/5.0"
    xml.setRequestHeader "Referer", "https://finance.naver.com/"
    xml.send

    Dim html As String
    html = xml.responseText

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False

    Select Case LCase(Trim(field))
        Case "nav", ""
            re.Pattern = "현재가 ([\d,]+)"
            Dim m As Object
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetACEGoldNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetACEGoldNav = "현재가패턴불일치"
            End If
        Case "change"
            re.Pattern = "전일대비 [^\d]+([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetACEGoldNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetACEGoldNav = "등락패턴불일치"
            End If
        Case "change_pct"
            re.Pattern = "([\d\.]+) 퍼센트"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetACEGoldNav = CDbl(m(0).SubMatches(0))
            Else
                GetACEGoldNav = "등락%패턴불일치"
            End If
        Case Else
            GetACEGoldNav = "nav / change / change_pct 중 입력"
    End Select

    Exit Function
ErrHandler:
    GetACEGoldNav = "오류:" & Err.Description

End Function
