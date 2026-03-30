Attribute VB_Name = "dax"
Function GetKiwoomDAXNav(Optional field As String = "nav") As Variant

    Dim url As String
    url = "https://finance.naver.com/item/sise.naver?code=411860"

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
            ' 패턴: 현재가 18,115
            re.Pattern = "현재가 ([\d,]+)"
            Dim m As Object
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKiwoomDAXNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetKiwoomDAXNav = "현재가패턴불일치"
            End If

        Case "change"
            ' 패턴: 전일대비 하락/상승 70
            re.Pattern = "전일대비 [^\d]+([\d,]+)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKiwoomDAXNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetKiwoomDAXNav = "등락패턴불일치"
            End If

        Case "change_pct"
            ' 패턴: 0.38 퍼센트
            re.Pattern = "([\d\.]+) 퍼센트"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKiwoomDAXNav = CDbl(m(0).SubMatches(0))
            Else
                GetKiwoomDAXNav = "등락%패턴불일치"
            End If

        Case Else
            GetKiwoomDAXNav = "nav / change / change_pct 중 입력"

    End Select

    Exit Function
ErrHandler:
    GetKiwoomDAXNav = "오류:" & Err.Description

End Function
