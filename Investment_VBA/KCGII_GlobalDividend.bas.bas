Attribute VB_Name = "KCGIINCOME"
Function GetKCGIGlobalDivNav(Optional field As String = "nav") As Variant

    ' KCGI글로벌고배당증권자투자신탁[주식-재간접형] A 클래스
    ' 표준코드: K55365D69567
    Dim url As String
    url = "https://www.funetf.co.kr/product/fund/view/K55365D69567"

    Dim xml As Object
    Set xml = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrHandler

    xml.Open "GET", url, False
    xml.setRequestHeader "User-Agent", "Mozilla/5.0"
    xml.setRequestHeader "Referer", "https://www.funetf.co.kr/"
    xml.send

    Dim html As String
    html = xml.responseText

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False

    Select Case LCase(Trim(field))
        Case "nav", ""
            re.Pattern = "기준가\(전일대비\)[\s\S]{1,200}?([\d,]+\.?\d*)원"
            Dim m As Object
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKCGIGlobalDivNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetKCGIGlobalDivNav = "NAV패턴불일치"
            End If
        Case "change"
            re.Pattern = "기준가\(전일대비\)[\s\S]{1,200}?[\d,]+\.?\d*원\s*[\s\S]{1,50}?([\+\-]?[\d,]+\.?\d*)\s*\("
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKCGIGlobalDivNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetKCGIGlobalDivNav = "등락패턴불일치"
            End If
        Case "change_pct"
            re.Pattern = "\(([\d\.]+)%\)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKCGIGlobalDivNav = CDbl(m(0).SubMatches(0))
            Else
                GetKCGIGlobalDivNav = "등락%패턴불일치"
            End If
        Case "date"
            re.Pattern = "(\d{2}\.\d{2}\.\d{2})\s*기준"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetKCGIGlobalDivNav = m(0).SubMatches(0)
            Else
                GetKCGIGlobalDivNav = "날짜패턴불일치"
            End If
        Case Else
            GetKCGIGlobalDivNav = "nav / change / change_pct / date 중 입력"
    End Select

    Exit Function
ErrHandler:
    GetKCGIGlobalDivNav = "오류:" & Err.Description

End Function
