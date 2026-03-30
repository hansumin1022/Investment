Attribute VB_Name = "KOR"
' ============================================================
' 피델리티 재팬증권자투자신탁(주식-재간접형) C-e
' ============================================================
Function GetFidelityJapanNav(Optional field As String = "nav") As Variant
    GetFidelityJapanNav = GetFunETFNav("KR5235AB4750", field)
End Function

' ============================================================
' 미래에셋 인도중소형포커스증권자투자신탁 1호(주식) C-e
' ============================================================
Function GetMiraeIndiaNav(Optional field As String = "nav") As Variant
    GetMiraeIndiaNav = GetFunETFNav("K55301B58619", field)
End Function

' ============================================================
' 미래에셋 차이나과창판증권투자신탁(주식) C-e
' ============================================================
Function GetMiraeChinaNav(Optional field As String = "nav") As Variant
    GetMiraeChinaNav = GetFunETFNav("K55301DA3336", field)
End Function

' ============================================================
' 공통 파싱 함수 (FunETF 기반 - 모든 펀드 공유)
' ============================================================
Function GetFunETFNav(standardCd As String, Optional field As String = "nav") As Variant

    Dim xml As Object
    Set xml = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrHandler

    xml.Open "GET", "https://www.funetf.co.kr/product/fund/view/" & standardCd, False
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
                GetFunETFNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetFunETFNav = "NAV패턴불일치"
            End If
        Case "change"
            re.Pattern = "기준가\(전일대비\)[\s\S]{1,200}?[\d,]+\.?\d*원\s*[\s\S]{1,50}?([\+\-]?[\d,]+\.?\d*)\s*\("
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetFunETFNav = CDbl(Replace(m(0).SubMatches(0), ",", ""))
            Else
                GetFunETFNav = "등락패턴불일치"
            End If
        Case "change_pct"
            re.Pattern = "\(([\d\.]+)%\)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetFunETFNav = CDbl(m(0).SubMatches(0))
            Else
                GetFunETFNav = "등락%패턴불일치"
            End If
        Case "date"
            re.Pattern = "(\d{2}\.\d{2}\.\d{2})\s*기준"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetFunETFNav = m(0).SubMatches(0)
            Else
                GetFunETFNav = "날짜패턴불일치"
            End If
        Case Else
            GetFunETFNav = "nav / change / change_pct / date 중 입력"
    End Select

    Exit Function
ErrHandler:
    GetFunETFNav = "오류:" & Err.Description
End Function
