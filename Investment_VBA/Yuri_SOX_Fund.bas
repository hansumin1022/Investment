Attribute VB_Name = "semi"
Function GetYuriNav(Optional field As String = "nav") As Variant
    ' 유리필라델피아반도체인덱스UH C-e 클래스
    ' 표준코드: K55307D05969
    Dim url As String
    url = "https://www.funetf.co.kr/product/fund/view/K55307D05969"
    
    Dim xml As Object
    Set xml = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrHandler
    
    xml.Open "GET", url, False
    xml.setRequestHeader "User-Agent", "Mozilla/5.0"
    xml.setRequestHeader "Referer", "https://www.funetf.co.kr/"
    xml.send
    
    If xml.Status <> 200 Then
        GetYuriNav = "HTTP오류:" & xml.Status
        Exit Function
    End If
    
    Dim html As String
    html = xml.responseText
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    
    Select Case LCase(Trim(field))
        Case "nav", ""
            re.Pattern = "기준가\(전일대비\)[\s\S]{1,200}?([\d,]+\.?\d*)원"
            Dim m As Object
            Set m = re.Execute(html)
            If m.Count > 0 Then
                Dim navStr As String
                navStr = Replace(m(0).SubMatches(0), ",", "")
                GetYuriNav = CDbl(navStr)
            Else
                GetYuriNav = "NAV패턴불일치"
            End If
            
        Case "change"
            re.Pattern = "기준가\(전일대비\)[\s\S]{1,200}?[\d,]+\.?\d*원\s*[\s\S]{1,50}?([\+\-]?[\d,]+\.?\d*)\s*\("
            Set m = re.Execute(html)
            If m.Count > 0 Then
                Dim chgStr As String
                chgStr = Replace(m(0).SubMatches(0), ",", "")
                GetYuriNav = CDbl(chgStr)
            Else
                GetYuriNav = "등락패턴불일치"
            End If
            
        Case "change_pct"
            re.Pattern = "\(([\d\.]+)%\)"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetYuriNav = CDbl(m(0).SubMatches(0))
            Else
                GetYuriNav = "등락%패턴불일치"
            End If
            
        Case "date"
            re.Pattern = "(\d{2}\.\d{2}\.\d{2})\s*기준"
            Set m = re.Execute(html)
            If m.Count > 0 Then
                GetYuriNav = m(0).SubMatches(0)
            Else
                GetYuriNav = "날짜패턴불일치"
            End If
            
        Case Else
            GetYuriNav = "nav / change / change_pct / date 중 입력"
    End Select
    
    Exit Function
ErrHandler:
    GetYuriNav = "오류:" & Err.Description
End Function

