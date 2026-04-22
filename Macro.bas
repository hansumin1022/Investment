Sub 새로고침()
'
' 새로고침 매크로
'
' 바로 가기 키: Ctrl+a
'
    Range("D4").Select
    ActiveCell.Formula2R1C1 = "=GetACEGoldNav(""nav"")"
    Range("D5").Select
    ActiveCell.Formula2R1C1 = "=GetKCGIGlobalBaedangNav()"
    Range("D6").Select
    ActiveCell.Formula2R1C1 = "=GetRISE200Nav()"
    Range("D7").Select
    ActiveCell.Formula2R1C1 = "=GetVietnamNav()"
    Range("D8").Select
    ActiveCell.Formula2R1C1 = "=GetMiraeIndiaNav()"
    Range("D9").Select
    ActiveCell.Formula2R1C1 = "=GetMiraeChinaNav()"
    Range("D10").Select
    ActiveCell.Formula2R1C1 = "=GetMiraeJapanNav()"
    Range("D11").Select
    ActiveCell.Formula2R1C1 = "=GetKCGIGlobalnav()"
    Range("D12").Select
    ActiveCell.Formula2R1C1 = "=GetKCGISmallCapAeNav()"
    Range("D13").Select
    ActiveCell.Formula2R1C1 = "=GetKCGIAINav()"
    Range("D14").Select
End Sub
