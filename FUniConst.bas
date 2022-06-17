Attribute VB_Name = "FUniConst"
Public Function UniConst(Text As eUnicodeConst) As String
'---------------------------------------------------------------------------------------
' Procedure : UniConst
' Author    : Adam Waller
' Date      : 7/7/2020
' Purpose   : Search for characters: https://emojipedia.org/
'           : Look up UTF-16 Decimal value(s) from the following site:
'           : http://www.fileformat.info/info/unicode/char/search.htm
'           : https://www.fileformat.info/info/unicode/char/search.htm?q=NAME OF EMOJI&preview=entity
'---------------------------------------------------------------------------------------
'
    Select Case Text
        Case LeftArrow:     UniConst = ChrW(8592)
        Case RightArrow:    UniConst = ChrW(8594)
        Case Clock2:        UniConst = ChrW(55357) & ChrW(56657)
        Case SailBoat:      UniConst = ChrW(9973)
        Case PeaceLove:     UniConst = ChrW(9774)
        Case CheckOK:       UniConst = ChrW(9989)
        Case Ghost:         UniConst = ChrW(128123)
        Case MapleLeaf:     UniConst = ChrW(127809)
        Case Calaca:        UniConst = ChrW(128128)
        Case Corazon:       UniConst = ChrW(129505)
        Case Finger1:       UniConst = ChrW(128070)
        Case HandBro:       UniConst = ChrW(128587)
        Case Airplane:      UniConst = ChrW(9992)
        Case XMark:         UniConst = ChrW(917592)
        Case Umbrella:      UniConst = ChrW(127746)
        Case DogFace:       UniConst = ChrW(128054)
    End Select
End Function
