Attribute VB_Name = "CryptoHelpers"
'RC4 encrypts strText using strKey as the key.
Public Function RC4(strText As String, strKey As String) As String
    Dim intKeyLen As Integer, intTemp As Integer
    Dim intX1 As Integer, intX2 As Integer, intX3 As Integer
    Dim aintB(256) As Integer, aintK(256) As Integer

    'Initialise buffer and keystream
    intKeyLen = Len(strKey)
    For intX1 = 0 To 255
        aintB(intX1) = intX1
        aintK(intX1) = Asc(Mid(strKey, (intX1 Mod intKeyLen) + 1))
    Next

    'Scramble up the data in the buffer a bit
    intX2 = 0
    For intX1 = 0 To 255
        intX2 = (intX2 + aintB(intX1) + aintK(intX1)) Mod 255
        intTemp = aintB(intX1)
        aintB(intX1) = aintB(intX2)
        aintB(intX2) = intTemp
    Next

    'Encode/Decode (but Process n bytes through the stream before we start)
    intX2 = 0
    intX3 = 0
    For intX1 = 1 To 3072 + Len(strText)
        intX2 = (intX2 + 1) Mod 255
        intX3 = (intX3 + aintB(intX2)) Mod 255
        intTemp = aintB(intX2)
        aintB(intX2) = aintB(intX3)
        aintB(intX3) = intTemp
        If intX1 > 3072 Then
            RC4 = RC4 & _
                  Chr(Asc(Mid(strText, intX1 - 3072)) Xor _
                      aintB((aintB(intX2) + aintB(intX3)) Mod 255))
        End If
    Next
End Function
Sub Prueba()

    Dim Frase As String
    Dim Pw As String
    
    Dim Texto As String
    Dim cCifrado As clsCifrado
    Dim cCrypto As clsCryptoFilterBox
    Dim strCypher As String
    Dim strUnCypher As String
    
    Dim oDoc As DOMDocument
    Dim oRoot As IXMLDOMNode
    
    Set cCifrado = New clsCifrado
    Set cCrypto = New clsCryptoFilterBox

    
    Frase = "LaMamarron@00"
    Pw = "Monique00"
    
    cCrypto.Password = "Imagination is more important than knowledge-Albert Einstein"
    cCrypto.InBuffer = "M0n1ca"
    cCrypto.Encrypt
    strCypher = cCrypto.OutBuffer
    Debug.Print strCypher
    cCrypto.InBuffer = ""

    
    cCrypto.Password = "Imagination is more important than knowledge-Albert Einstein"
    cCrypto.InBuffer = strCypher
    cCrypto.Decrypt
    strUnCypher = cCrypto.OutBuffer
    Debug.Print strUnCypher
    '---poner la contraseña
    If Pw = "" Then
        MsgBox "The Password is missing"
        Exit Sub
    Else
        cCifrado.Clave = Pw
    End If

    '---Sacar los datos
    Texto = Frase

    '---cifrar el texto
    Texto = cCifrado.Cifrar(Texto)

    Frase = Texto

'Obtain the AddressEntry for CurrentUser
Set oExUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser

MsgBox oExUser.PrimarySmtpAddress
MsgBox oExUser.YomiCompanyName
MsgBox oExUser.Address
MsgBox oExUser.Department
  
End Sub