<SCRIPT Runat="Server" Language="VBScript">

Class TRC4Cipher

    Private FSBox(255)
    Private Fkeys(255)
    Private FPassword


    Private Sub Class_Initialize()
      'Password = "MeASPDefault"
      FPassword  = ""
    End Sub

    Private Sub Class_Terminate()
      FPassword  = ""
    End Sub

    Public Property Let Password(ByRef pValue)
        if FPassword <> pValue then
          FPassword = pValue
          if FPassword <> "" then RC4Initialize(FPassword)
        end if
    End Property

    Public Function Encrypt(ByRef pPlainText)
        Encrypt = iRC4CryptText(pPlainText)
    End Function

    Public Function Decrypt(ByRef pPlainText)
        Decrypt = iRC4CryptText(pPlainText)
    End Function

    ' return the encrypted or decrypted plain bin array: not test yet
    Public Function iRC4Crypt(ByRef pPlainArray)
       Dim temp
       Dim a
       Dim i
       Dim j
       Dim k
       Dim Result
    
       i = 0
       j = 0
    
       ReDim Result(UBound(pPlainArray) + 1)
       For a = 0 To UBound(pPlainArray)
          i = (i + 1) Mod 256
          j = (j + FSBox(i)) Mod 256
          temp = FSBox(i)
          FSBox(i) = FSBox(j)
          FSBox(j) = temp
    
          k = FSBox((FSBox(i) + FSBox(j)) Mod 256)
    
          Result(a) = pPlainArray(a) Xor k
          'response.write "<BR/>P["&a&"]="& Hex(pPlainText(a)) &"&nbsp;R="& Hex(Result(a))
       Next
    
       iRC4Crypt = Result
    End Function

    Public Function iRC4CryptText(ByRef pPlainText)
       Dim temp
       Dim a
       Dim i
       Dim j
       Dim k
       Dim vCryptedByte
       Dim Result
    
       i = 0
       j = 0
    
       Result = ""
       For a = 1 To LenB(pPlainText)
          i = (i + 1) Mod 256
          j = (j + FSBox(i)) Mod 256
          temp = FSBox(i)
          FSBox(i) = FSBox(j)
          FSBox(j) = temp
    
          k = FSBox((FSBox(i) + FSBox(j)) Mod 256)
    
          vCryptedByte = CByte(AscB(MidB(pPlainText, a, 1))) Xor k
          Result = Result + ChrB(vCryptedByte)
          'response.write "<BR/>P["&a&"]="& AscB(MidB(pPlainText, a, 1)) &"&nbsp;R="& Hex(AscB(MidB(Result, a, 1)) )
          'response.write "<BR/>P["&a&"]="& Hex(AscB(MidB(pPlainText, a, 1)) ) &"&nbsp;R="& Hex(AscB(MidB(Result, a, 1)) )
       Next
    
       iRC4CryptText = Result
    End Function

    Private Sub RC4Initialize(ByRef pPassword)
       Dim vTemp
       Dim vPwdLen
       Dim a, b
    
       vPwdLen = len(pPassword)
       For a = 0 To 255
          Fkeys(a) = ascB(midB(pPassword, (a mod vPwdLen)+1, 1))
          FSBox(a) = a
       next
    
       b = 0
       For a = 0 To 255
          b = (b + FSBox(a) + Fkeys(a)) Mod 256
          vTemp    = FSBox(a)
          FSBox(a) = FSBox(b)
          FSBox(b) = vTemp
       Next
    
    End Sub

End Class

   
Function RC4Encrypt(ByRef pPlainText, ByRef pPassword)
    Dim vCipher
    Set vCipher = New TRC4Cipher

    vCipher.Password = pPassword
    RC4Encrypt = vCipher.Encrypt(pPlainText)
    Set vCipher = Nothing

End Function

Function RC4Decrypt(ByRef pEncryptedText, ByRef pPassword)
    Dim vCipher
    Set vCipher = New TRC4Cipher

    vCipher.Password = pPassword
    RC4Decrypt = vCipher.Decrypt(pEncryptedText)
    Set vCipher = Nothing

End Function

Function RC4Crypt(ByRef pText, ByRef pPassword)
    Dim vCipher, Result
    Set vCipher = New TRC4Cipher

    vCipher.Password = pPassword
    Result = vCipher.iRC4CryptText(pText)
    Set vCipher = Nothing
    RC4Crypt = Result

End Function

Private Const CHARS_TO_A_WORD = 4
Private Const CHARS_TO_A_BYTE = 2

' the hex string like this: "0ABE3450AE00"
Function HexStrToWordArray(pHexStr)
  Dim Result, i, j
  Redim Result(Len(pHexStr) \ CHARS_TO_A_WORD)
  j = 0
  For i = 1 to ubound(Result) Step CHARS_TO_A_WORD
    'response.write "hexSrt="&Mid(pHexStr, i, CHARS_TO_A_WORD)&"&nbsp;"
    Result(j) = CInt(Mid(pHexStr, i, CHARS_TO_A_WORD))
    response.write "hex="&Result(j)
    j = j + 1
  Next
  HexStrToWordArray = Result
End Function

Function HexStrToByteArray(pHexStr)
  Dim Result, i, j
  Redim Result(Len(pHexStr) \ CHARS_TO_A_BYTE)
  j = 0
  For i = 1 to ubound(Result) Step CHARS_TO_A_BYTE
    'response.write "hexSrt="&Mid(pHexStr, i, CHARS_TO_A_BYTE)&"&nbsp;"
    Result(j) = CByte(Mid(pHexStr, i, CHARS_TO_A_BYTE))
    'response.write "hex="&Result(j)
    j = j + 1
  Next
  HexStrToByteArray = Result
End Function

Function StrToByteArray(pStr)
  Dim Result, i, c
  ReDim Result(Len(pStr)-1)
  
  For i = 1 to Len(pStr)
    Result(i-1) = AscB(pStr)
  Next

  StrToByteArray = Result
End Function

Function ByteArrayToStr(pByteArray)
  Dim Result, i, c
  
  Result = ""
  For i = 0 to UBound(pByteArray)
    Result = Result + Hex(pByteArray(i))
  Next

  ByteArrayToStr = Result
End Function
</SCRIPT>

