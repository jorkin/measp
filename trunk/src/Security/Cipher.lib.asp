<SCRIPT Runat="Server" Language="VBScript">

Lib.Require("Security.Cipher.RC4")
'Lib.Require("util.string")

Private Const cDefaultPassword = "A2eDefaultSDfieUA03218e4fFJ4"

Public Function DefaultEncrypt(ByRef pPlainText, ByRef pPassword)
    DefaultEncrypt = RC4Encrypt(pPlainText, pPassword)
End Function

Public Function DefaultDecrypt(ByRef pEncryptedText, ByRef pPassword)
    DefaultDecrypt = RC4Decrypt(pEncryptedText, pPassword)
End Function

'=================================================================
'This function Encrypt/Decrypt Ansi(multi-bytes) text and returns the Ansi(multi-bytes) text
'-----------------------------------------------------------------
Function EnDeCryptXOR(ByVal aAnsiText, ByVal aPassword) 
  Dim Size, i,  KeySize, vChar, vPwdChar
  aPassword = CStr(aPassword)
  i = LenB(aPassword)
  if i > 0 then
    aPassword = aPassword + CStr(i)
  else
    aPassword = cDefaultPassword
  end if
  KeySize = LenB(aPassword)

  'aAnsiText = CStr(aAnsiText)
  Size = LenB(aAnsiText)

  For i = 1 To Size
    vChar = AscB(MidB(aAnsiText, i, 1))
    vPwdChar = AscB(MidB(aPassword, ((i mod KeySize) + 1), 1))
    EnDeCryptXOR = EnDeCryptXOR + ChrB(vChar Xor vPwdChar)
  Next 
End Function ' EnDeCryptXOR 

'Encrypt/Descrypt binary array and return the binary array.
Function EnDeCryptXORBinary(ByRef aBinary, ByVal aPassword) 
  Dim Result, i, j,  KeySize, vPwdByte
  aPassword = CStr(aPassword)
  i = LenB(aPassword)
  if i > 0 then
    aPassword = aPassword + CStr(i)
  else
    aPassword = cDefaultPassword
  end if
  KeySize = LenB(aPassword)

  if IsArray(aBinary) then
    Redim Result(UBound(aBinary) - LBound(aBinary))
    j = 0
    For i = LBound(aBinary) To UBound(aBinary)
      vPwdByte = AscB(MidB(aPassword, ((i mod KeySize) + 1), 1))
      Result(j) = aBinary(i) Xor vPwdByte
      j = j + 1
    Next 
  end if
  EnDeCryptXOR = Result
End Function ' EnDeCryptXORBinary 
'=================================================================

</SCRIPT>

