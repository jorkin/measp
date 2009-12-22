<%
'< S CRIPT Runat="Server" Language="VBScript">

' Conversion of Visual Basic code to VBScript code 
' VBScript code
' First published on <www.di-mgt.com.au/crypto.html> 2 February 2002
' Revised 13 August 2002: Fixed Mid() error in Base64ToAnsi


' basRadix64: Radix 64 en/decoding functions
' Version 3. Published January 2002 with even faster SHR/SHL functions
'            and using Mid$ function instead of appending to strings.
' Version 2. Published 12 May 2001
' Version 1. Published 28 December 2000
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-2 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact and acknowledge
' its authorship with the words:

'   "Contains cryptography software by David Ireland of
'   DI Management Services Pty Ltd <www.di-mgt.com.au>."

' This code may only be used as part of an application. It may
' not be reproduced or distributed separately by any means without
' the express written permission of the author.

' David Ireland and DI Management Services Pty Limited make no
' representations concerning either the merchantability of this
' software or the suitability of this software for any particular
' purpose. It is provided "as is" without express or implied
' warranty of any kind.

' Please forward comments or bug reports to <code@di-mgt.com.au>.
' The latest version of this source code can be downloaded from
' www.di-mgt.com.au/crypto.html.

' Credit where credit is due:
' Some parts of this VB code are based on original C code
' by Carl M. Ellison. See "cod64.c" published 1995.
'****************** END OF COPYRIGHT NOTICE*************************

Lib.Require("util.math")

Private cBase64DecodeTable(255)
Private Const cBase64EncodeTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'Private Const cBase64DelFlagChar = "="
Private Const cBase64DelFlagChar = "-" '×îÄ©Î²µÄÌîÁã²¹³ä×Ö·û

MakeDecodeTable()

'encode ansi string to base64 string
Public Function AnsiToBase64(sInput)
' Return radix64 encoding of string of binary values
' Does not insert CRLFs. Just returns one long string,
' so it's up to the user to add line breaks or other formatting.
' Version 3: Use Mid() function instead of appending
' VBScript: Doesn't work. Go back to appending
    Dim sOutput, sLast
    Dim b(2)
    Dim j
    Dim i, nLen, nQuants
    Dim iIndex

    nLen = LenB(sInput)
    nQuants = nLen \ 3
    ' sOutput = String(nQuants * 4, " ")
    sOutput = ""
    iIndex = 0
    ' Now start reading in 3 bytes at a time
    For i = 0 To nQuants - 1
        For j = 0 To 2
           b(j) = AscB(MidB(sInput, (i * 3) + j + 1, 1))
        Next
        ' Mid(sOutput, iIndex + 1, 4) = EncodeQuantum(b)
        sOutput = sOutput + EncodeQuantum(b)
        iIndex = iIndex + 4
    Next

    ' Cope with odd bytes
    Select Case nLen Mod 3
    Case 0
        sLast = ""
    Case 1
        b(0) = AscB(MidB(sInput, nLen, 1))
        b(1) = 0
        b(2) = 0
        sLast = EncodeQuantum(b)
        ' Replace last 2 with =
        sLast = Left(sLast, 2) + cBase64DelFlagChar + cBase64DelFlagChar
    Case 2
        b(0) = AscB(MidB(sInput, nLen - 1, 1))
        b(1) = AscB(MidB(sInput, nLen, 1))
        b(2) = 0
        sLast = EncodeQuantum(b)
        ' Replace last with =
        sLast = Left(sLast, 3) + cBase64DelFlagChar
    End Select

    AnsiToBase64 = sOutput + sLast
End Function

'decode base64 string to ansi string
Public Function Base64ToAnsi(sEncoded)
' Return string of decoded binary values given radix64 string
' Ignores any chars not in the 64-char subset
' Version 3: Use Mid) function instead of appending
' VBScript Revised 13 Aug 2002: Use appending instead of Mid()
' (VBScript doesn't seem to like Mid(str, i, 1) = "A")
    Dim sDecoded
    Dim d(3)
    Dim C
    Dim di
    Dim i
    Dim nLen
    Dim iIndex

    nLen = Len(sEncoded)
    'sDecoded = String((nLen \ 4) * 3, " ")
    sDecoded = ""
    iIndex = 0
    di = 0
    'Call MakeDecodeTable
    ' Read in each char in trun
    For i = 1 To Len(sEncoded)
        C = CByte(Asc(Mid(sEncoded, i, 1)))
        C = cBase64DecodeTable(C)
        If C >= 0 Then
            d(di) = C
            di = di + 1
            If di = 4 Then
                'Mid(sDecoded, iIndex + 1, 3) = DecodeQuantum(d)
                sDecoded = sDecoded & DecodeQuantum(d)
                iIndex = iIndex + 3
                If d(3) = 64 Then
                    sDecoded = LeftB(sDecoded, LenB(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                If d(2) = 64 Then
                    sDecoded = LeftB(sDecoded, LenB(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                di = 0
            End If
        End If
    Next

    Base64ToAnsi = sDecoded
End Function

Private Function EncodeQuantum(b())
    Dim sOutput
    Dim C

    sOutput = ""
    C = SHR2(b(0)) And &H3F
    sOutput = sOutput & Mid(cBase64EncodeTable, C + 1, 1)
    C = SHL4(b(0) And &H3) Or (SHR4(b(1)) And &HF)
    sOutput = sOutput & Mid(cBase64EncodeTable, C + 1, 1)
    C = SHL2(b(1) And &HF) Or (SHR6(b(2)) And &H3)
    sOutput = sOutput & Mid(cBase64EncodeTable, C + 1, 1)
    C = b(2) And &H3F
    sOutput = sOutput & Mid(cBase64EncodeTable, C + 1, 1)

    EncodeQuantum = sOutput

End Function

Private Function DecodeQuantum(d())
    Dim sOutput
    Dim C

    sOutput = ""
    C = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
    sOutput = sOutput + ChrB(C)
    C = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
    sOutput = sOutput + ChrB(C)
    C = SHL6(d(2) And &H3) Or d(3)
    sOutput = sOutput + ChrB(C)

    DecodeQuantum = sOutput

End Function

Private Function MakeDecodeTable()
' Set up Radix 64 decoding table
    Dim t
    Dim C

    For C = 0 To 255
        cBase64DecodeTable(C) = -1
    Next

    t = 0
    For C = Asc("A") To Asc("Z")
        cBase64DecodeTable(C) = t
        t = t + 1
    Next

    For C = Asc("a") To Asc("z")
        cBase64DecodeTable(C) = t
        t = t + 1
    Next

    For C = Asc("0") To Asc("9")
        cBase64DecodeTable(C) = t
        t = t + 1
    Next

    C = Asc("+")
    cBase64DecodeTable(C) = t
    t = t + 1

    C = Asc("/")
    cBase64DecodeTable(C) = t
    t = t + 1

    C = Asc(cBase64DelFlagChar)    ' flag for the byte-deleting char
    cBase64DecodeTable(C) = t  ' should be 64

End Function



'< / SCRIPT>
%>
