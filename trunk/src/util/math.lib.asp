<%
'< S CRIPT Runat="Server" Language="VBScript">

Public Const BITS_IN_BYTE = 8
Public Const BYTES_IN_WORD = 4
Public Const BITS_IN_WORD = 32

Public cOnBitsTable(30)
Public cPower2Table(30)

InitMathUnit()

Private Sub InitMathUnit()
    cOnBitsTable(0) = CLng(1)
    cOnBitsTable(1) = CLng(3)
    cOnBitsTable(2) = CLng(7)
    cOnBitsTable(3) = CLng(15)
    cOnBitsTable(4) = CLng(31)
    cOnBitsTable(5) = CLng(63)
    cOnBitsTable(6) = CLng(127)
    cOnBitsTable(7) = CLng(255)
    cOnBitsTable(8) = CLng(511)
    cOnBitsTable(9) = CLng(1023)
    cOnBitsTable(10) = CLng(2047)
    cOnBitsTable(11) = CLng(4095)
    cOnBitsTable(12) = CLng(8191)
    cOnBitsTable(13) = CLng(16383)
    cOnBitsTable(14) = CLng(32767)
    cOnBitsTable(15) = CLng(65535)
    cOnBitsTable(16) = CLng(131071)
    cOnBitsTable(17) = CLng(262143)
    cOnBitsTable(18) = CLng(524287)
    cOnBitsTable(19) = CLng(1048575)
    cOnBitsTable(20) = CLng(2097151)
    cOnBitsTable(21) = CLng(4194303)
    cOnBitsTable(22) = CLng(8388607)
    cOnBitsTable(23) = CLng(16777215)
    cOnBitsTable(24) = CLng(33554431)
    cOnBitsTable(25) = CLng(67108863)
    cOnBitsTable(26) = CLng(134217727)
    cOnBitsTable(27) = CLng(268435455)
    cOnBitsTable(28) = CLng(536870911)
    cOnBitsTable(29) = CLng(1073741823)
    cOnBitsTable(30) = CLng(2147483647)
    
    cPower2Table(0) = CLng(1)
    cPower2Table(1) = CLng(2)
    cPower2Table(2) = CLng(4)
    cPower2Table(3) = CLng(8)
    cPower2Table(4) = CLng(16)
    cPower2Table(5) = CLng(32)
    cPower2Table(6) = CLng(64)
    cPower2Table(7) = CLng(128)
    cPower2Table(8) = CLng(256)
    cPower2Table(9) = CLng(512)
    cPower2Table(10) = CLng(1024)
    cPower2Table(11) = CLng(2048)
    cPower2Table(12) = CLng(4096)
    cPower2Table(13) = CLng(8192)
    cPower2Table(14) = CLng(16384)
    cPower2Table(15) = CLng(32768)
    cPower2Table(16) = CLng(65536)
    cPower2Table(17) = CLng(131072)
    cPower2Table(18) = CLng(262144)
    cPower2Table(19) = CLng(524288)
    cPower2Table(20) = CLng(1048576)
    cPower2Table(21) = CLng(2097152)
    cPower2Table(22) = CLng(4194304)
    cPower2Table(23) = CLng(8388608)
    cPower2Table(24) = CLng(16777216)
    cPower2Table(25) = CLng(33554432)
    cPower2Table(26) = CLng(67108864)
    cPower2Table(27) = CLng(134217728)
    cPower2Table(28) = CLng(268435456)
    cPower2Table(29) = CLng(536870912)
    cPower2Table(30) = CLng(1073741824)
End Sub

Public Function LShift(lvalue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lvalue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lvalue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lvalue And cPower2Table(31 - iShiftBits)) Then
        LShift = ((lvalue And cOnBitsTable(31 - (iShiftBits + 1))) * cPower2Table(iShiftBits)) Or &H80000000
    Else
        LShift = ((lvalue And cOnBitsTable(31 - iShiftBits)) * cPower2Table(iShiftBits))
    End If
End Function

Public Function RShift(lvalue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lvalue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lvalue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lvalue And &H7FFFFFFE) \ cPower2Table(iShiftBits)

    If (lvalue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ cPower2Table(iShiftBits - 1)))
    End If
End Function

Public Function RotateLeft(lvalue, iShiftBits)
    RotateLeft = LShift(lvalue, iShiftBits) Or RShift(lvalue, (32 - iShiftBits))
End Function


' Version 3: ShiftLeft and ShiftRight functions improved.
Public Function SHL2(ByVal bytValue)
' Shift 8-bit value to left by 2 bits
' i.e. VB equivalent of "bytValue << 2" in C
    SHL2 = (bytValue * &H4) And &HFF
End Function

Public Function SHL4(ByVal bytValue)
' Shift 8-bit value to left by 4 bits
' i.e. VB equivalent of "bytValue << 4" in C
    SHL4 = (bytValue * &H10) And &HFF
End Function

Public Function SHL6(ByVal bytValue)
' Shift 8-bit value to left by 6 bits
' i.e. VB equivalent of "bytValue << 6" in C
    SHL6 = (bytValue * &H40) And &HFF
End Function

Public Function SHR2(ByVal bytValue)
' Shift 8-bit value to right by 2 bits
' i.e. VB equivalent of "bytValue >> 2" in C
    SHR2 = bytValue \ &H4
End Function

Public Function SHR4(ByVal bytValue)
' Shift 8-bit value to right by 4 bits
' i.e. VB equivalent of "bytValue >> 4" in C
    SHR4 = bytValue \ &H10
End Function

Public Function SHR6(ByVal bytValue)
' Shift 8-bit value to right by 6 bits
' i.e. VB equivalent of "bytValue >> 6" in C
    SHR6 = bytValue \ &H40
End Function

'< / SCRIPT>
%>