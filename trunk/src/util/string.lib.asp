<%
'<   S CRIPT Runat="Server" Language="VBScript">

Lib.Require("ADOConsts")

Private Const cDefaultCharset = "utf-8" ' "us-ascii"

' convert hex string to ansi string
Function HexToAnsi(ByRef aHexStr)
  Dim Result, i, vHex
  Result = ""
  For i = 1 to Len(aHexStr) Step 2
    vHex = Mid(aHexStr, i, 2)
    Result = Result + ChrB("&H"+vHex)
  Next 'i
  HexToAnsi = Result
End Function

' convert Ansi string to Hex string
Function AnsiToHex(ByRef aAnsiStr)
  Dim Result, i, vByte
  Result = ""
  For i = 1 to LenB(aAnsiStr)
    vByte = Hex(AscB(MidB(aAnsiStr, i, 1)))
    if Len(vByte) = 1 then vByte = "0"+vByte
      'Writeln(vByte)
    Result = Result + vByte
  Next 'i
  AnsiToHex = Result
End Function

Function StringToAnsi(ByRef aStr)
  Dim i, Result, vByte, vL, vH ', s
  Result = ""
  For i = 1 to Len(aStr)
    vByte = Asc(Mid(aStr,i,1))
    if vByte < 0 then vByte = vByte + 65536
    if vByte > 255 then
      vByte = Hex(vByte)
      vL = "&H"+Right(vByte, 2)
      vH = "&H"+Left(vByte, 2)
      'writeln vL + ":" + vH + ":" + vByte
      Result = Result + ChrB(vL) + ChrB(vH)
      's = Mid(aStr,i,1)
      'writeln s+Hex(AscB(MidB(s,1,1))) + Hex(AscB(MidB(s,2,1)))
    else
      Result = Result + ChrB(vByte)
    end if
  Next 'i
  StringToAnsi = Result
End Function

Function AnsiToString(ByRef aAnsiStr)
  Dim i, Result, vByte, vL, vH
  Result = ""
  i = 1
  Do While i <= LenB(aAnsiStr)
    vByte = AscB(MidB(aAnsiStr,i,1))
    if vByte > &H7E then
      i = i + 1
      'Writeln Hex(AscB(MidB(aAnsiStr,i,1))) + Hex(vByte)
      vByte = "&H" + Hex(AscB(MidB(aAnsiStr,i,1)))+Hex(vByte)
      'Writeln vByte
      Result = Result + Chr(vByte) 
    else
      Result = Result + Chr(vByte)
    end if
    i = i + 1
  Loop
  AnsiToString = Result
End Function


Function StringToBinary(ByRef Text)
  StringToBinary = CharsetStringToBinary(Text, "")
End Function

Function BinaryToString(ByRef Binary)
  BinaryToString = BinaryToCharsetString(Binary, "")
End Function

' the default charset is utf-8
' the binary array or ansi(ascii) string to charset string
Function BinaryToCharsetString(ByRef Binary, ByVal Charset)
  'Create Stream object
  Dim vBinaryStream 'As New Stream
  Set vBinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type
  if IsArray(Binary) then
    vBinaryStream.Type = adTypeBinary
  else
    vBinaryStream.Type = adTypeText
    'Specify charset For the source text (ansi) data.
    vBinaryStream.CharSet = "ascii"
  end if
  
  'Open the stream And write text/string data To the object
  vBinaryStream.Open
  if IsArray(Binary) then
    vBinaryStream.Write Binary
  else
    vBinaryStream.WriteText Binary
  end if
  
  
  'Change stream type To binary
  vBinaryStream.Position = 0
  vBinaryStream.Type = adTypeText

  If Len(CharSet) > 0 Then
    vBinaryStream.CharSet = CharSet
  Else
    vBinaryStream.CharSet = cDefaultCharset
  End If

  'Ignore first two bytes - sign of
  'vBinaryStream.Position = 0
  
  'Open the stream And get text from the object
  BinaryToCharsetString = vBinaryStream.ReadText
  vBinaryStream.Close
  Set vBinaryStream = Nothing
End Function

Function CharsetStringToBinary(ByRef Text, ByVal Charset)
  'Create Stream object
  Dim vBinaryStream 'As New Stream
  Set vBinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save text/string data.
  vBinaryStream.Type = adTypeText
  
  'Specify charset For the source text (unicode) data.
  If Len(CharSet) > 0 Then
    vBinaryStream.CharSet = CharSet
  Else
    vBinaryStream.CharSet = cDefaultCharset
  End If
  
  'Open the stream And write text/string data To the object
  vBinaryStream.Open
  vBinaryStream.WriteText Text
  
  
  'Change stream type To binary
  vBinaryStream.Position = 0
  vBinaryStream.Type = adTypeBinary
  
  'Ignore first two bytes - sign of
  vBinaryStream.Position = 0
  
  'Open the stream And get binary data from the object
  CharsetStringToBinary = vBinaryStream.Read
  vBinaryStream.Close
  Set vBinaryStream = Nothing
End Function

'SimpleBinaryToString is clear function, but the function takes much time for large data. 
'You can use it to convert data with up to 100kB of size (concatenation of large string takes much processor time).
Function SimpleBinaryToString(ByRef Binary)
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  SimpleBinaryToString = S
End Function

' This function is up to 20 times faster than SimpleBinaryToString. You can use it to convert up to 2 MB of binary data.
Function SimpleBinaryToStringEx(ByRef Binary)
  'Antonin Foller, http://www.motobit.com
  'Optimized version of a simple BinaryToString algorithm.
  
  Dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    If cl3>300 Then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      If cl2>200 Then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  SimpleBinaryToStringEx = pl1 & pl2 & pl3
End Function

' ADODB.Recordset lets you work with all supported VARIANT data types - also with binary and String data (VT_UI1 | VT_ARRAY, BSTR). 
' It lets you convert between these two data formats :
' RSBinaryToString is not limitted by time - only by physical memory. 
' The function is up to 100 times faster than MultiByte conversions - you can use it to convert up to 100MB data.
Function RSBinaryToString(xBinary)
  'Antonin Foller, http://www.motobit.com
  'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using ADO recordset
  Dim Binary
  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
  If vartype(xBinary)=vbString Then Binary = AnsiToBinary(xBinary) Else Binary = xBinary
  
  Dim RS, LBinary
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    Call RS.Fields.Append("mBinary", adLongVarChar, LBinary)
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function

' the ANSI(MultiByte) string to Binary array
Function AnsiToBinary(ByRef aMultiByte)
  '? 2000 Antonin Foller, http://www.motobit.com
  ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
  ' Using recordset
  Dim RS, LMultiByte, Binary
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(aMultiByte)
  If LMultiByte>0 Then
    Call RS.Fields.Append("mBinary", adLongVarBinary, LMultiByte)
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk aMultiByte & ChrB(0)
    RS.Update
    Binary = RS("mBinary").GetChunk(LMultiByte)
    RS.Close
  End If
  Set RS = Nothing
  AnsiToBinary = Binary
End Function

' TODO: Not test yet.
Function BinaryToAnsi(ByRef aBinary)
  Dim RS, LBinarySize, Result
  Set RS = CreateObject("ADODB.Recordset")
  LBinarySize = LenB(aBinary)
  If LBinarySize>0 Then
    Call RS.Fields.Append("mBinary", adLongVarChar, LBinarySize)
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk aBinary & ChrB(0)
    RS.Update
    Result = RS("mBinary").GetChunk(LBinarySize)
    RS.Close
  End If
  Set RS = Nothing
  BinaryToAnsi = Result
End Function

'VBScript QuotedPrintable encoding
'2005 Antonin Foller http://www.motobit.com
' SourceString - string variable with source data, BSTR type
' Charset - Charset of the destination data 
Function StringToQuotedPrintable(SourceString, CharSet)
  'Create CDO.Message object For the encoding.
  Dim Message: Set Message = CreateObject("CDO.Message")

  'Set the encoding
  Message.BodyPart.ContentTransferEncoding = "quoted-printable"
  
  'Get the data stream To write source string data
  Dim Stream 'As ADODB.Stream
  Set Stream = Message.BodyPart.GetDecodedContentStream
  
  'Set the charset For the destination data, If required
  If Len(CharSet) > 0 Then Stream.CharSet = CharSet
  
  'Write the VBScript string To the stream.
  Stream.WriteText SourceString
  
  'Store the data To the message BodyPart
  Stream.Flush
  
  'Get an encoded stream
  Set Stream = Message.BodyPart.GetEncodedContentStream
  
  'read the encoded data As a string
  StringToQuotedPrintable = Stream.ReadText
  
  'You can use Read method To get a binary data.
  'Stream.Type = 1
  'StringToQuotedPrintable = Stream.Read
End Function

'VBScript BinaryToQuotedPrintable encoding
'2005 Antonin Foller http://www.motobit.com
Function BinaryToQuotedPrintable(SourceBinary)
  'Create CDO.Message object For the encoding.
  Dim Message: Set Message = CreateObject("CDO.Message")

  'Set the encoding
  Message.BodyPart.ContentTransferEncoding = "quoted-printable"
  
  'Get the data stream To write source string data
  Dim Stream 'As ADODB.Stream
  Set Stream = Message.BodyPart.GetDecodedContentStream
  
  'Set the type of the stream To adTypeBinary.
  Stream.Type = 1
  'Write the VBScript string To the stream.
  Stream.Write SourceBinary
  
  'Store the data To the message BodyPart
  Stream.Flush
  
  'Get an encoded stream
  Set Stream = Message.BodyPart.GetEncodedContentStream
  
  'Set the type of the stream To adTypeBinary.
  Stream.Type = 1
  
  'You can use Read method To get a binary data.
  BinaryToQuotedPrintable = Stream.Read
End Function

'VBScript QuotedPrintableToString decoding Function
'2005 Antonin Foller http://www.motobit.com
Function QuotedPrintableToString(SourceData, CharSet)
  'Create CDO.Message object For the encoding.
  Dim Message: Set Message = CreateObject("CDO.Message")

  'Set the encoding
  Message.BodyPart.ContentTransferEncoding = "quoted-printable"
  
  'Get the data stream To write source string data
  Dim Stream 'As ADODB.Stream
  Set Stream = Message.BodyPart.GetEncodedContentStream
  If VarType(SourceData) = vbString Then
    'Write the VBScript string To the stream.
    Stream.WriteText SourceData
  Else
    'Set the type of the stream To adTypeBinary.
    Stream.Type = 1
    'Write the source binary data To the stream.
    Stream.Write SourceData
  End If
  
  'Store the data To the message BodyPart
  Stream.Flush
  
  'Get an encoded stream
  Set Stream = Message.BodyPart.GetDecodedContentStream
  
  'Set the type of the stream To adTypeBinary.
  Stream.CharSet = CharSet
  
  'You can use Read method To get a binary data.
  QuotedPrintableToString = Stream.ReadText
End Function

'VBScript QuotedPrintableToBinary decoding Function
'2005 Antonin Foller http://www.motobit.com
Function QuotedPrintableToBinary(SourceData)
  'Create CDO.Message object For the encoding.
  Dim Message: Set Message = CreateObject("CDO.Message")

  'Set the encoding
  Message.BodyPart.ContentTransferEncoding = "quoted-printable"
  
  'Get the data stream To write source string data
  Dim Stream 'As ADODB.Stream
  Set Stream = Message.BodyPart.GetEncodedContentStream
  If VarType(SourceData) = vbString Then
    'Write the VBScript string To the stream.
    Stream.WriteWritetext SourceData
  Else
    'Set the type of the stream To adTypeBinary.
    Stream.Type = 1
    'Write the source binary data To the stream.
    Stream.Write SourceData
  End If
  
  'Store the data To the message BodyPart
  Stream.Flush
  
  'Get an encoded stream
  Set Stream = Message.BodyPart.GetDecodedContentStream
  
  'Set the type of the stream To adTypeBinary.
  Stream.Type = 1
  
  'You can use Read method To get a binary data.
  QuotedPrintableToBinary = Stream.Read
End Function


'<  /S CRIPT>
%>