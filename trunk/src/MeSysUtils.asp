<SCRIPT Runat="Server" Language="VBScript">

' the MeASP System Utilites Functions

' the constants for GetPathType function
Const cPhysicPathType = 1
Const cAbsolutePathType = 2
Const cRelatedPathType = 3
Const cURLPathType = 4

'the constants for getScriptType
Const stVBScriptType = 0
Const stJScriptType  = 1

Public Sub RaiseError(ByRef pNumber, ByRef pSource, ByRef pDesc)
  if gRaiseException then
    'Err.Clear
    Call Err.Raise(pNumber, pSource, pDesc)
  else
    gErrorMessage = gErrorMessage & "<me:err id="""& Str(pNumber) &""" src="""& pSource &""">"& pDesc & "</me:err>"
  end if
End Sub

'wrap the response.Write
Public Sub Write(ByRef pMsg)
  response.Write(pMsg)
End Sub

Public Sub WriteLn(ByRef pMsg)
  response.Write(pMsg & "<br/>")
End Sub

Public Function GetClientIp()
  GetClientIp = Request.ServerVariables("REMOTE_ADDR")
End Function

Public Function GetClientHost()
  GetClientHost = Request.ServerVariables("REMOTE_HOST")
End Function

' find the file in the pFolder and child folders
' Note: the pFolder MUST is a Folder Object in FileSystemObject
' and the pFileName must LCase(pFileName) before call
' return the file object if success, or return Nothing object.
Public Function FindFile(ByRef pFolder, ByRef pFileName)
  Dim vSubItem, vResult
  Set vResult = Nothing
  For Each vSubItem In pFolder.Files
    if LCase(vSubItem.Name) = pFileName then Set FindFile = vSubItem: Exit Function
  Next

  For Each vSubItem In pFolder.SubFolders
    Set vResult = FindFile(vSubItem, pFileName)
    if not (vResult is Nothing) then Exit For
  Next
  Set FindFile = vResult
End Function

Public Function FolderExists(ByRef pFolderName)
  Dim vFSO
  Set vFSO = Server.CreateObject(FileSystemObjectName)
  FolderExists = vFSO.FolderExists(GetPhysicPath(pFolderName))
End Function

' Create all the subfolders if they do not exist yet.
Function ForceDirectories(ByVal strDir)
  Dim i, s, vDir, fso, DirList, vBaseFolder
  Set fso = CreateObject(FileSystemObjectName)
  vBaseFolder = server.mappath(".")
  strDir = Replace(strDir, "/", "\")
  i = Instr(strDir, vBaseFolder)
  'writeln("vBaseFolder=" & vBaseFolder)
  'writeln("strDir=" & strDir)
  if i > 0 then 
    strDir = Mid(strDir, i+ Len(vBaseFolder))
  elseif InStr(strDir, ":") = 2 then
    s = server.mappath("/")
    i = Instr(strDir, s)
    if i > 0 then 
      vBaseFolder = s
      strDir = Mid(strDir, i+ Len(vBaseFolder))
    else
      vBaseFolder = Left(strDir, 2)
      strDir = Mid(strDir, 3)
    end if
  end if
  'response.write "<br>strDir=" & strDir
  'response.write "<br>vBaseFolder=" & vBaseFolder
  DirList = Split(strDir, "\")
  s = vBaseFolder + DirList(0)
  For i = 1 to UBound(DirList)
   vDir = Trim(DirList(i))
   if vDir <> "" then
     s = s + "\" + DirList(i)
     'response.write "<br>sDir=" & s
     'response.write "<br>DirExists=" & fso.FolderExists(s)
     if not fso.FolderExists(s) then fso.CreateFolder(s)
   end if
  Next
  Set fso = nothing
End Function

Public Function ExtractFileName(ByVal aFileName)
  Dim i

  i = InStrRev("\", aFileName)
  if i > 0 then aFileName = Mid(aFileName, i+1)

  ' prevent hacker from using the "../" to enter other folder.
  i = InStrRev("/", aFileName)
  if i > 0 then aFileName = Mid(aFileName, i+1)
  ExtractFileName = aFileName
End Function

Public Function ExtractFilePath(ByVal aFileName)
  Dim i

  i = InStrRev("\", aFileName)
  if i > 0 then 
    aFileName = Left(aFileName, i)
  else
    i = InStrRev("/", aFileName)
    if i > 0 then aFileName = Left(aFileName, i)
  end if

  ExtractFilePath = aFileName
End Function

Public Function GetPathType(ByRef pPath)
    Dim vChar
    vChar = UCase((Mid(pPath,1,1)))
    vChar = Asc(vChar)
    if Len(pPath) >= 2 then
      if vChar >= Asc("A") and vChar <= Asc("Z") and Mid(pPath, 2, 1) = ":" then
        GetPathType = cPhysicPathType
        Exit Function
      end if
    end if

    if InStr(pPath, "http://") = 1 then
        GetPathType = cURLPathType
    elseif InStr(pPath, "https://") = 1 then
        GetPathType = cURLPathType
    elseif InStr(pPath, "ftp://") = 1 then
        GetPathType = cURLPathType
    elseif vChar = Asc("/") then
        GetPathType = cAbsolutePathType
    else
        GetPathType = cRelatedPathType
    end if
End Function

Public Function GetPhysicPath(ByVal aPath)
    Select Case GetPathType(aPath)
      Case cRelatedPathType
        'response.write "<br/>GetPhysicPath.aPath=" & aPath
        aPath = Server.MapPath(gMeSysPath + aPath)
      Case cAbsolutePathType
        'response.write "<br/>GetPhysicPath.aPath=" & aPath
        aPath = Server.MapPath(aPath)
      'Case cURLPathType
      Case Else
        aPath = ""
        ' Not Supports
    End Select
    GetPhysicPath = aPath
End Function

Public Function IncludeTrailingPathDelimiter(ByRef pStr)
  Dim vPathDelimiter
  
  if GetPathType(pStr) = cPhysicPathType then vPathDelimiter = "\" else vPathDelimiter = "/" 
  IncludeTrailingPathDelimiter = IncludeTrailingDelimiter(pStr, vPathDelimiter) 
End Function

Public Function IncludeTrailingDelimiter(ByRef pStr, ByRef pDelimiter)
  if pStr = "" then
    IncludeTrailingDelimiter=""
  else
    if Mid(pStr, Len(pStr), 1) = pDelimiter then IncludeTrailingDelimiter = pStr else IncludeTrailingDelimiter = pStr + pDelimiter
  end if
End Function

Public Function TrimChar(ByRef pStr, ByRef pChar)
  Dim Result
  Result = LTrimChar(pStr, pChar)
  Result = RTrimChar(Result, pChar)
  TrimChar = Result
End Function

Public Function LTrimChar(pStr, pChar)
  Dim i, vCount
  i = 1
  vCount = 0
  Do
    i = InStr(i, pStr, pChar, vbTextCompare)
    if i > 0 then
      if vCount = 0 and i <> 1 then Exit Do
      vCount = vCount + Len(PChar)
      i = i + Len(PChar) - 1
      'Debugger.Print "LTrimChar.i"&i, vCount
      if vCount <> i then vCount = vCount -Len(pChar):Exit Do
    end if
  Loop Until IsNull(i) or (i <= 0)
  if vCount > 0 then
    'Debugger.Print "LTrimChar.Count"&i, vCount
    LTrimChar = Mid(pStr, vCount+1)
  else
    LTrimChar = pStr
  end if
End Function

Public Function RTrimChar(ByRef pStr, ByRef pChar)
  Dim i, vCount
  i = -1
  vCount = -1
  'Debugger.Print "RTrimChar.Len", Len(pStr)
  'Debugger.Print "RTrimChar.Len(Char)", Len(pChar)
  Do
    i = InStrRev(pStr, pChar, i, vbTextCompare)
    if i > 0 then
      i = i - 1
      if vCount = -1 then
        if i <> Len(pStr) - Len(pChar) then Exit Do
        vCount = Len(pStr)
      end if
      vCount = vCount - Len(pChar) 
      'Debugger.Print "RTrimChar.i"&i, vCount
      if vCount <> i then vCount = vCount + Len(pChar):Exit Do
    end if
  Loop Until IsNull(i) or (i <= 0)
  if vCount > 0 and vCount < Len(pStr) then
    RTrimChar = Left(pStr, vCount)
  else
    RTrimChar = pStr
  end if
End Function

' the bit test function
' Is the pBit in the pValue?
Public Function IsBitIn(ByRef pValue, ByRef pBit)
  IsBitIn = ((pValue and pBit) = pBit)
End Function

Public Sub FillValue(ByRef pX, ByRef pValue)
  Dim i, vItem
  if IsArray(pX) or IsObject(pX) then
    For Each vItem In pX
      if IsObject(pValue) then
        Set vItem = pValue
      else
        vItem = pValue
      end if
    Next
  end if
End Sub

Function QuotedString(ByRef aStr, ByVal aQuote)
  Dim Result
  if IsEmpty(aQuote) or (aQuote = "") then aQuote = """"
  Result = Replace(aStr, aQuote, aQuote+aQuote)
  Result = aQuote + Result + aQuote
  QuotedString = Result
End Function

Function QuoteXml(ByRef pText)
    QuoteXml = Replace(pText, "&", "&amp;")
    QuoteXml = Replace(QuoteXml, "<", "&lt;")
    QuoteXml = Replace(QuoteXml, ">", "&gt;")
End Function

Function CDATAEncode(ByRef pText)
    If pText <> "" Then
        CDATAEncode = Replace(pText, "&", "&amp;")
        CDATAEncode = Replace(CDATAEncode, "<", "&lt;")
        CDATAEncode = Replace(CDATAEncode, "'", "&apos;")
    End If
End Function

Function PCDATAEncode(ByRef pText)
    If pText <> "" Then
        PCDATAEncode = Replace(pText, "&", "&amp;")
        PCDATAEncode = Replace(PCDATAEncode, "<", "&lt;")
        PCDATAEncode = Replace(PCDATAEncode, "]]>", "]]&gt;")
    End If
End Function

Function URLDecode(ByRef pURL)
    Dim vPos
    If pURL <> "" Then
        pURL = Replace(pURL, "+", " ")
        vPos = InStr(pURL, "%")
        Do While vPos > 0
            pURL = Left(pURL, vPos - 1) _
                 & Chr(CLng("&H" & Mid(pURL, vPos + 1, 2))) _
                 & Mid(pURL, vPos + 3)
            vPos = InStr(vPos + 1, pURL, "%")
        Loop
    End If
    URLDecode = pURL
End Function

Function URLEncode(ByRef pURL) 
  URLEncode = Server.URLEncode(pURL) 
End Function 

Function HTMLEncode(sText)
  HTMLEncode = Server.HTMLEncode(sText)
End Function 

Function HTMLDecode(sText)
    Dim I
    sText = Replace(sText, "&quot;", Chr(34))
    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))
    For I = 1 to 255
        sText = Replace(sText, "&#" & I & ";", Chr(I))
    Next
    HTMLDecode = sText
End Function

Function StringEncode(ByRef pText)
  Dim i, Result, vChar
  For i=1 To LenB(S)
    vChar = AscB(MidB(S,i,1))
    'if vChar < 0 then vChar = vChar + 65536
    if vChar = AscB("#") then
      Result = Result & ChrB(vChar)& ChrB(vChar)
    elseif vChar >= &H20 and vChar < &HFF then
      Result = Result & ChrB(vChar)
    else
      vChar = Hex(vChar)
      Result = Result & "#" & vChar &"#"
    end if
  Next
  StringEncode = Result
End Function

' TODO: CAN NOT work On Binary string
Function StringDecode(ByRef pText)
  Dim i, Result, vChar, vN
  i = 1
  Do
    vChar = Mid(S,i,1)
    if vChar = "#" then
        i = i + 1
        vN = ""
        vChar = Mid(s,i,1)
        Do While vChar <> "#"
          vN = vN + vChar
          i = i + 1
          if i > Len(pText) then Exit Do
          vChar = Mid(s,i,1)
        Loop 
        writeln i&":N:"+vN
        if vN <> "" then
          vN = "&H" + vN
          if Len(vN) = 4 then
            Result = Result & ChrB(vN)
          else
            Result = Result & Chr(vN)
          end if
        else
          Result = Result & "#"
        end if
    else
      Result = Result & vChar
    end if
    i = i + 1
  Loop Until i > Len(pText)
  StringDecode = Result
End Function

Function IsClientCookieSupported()
    Dim Result
    Result = Request.ServerVariables("HTTP_COOKIE")
    IsClientCookieSupported = (Result <> "") and (Len(Result) >= 2)
End Function

Function TryCreateObject(ByRef pClassName)
    Dim Result
    On Error Resume Next
    Set Result = Eval("New "+ pClassName)
    On Error Goto 0
    If not IsObject(Result) then
      Set Result = Nothing
    End if
    Set TryCreateObject = Result
End Function

Function TryCreateServerObject(ByRef pClassName)
    Dim Result
    On Error Resume Next
    Set Result = Server.CreateObject(pClassName)
    On Error Goto 0
    If not IsObject(Result) then
      Set Result = Nothing
    End if
    Set TryCreateServerObject = Result
End Function

Function GetScriptType(ByRef pStr)
    if InStr(1, pStr, "JSCRIPT", vbTextCompare) > 0 then
      GetScriptType = stJScriptType
    else
      GetScriptType = stVBScriptType
    end if
End Function

Function ExecuteVBScript(ByRef pLibName, ByRef pScript)
    Dim vNumber, vSource, vDesc
    ExecuteVBScript = True
    On Error Resume Next
    ExecuteGlobal(pScript)
    vNumber = Err.Number
    if vNumber <> 0 then
      vSource = Err.Source
      vDesc = Err.Description
      'Err.Clear 当调用 On Error Goto 0 就已经清理了。
    end if
    On Error Goto 0
    if  vNumber <> 0 then
      ExecuteVBScript = False
      Call RaiseError(vbExecuteLibScriptError, "TMeLib.Compiling the VB library:"""+pLibName+"""", vSource& "[" & vNumber & "]: " & vDesc)
    end if
End Function

Function ExecuteJScript(ByRef pLibName, ByRef pScript)
    Dim vNumber, vSource, vDesc
    ExecuteJScript = True
    On Error Resume Next
    jsEval(pScript)
    vNumber = Err.Number
    if vNumber <> 0 then
      vSource = Err.Source
      vDesc = Err.Description
      'Err.Clear 当调用 On Error Goto 0 就已经清理了。
    end if
    On Error Goto 0
    if  vNumber <> 0 then
      ExecuteJScript = False
      Call RaiseError(vbExecuteLibScriptError, "TMeLib.Compiling the JS library:"""+pLibName+"""", vSource& "[" & vNumber & "]: " & vDesc)
    end if
End Function

Function ExecuteScript(ByRef pLibName, ByRef pScript, ByRef pScriptType)
  Select Case pScriptType
    Case stJScriptType
      ExecuteScript = ExecuteJScript(pLibName, pScript)
    Case Else
      ExecuteScript = ExecuteVBScript(pLibName, pScript)
    'Case stVBScriptType
  End Select
End Function

</SCRIPT> 
