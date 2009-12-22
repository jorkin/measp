<%
'< S CRIPT Runat="Server" Language="VBScript">
'
' ---------------------------------------------------------------------------
'      $Source: /home/cvs/MeCMS/src/util/RegExp.lib.asp,v $
'      $Revision: 1.3 $
'      $Author: riceball $
' ---------------------------------------------------------------------------
'
' These functions simulate the m and s operations as available in the
' programming language perl. You can usually literally copy perl regular
' expressions and expect them to work with these functions.
'
' In perl you can do something like:
'
'     s/A(.*?)B(.*?)C/&MyMethod($1, $2)/ge
'
' When the match is made this will call the sub "MyMethod", pass it
' the two matched variables, and finally the match is substituted
' by whatever the sub returns.
'
' The function s below can behave in a similar manner. The perl expression
' shown above would be written as:
'
'     myText = gRegExp.s(myText, "A(.*?)B(.*?)C", "&MyMethod($1, $2)", True, True)
'
' In ASP the function must return the value to be substituted. E.g.
'
'     Function MyMethod(pParam1, pParam2)
'         If pParam1 = pParam2 Then
'             MyMethod = "Same"
'         Else
'             MyMethod = "Different"
'         End If
'     End Function
'

Public gRegExp

Set gRegExp = New TMeRegExp

Class TMeRegExp
    ' Reuse regular expression object
    Private FRegEx

    Private Sub Class_Initialize
      On Error Resume Next
      Set FRegEx = New RegExp
      On Error Goto 0
      FRegEx.MultiLine = True
      
      If Not IsObject(FRegEx) Then
          Call RaiseError(vbRegistryPermissionError, "RegExp", "<h2>Error:</h2><p>Probable cause: Registry permission problem.</p>"_
            & "This is a known problem with Microsoft.<br />" _
            & "You can find more information about this problem in the following  " _
            & "<a href=""http://support.microsoft.com/support/kb/articles/Q274/0/38.ASP"">Microsoft knowledge base article</a>."_
          )
      End If
    End Sub

    Private Sub Class_Terminate
      Set FRegEx = Nothing
    End Sub

    Public Property Get RegEx()
        Set RegEx = FRegEx
    End Property

    ' Get the Matched Result.
    Public Function GetMatches(ByRef pText, ByRef pSearchPattern, ByRef pIgnoreCase, aMultiSearch)
        If IsNull(pText) Then
            Exit Function
        End If
        FRegEx.IgnoreCase = pIgnoreCase
        FRegEx.Global     = aMultiSearch
        FRegEx.Pattern    = pSearchPattern
        Set GetMatches = FRegEx.Execute(pText)
    End Function

    Public Function g(ByRef pText, ByRef pSearchPattern, ByRef pIgnoreCase)
      Dim vMatches, vMatch
        g = ""
        If IsNull(pText) Then
            Exit Function
        End If
        Set vMatches = GetMatches(pText, pSearchPattern, pIgnoreCase, False) 'False means search once. 只找一次
        if not IsNull(vMatches) then
          For Each vMatch in vMatches
            'if g = "" then g = vMatch.Value else g = g + "," + vMatch.Value
            g = vMatch.Value 'Get the matched value.
          Next
        end if
    End Function
    
    'test whether the pText is match the pPattern or not
    Public Function m(ByRef pText, ByRef pPattern, ByRef pIgnoreCase, ByRef pGlobal)
        If IsNull(pText) Then
            m = False
            Exit Function
        End If
        FRegEx.IgnoreCase = pIgnoreCase
        FRegEx.Global     = pGlobal
        FRegEx.Pattern    = pPattern
        m = FRegEx.Test(pText)
    End Function
    
    
    'substitute, supports the function callback like perl s function:
    '  myText = s(myText, "A(.*?)B(.*?)C", "&MyMethod($1, $2)", True, True)
    'of cause the common substitute is support too:
    '  s(pText, "\&lt;br(\s[^<>/]+?)?\&gt;", "<br $1 />", True, True)
    '注意如果使用函数回调，那么aReplacePattern必须为函数！
    Public Function s(ByRef pText, ByRef pSearchPattern, ByVal aReplacePattern, ByRef pIgnoreCase, ByRef pGlobal)
        'Response.Write("<br /><br />Text: " & Server.HTMLEncode(pText))
        'Response.Write("<br />Patterns: " & Server.HTMLEncode(pSearchPattern) & " --> " & Server.HTMLEncode(aReplacePattern))
    
        If IsNull(pText) Then
            s = ""
            Exit Function
        End If
    
        FRegEx.IgnoreCase = pIgnoreCase
        FRegEx.Global     = pGlobal
        FRegEx.Pattern    = pSearchPattern
        If (Left(aReplacePattern, 1) <> "&") Then
            s = FRegEx.Replace(pText, aReplacePattern)
        Else 'it's function callback
            Dim vText, vPrevLastIndex, vPrevNewPos
            Dim vMatch, vMatches, vSubMatch, i, vCmd, vReplacement ', j, vStr
    
            vText          = pText
            vPrevLastIndex = 0
            vPrevNewPos    = 0
    
            aReplacePattern = Mid(aReplacePattern, 2)
    
            Set vMatches = FRegEx.Execute(pText)
            For Each vMatch In vMatches
                vCmd = aReplacePattern
    
                i = 0
                For Each vSubMatch in vMatch.SubMatches
                    'vStr = Trim(vSubMatch)
                    'vStr = Replace(vStr, Chr(0), "0")
                    'WriteLn("SubMatchA: " )
                    'For j = 1 to Len(vStr) 
                    'WriteLn("#"&Asc(Mid(vStr, j ,1)))
                    'Next
                    'WriteLn("SubMatch: [" & Server.HTMLEncode(vStr) & "]")
                    'vStr = Replace(vStr, """", """""")
                    'vCmd = Replace(vCmd, "$" & (i + 1), """" &  vStr & """")
                    vCmd = Replace(vCmd, "$" + CStr(i + 1), """" +  Replace(vSubMatch, """", """""") + """")
                    'Response.Write("<br />SubCmd: " & Server.HTMLEncode(vCmd))
                    i = i+ 1
                Next
    
                'WriteLn("REGEXP CMD: " & HTMLEncode(vCmd))
    
                vCmd = Replace(vCmd, vbCRLF, """ & vbCRLF & """)
                vReplacement = EVal(vCmd)
    
                ' replace vMatch.Value in vText by vReplacement
                vPrevNewPos = vPrevNewPos + (vMatch.FirstIndex - vPrevLastIndex)
                vText = Mid(vText, 1, vPrevNewPos) + CStr(vReplacement) + Mid(vText, vPrevNewPos + vMatch.Length + 1)
                vPrevNewPos = vPrevNewPos + Len(vReplacement) + 1
                vPrevLastIndex = vMatch.FirstIndex + vMatch.Length + 1
            Next
            s = vText
        End If
    End Function

    ' escape sequences to be interpreted literally by 'escaping' them by preceding them with a backslash "\", for instance: metacharacter "^" match beginning of string, but "\^" match character "^", "\\" match "\" and so on.
    ' for the text pattern
    Public Function EscapeString(ByRef pText)
        Dim Result
        Result = pText
        If pText="" Or IsNull(Result) Then
            Result=""
        Else
            Result=Replace(Result,"\","\\")
            Result=Replace(Result,"(","\(")
            Result=Replace(Result,")","\)")
            Result=Replace(Result,"*","\*")
            Result=Replace(Result,"?","\?")
            Result=Replace(Result,"{","\{")
            Result=Replace(Result,"}","\}")
            Result=Replace(Result,".","\.")
            Result=Replace(Result,"+","\+")
            Result=Replace(Result,"[","\[")
            Result=Replace(Result,"]","\]")
            Result=Replace(Result,"^","\^")
        End If
        EscapeString = Result
    End Function
End Class

'< / SCRIPT> 
%>