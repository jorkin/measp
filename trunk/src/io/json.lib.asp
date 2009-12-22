<SCRIPT Runat="Server" Language="VBScript">
' for Ajax convert vb object to java json object jscript.
Public Function ObjectToJsObject(ByRef pObject)
    Dim i, vProps, Result
    Result = ""
    vProps = pObject.GetMetaObject
    if IsArray(vProps) then
      For i = 1 to UBound(vProps)
         if Result <> "" then Result = Result + "," 
         Result = Result + vProps(i) + ":"
         Select Case Eval("VarType(pObject."+vProps(i)+")")
           Case vbString
             Result = Result  + """" + Eval("pObject."+vProps(i)) + """"
           Case vbDate
             'todo: i dont know whether the vbs return the date type as the ms number from 1970.1.1
             ' VBS 不是从1970年开始的！
             Result = Result  + "new Date(" + CStr(Eval("CSng(pObject."+vProps(i)+")") - CSng(#1970-1-1#)) + ")"
           Case vbArray
             Result = Result  + "[" + ArrayToString(Eval("pObject."+vProps(i))) + "]"
           Case vbObject
             Result = Result + ObjectToJsObject(Eval("pObject."+vProps(i)))
           Case Else 'vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
             Result = Result + CStr(Eval("pObject."+vProps(i)))
         End Select
      Next 'i
    end if
    if Result <> "" then Result = "{" + Result + "}"
    ObjectToJsObject = Result
End Function

Private Function ArrayToString(ByRef pArray)
    Dim i, vProps, Result
    Result = ""
    if IsArray(pArray) then
      For i = LBound(pArray) to UBound(pArray)
         if Result <> "" then Result = Result + "," 
         Select Case VarType(pArray(i))
           Case vbString
             Result = Result + """" + pArray(i) + """"
           Case vbDate
             Result = Result  + "new Date(" + CStr(CSng(pArray(i) - #1970-1-1#)) + ")"
           Case Else 'vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
             Result = Result + CStr(pArray(i))
         End Select
      Next 'i
    end if
    ArrayToString = Result
End Function

</SCRIPT>
