<SCRIPT Runat="Server" Language="VBScript">

'require the MeConsts.asp
Class TMeDebugger
    Dim FEnabled
    Dim FRequestTime
    Dim FFinishTime
    Dim FObjStorage

    Public Property Get Enabled()
        Enabled = FEnabled
    End Property

    Public Property Let Enabled(pValue)
        FEnabled = pValue
    End Property

    Private Sub Class_Initialize()
        FRequestTime = Timer()
        Set FObjStorage = Server.CreateObject(DictionaryObjectName)
    End Sub

    Public Sub Print(label, output)
        If Enabled then
            FObjStorage.Add label, output
        End if
    End Sub

    Public Sub [End]()
        FFinishTime = Timer()
        If Enabled then
            PrintSummaryInfo()
            Call PrintCollection("VARIABLE STORAGE", FObjStorage)
            Call PrintCollection("QUERYSTRING COLLECTION", Request.QueryString())
            Call PrintCollection("FORM COLLECTION", Request.Form())
            Call PrintCollection("COOKIES COLLECTION", Request.Cookies())
            Call PrintCollection("SERVER VARIABLES COLLECTION", Request.ServerVariables())
            Call PrintCollection("APPLICATION CONTENTS COLLECTION", Application.Contents())
            Call PrintCollection("APPLICATION STATICOBJECTS COLLECTION", Application.StaticObjects())
            On Error Resume Next
            Call PrintCollection("SESSION CONTENTS COLLECTION", Session.Contents())
            Call PrintCollection("SESSION STATICOBJECTS COLLECTION", Session.StaticObjects())
            On Error Goto 0
        End if
    End Sub

    Private Sub PrintSummaryInfo()
        With Response
            .Write("<hr>")
            .Write("<b>SUMMARY INFO</b></br>")
            '.Write("Time of Request = " & FRequestTime) & "<br>"
            '.Write("Time Finished = " & FFinishTime) & "<br>"
            .Write("Elapsed Time = " & (FFinishTime - FRequestTime) * 1000 & " ms<br>")
            .Write("Request Type = " & Request.ServerVariables("REQUEST_METHOD") & "<br>")
            .Write("Status Code = " & Response.Status & "<br>")
        End With
    End Sub


    Private Sub Class_Terminate()
        Set FObjStorage = Nothing
    End Sub

    Public Function PrintCollection(Byval Name, Byval Collection)
        Dim varItem, I
        WriteLn("<br><b>" & Name & "</b>")
        On Error Resume Next
        if IsArray(Collection) then I = LBound(Collection)
        For Each varItem in Collection
          if IsArray(Collection) then
            WriteLn("&nbsp;&nbsp;Array["& I & "]=" & varItem & "")
            I = I + 1
          elseif IsArray(Collection(varItem)) then
            Call PrintCollection("&nbsp;"& Name & "."&varItem,  Collection(varItem))
          else
            WriteLn(varItem & "=" & Collection(varItem))
          end if
          if Err.Number <> 0 then
            Err.Clear
            WriteLn(varItem & "=" & Collection(varItem))
          end if
        Next
        On Error Goto 0
        WriteLn("")
    End Function
End Class

</SCRIPT>
